# -*- coding: utf-8 -*-
"""
Created on Thu Aug  7 08:57:52 2025

@author: ann.calleja
"""

from flask import Flask, redirect, request, jsonify, render_template_string
import urllib.parse
import requests
import pandas as pd
import json
from datetime import datetime
import os

app = Flask(__name__)

# Xero app credentials
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
REDIRECT_URI = os.getenv("REDIRECT_URI", "http://localhost:5000/callback")
SCOPES = 'openid profile email accounting.transactions offline_access'

# Globals to store tokens & tenant_id 
tokens = {}
tenant_id = None


# Step 1: Redirect to Xero for Authorization
@app.route('/')
def login():
    auth_url = (
        "https://login.xero.com/identity/connect/authorize?"
        f"response_type=code"
        f"&client_id={CLIENT_ID}"
        f"&redirect_uri={urllib.parse.quote(REDIRECT_URI)}"
        f"&scope={urllib.parse.quote(SCOPES)}"
        f"&state=12345"
    )
    return redirect(auth_url)


# Step 2: Callback to get authorization code and exchange tokens
@app.route('/callback')
def callback():
    global tokens, tenant_id

    error = request.args.get('error')
    if error:
        return f"❌ Authorization failed: {error}"

    code = request.args.get('code')
    if not code:
        return "⚠️ No authorization code received."

    # Exchange code for tokens
    token_url = "https://identity.xero.com/connect/token"
    response = requests.post(token_url, data={
        "grant_type": "authorization_code",
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET
    })

    tokens = response.json()
    if 'access_token' not in tokens:
        return f"❌ Failed to get access token: {tokens}"

    # Get tenant_id
    headers = {"Authorization": f"Bearer {tokens['access_token']}"}
    connections = requests.get("https://api.xero.com/connections", headers=headers)
    tenants = connections.json()
    if not tenants:
        return "❌ No Xero tenant found."

    tenant_id = tenants[0]['tenantId']

    return """
    ✅ Authorization successful!<br><br>
    Now upload your Excel file at <a href="/upload">/upload</a>
    """


# Upload Excel file and build PO
@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'GET':
        return '''
        <h2>Upload your Excel file (.xls or .xlsx)</h2>
        <form method="post" enctype="multipart/form-data">
          <input type="file" name="file" accept=".xls,.xlsx" required>
          <input type="submit" value="Upload and Build PO">
        </form>
        '''

    # POST - process file
    file = request.files.get('file')
    if not file:
        return "No file uploaded", 400

    # Load Excel into DataFrame
    try:
        df = pd.read_excel(file, header=None)
    except Exception as e:
        return f"Failed to read Excel: {e}", 400

    # Extract functions
    def extract_column_values(field_name):
        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                cell_value = str(df.iat[row_idx, col_idx]).strip()
                if field_name.lower() == cell_value.lower():
                    values = []
                    current_row = row_idx + 1
                    while current_row < len(df):
                        value = df.iat[current_row, col_idx]
                        if pd.isna(value):
                            break
                        values.append(value)
                        current_row += 1
                    return values
        return None

    def extract_quote_info(section_title="QUOTE INFORMATION"):
        quote_info = {}
        for i in range(len(df)):
            for j in range(len(df.columns)):
                if str(df.iat[i, j]).strip() == section_title:
                    start_row = i + 1
                    for k in range(start_row, start_row + 10):
                        if k >= len(df): break
                        row_values = df.iloc[k].dropna().tolist()
                        for idx in range(0, len(row_values) - 1, 2):
                            key = str(row_values[idx]).strip()
                            value = str(row_values[idx + 1]).strip()
                            if key:
                                quote_info[key] = value
                    return quote_info
        return {}

    # Extract Excel data
    item_numbers = extract_column_values("Item Number") or []
    descriptions = extract_column_values("Description") or []
    quantities = extract_column_values("Qty") or []
    unit_prices = extract_column_values("Unit Price") or []
    quote_info = extract_quote_info()

    # Build line items
    line_items = []
    for i in range(len(descriptions)):
        line_items.append({
            "Description": str(descriptions[i]),
            "Quantity": float(quantities[i]) if i < len(quantities) else 1,
            "UnitAmount": float(unit_prices[i]) if i < len(unit_prices) else 0,
            "AccountCode": "400",   # Changed to valid code
            "TaxType": "INPUT"
        })

    # PO Metadata from quote info
    contact_name = quote_info.get("Reseller Contact", "Unknown Supplier")
    reference = quote_info.get("Sales Quotation", "AutoPO")
    currency_code = quote_info.get("Currency", "AUD")
    if currency_code not in ["AUD", "NZD"]:  # supported currencies
        currency_code = "AUD"
    #delivery_date = quote_info.get("Validity End Date", datetime.today().strftime('%Y-%m-%d'))
    raw_date = quote_info.get("Validity End Date", "")
    try:
        delivery_date = datetime.strptime(raw_date, "%d/%m/%Y").strftime("%Y-%m-%d")
    except Exception:
        delivery_date = datetime.today().strftime("%Y-%m-%d")

    # Get or create ContactID in Xero
    try:
        contact_id = get_or_create_contact_id(contact_name)
    except Exception as e:
        return f"❌ Error getting/creating contact: {e}", 400

    # Build PO JSON payload
    po_data = {
        "Contact": {
            "ContactID": contact_id,
            "Name": contact_name
        },
        "Date": datetime.today().strftime('%Y-%m-%d'),
        "DeliveryDate": delivery_date,
        "LineItems": line_items,
        "DeliveryAddress": "Enablis Office",
        "Reference": reference,
        "CurrencyCode": currency_code,
        "Status": "DRAFT"
    }

    # Save PO data
    global po_payload
    po_payload = po_data

    # Show PO JSON and a send button
    return render_template_string('''
    <h2>Purchase Order JSON Payload</h2>
    <pre>{{ po_json }}</pre>
    <form action="/send_po" method="post">
      <button type="submit">Send PO to Xero</button>
    </form>
    ''', po_json=json.dumps(po_data, indent=4))


def get_or_create_contact_id(name):
    # Search for contact
    url = "https://api.xero.com/api.xro/2.0/Contacts"
    headers = {
        "Authorization": f"Bearer {tokens['access_token']}",
        "Xero-tenant-id": tenant_id,
        "Accept": "application/json"
    }
    resp = requests.get(url, headers=headers, params={"where": f'Name=="{name}"'})
    data = resp.json()
    contacts = data.get("Contacts", [])
    if contacts:
        return contacts[0]["ContactID"]

    # Create contact
    create_resp = requests.post(url, headers={**headers, "Content-Type": "application/json"},
                                json={"Name": name})

    if create_resp.status_code not in [200, 201]:
        raise Exception(f"Failed to create contact: {create_resp.text}")

    created_contact = create_resp.json()["Contacts"][0]
    return created_contact["ContactID"]


# Send the PO to Xero
@app.route('/send_po', methods=['POST'])
def send_po():
    global po_payload
    if not po_payload:
        return "❌ No PO payload available. Upload an Excel file first.", 400

    url = "https://api.xero.com/api.xro/2.0/PurchaseOrders"
    headers = {
        "Authorization": f"Bearer {tokens['access_token']}",
        "Xero-tenant-id": tenant_id,
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, json=po_payload)
    if response.status_code not in [200, 201]:
        return f"❌ Failed to send PO: {response.text}", 500

    return f"✅ Purchase Order sent successfully!<br>Response:<br><pre>{response.text}</pre>" 


if __name__ == '__main__':
    app.run(port=5000, debug=True)
