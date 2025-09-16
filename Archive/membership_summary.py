#!/usr/bin/env python3
import pandas as pd
import numpy as np

# Load data
trans_file = '/Users/vijayaraghavandevaraj/Library/Mobile Documents/com~apple~CloudDocs/Common/WoodLandPlayCafeAnalysis/MemebershipData/Transv1.xlsx'
tax_file = '/Users/vijayaraghavandevaraj/Library/Mobile Documents/com~apple~CloudDocs/Common/WoodLandPlayCafeAnalysis/MemebershipData/Tax.xlsx'
membership_file = '/Users/vijayaraghavandevaraj/Library/Mobile Documents/com~apple~CloudDocs/Common/WoodLandPlayCafeAnalysis/MemebershipData/Memebership.xlsx'

WoodTrans = pd.read_excel(trans_file)
WoodTax = pd.read_excel(tax_file)
SourceMembership = pd.read_excel(membership_file)

# Filter membership transactions
WoodTrans['Module'] = np.where(
    WoodTrans['Order ID'].astype(str).str.startswith('MEM'), 'memberships', 
    WoodTrans['Source']
)
WoodTrans_memberships = WoodTrans[WoodTrans['Module'] == 'memberships'].copy()

# Convert dates and filter for August 2025
WoodTrans_memberships['Transaction Date'] = pd.to_datetime(WoodTrans_memberships['Transaction Date'], errors='coerce')
WoodTrans_memberships['Month'] = WoodTrans_memberships['Transaction Date'].dt.to_period('M')
august_2025 = WoodTrans_memberships[WoodTrans_memberships['Month'] == '2025-08'].copy()

# Merge with membership data
Membership_df = SourceMembership[SourceMembership['membershipid'].notna()].copy()
august_2025 = august_2025.merge(
    Membership_df[['membershipid', 'subscriptionstartdate']],
    left_on='Order ID',
    right_on='membershipid',
    how='left'
)

# Create membership type
august_2025['Transaction Date'] = pd.to_datetime(august_2025['Transaction Date'], errors='coerce')
august_2025['subscriptionstartdate'] = pd.to_datetime(august_2025['subscriptionstartdate'], errors='coerce')
august_2025['Membership Type'] = np.where(
    august_2025['Transaction Date'].dt.date == august_2025['subscriptionstartdate'].dt.date,
    'New',
    'Recurring'
)

# Filter tax data for August 2025
WoodTax['Date'] = pd.to_datetime(WoodTax['Date'], errors='coerce')
WoodTax['Month'] = WoodTax['Date'].dt.to_period('M')
tax_august_2025 = WoodTax[WoodTax['Month'] == '2025-08'].copy()

print('🏪 WOODLAND PLAY CAFE - AUGUST 2025 MEMBERSHIP SUMMARY')
print('='*60)

# Match status
august_2025['Order ID'] = august_2025['Order ID'].astype(str)
tax_august_2025['Order ID'] = tax_august_2025['Order ID'].astype(str)

august_orders = set(august_2025['Order ID'])
tax_orders = set(tax_august_2025['Order ID'])

matched = august_orders.intersection(tax_orders)
not_matched = august_orders - tax_orders

print('📊 MATCH STATUS:')
print(f'   • Matched: {len(matched)}')
print(f'   • Not Matched: {len(not_matched)}')

# Membership types
membership_types = august_2025['Membership Type'].value_counts()
new_count = membership_types.get('New', 0)
recurring_count = membership_types.get('Recurring', 0)

print()
print('🎯 MEMBERSHIP TYPES:')
print(f'   • New: {new_count}')
print(f'   • Recurring: {recurring_count}')

print()
print('📋 DETAILED RECORDS:')
print()
print('🔵 MATCHED RECORDS (9):')
matched_records = august_2025[august_2025['Order ID'].isin(matched)]
for _, row in matched_records.iterrows():
    amount = row['Amount']
    print(f'   • Order ID: {row["Order ID"]} | Amount: ${amount:,.2f} | Type: {row["Membership Type"]} | Source: {row["Source"]}')

print()
print('🔴 NOT MATCHED RECORDS (22):')
not_matched_records = august_2025[august_2025['Order ID'].isin(not_matched)]
for _, row in not_matched_records.iterrows():
    amount = row['Amount']
    print(f'   • Order ID: {row["Order ID"]} | Amount: ${amount:,.2f} | Type: {row["Membership Type"]} | Source: {row["Source"]}')

print()
print('🆕 NEW MEMBERSHIPS (20):')
new_records = august_2025[august_2025['Membership Type'] == 'New']
for _, row in new_records.iterrows():
    amount = row['Amount']
    print(f'   • Order ID: {row["Order ID"]} | Amount: ${amount:,.2f} | Source: {row["Source"]}')

print()
print('🔄 RECURRING MEMBERSHIPS (12):')
recurring_records = august_2025[august_2025['Membership Type'] == 'Recurring']
for _, row in recurring_records.iterrows():
    amount = row['Amount']
    print(f'   • Order ID: {row["Order ID"]} | Amount: ${amount:,.2f} | Source: {row["Source"]}')
