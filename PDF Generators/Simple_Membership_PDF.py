#!/usr/bin/env python3
"""
Simple Membership Summary PDF Generator
Generates a concise PDF with match status, membership types, and detailed records
"""

import pandas as pd
import numpy as np
from datetime import datetime
import os
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch

def load_and_process_data():
    """Load and process membership data."""
    print("📁 Loading membership data...")
    
    # File paths
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

    # Match status
    august_2025['Order ID'] = august_2025['Order ID'].astype(str)
    tax_august_2025['Order ID'] = tax_august_2025['Order ID'].astype(str)

    august_orders = set(august_2025['Order ID'])
    tax_orders = set(tax_august_2025['Order ID'])

    matched = august_orders.intersection(tax_orders)
    not_matched = august_orders - tax_orders

    # Membership types
    membership_types = august_2025['Membership Type'].value_counts()
    new_count = membership_types.get('New', 0)
    recurring_count = membership_types.get('Recurring', 0)

    return august_2025, matched, not_matched, new_count, recurring_count

def create_match_status_table(august_2025, matched, not_matched):
    """Create match status table with amounts."""
    matched_records = august_2025[august_2025['Order ID'].isin(matched)]
    not_matched_records = august_2025[august_2025['Order ID'].isin(not_matched)]
    
    matched_amount = matched_records['Amount'].sum()
    not_matched_amount = not_matched_records['Amount'].sum()
    
    return [
        ['Match Status', 'Count', 'Total Amount'],
        ['Matched', str(len(matched)), f'${matched_amount:,.2f}'],
        ['Not Matched', str(len(not_matched)), f'${not_matched_amount:,.2f}']
    ]

def create_membership_types_table(august_2025, new_count, recurring_count):
    """Create membership types table with amounts."""
    new_records = august_2025[august_2025['Membership Type'] == 'New']
    recurring_records = august_2025[august_2025['Membership Type'] == 'Recurring']
    
    new_amount = new_records['Amount'].sum()
    recurring_amount = recurring_records['Amount'].sum()
    
    return [
        ['Membership Type', 'Count', 'Total Amount'],
        ['New', str(new_count), f'${new_amount:,.2f}'],
        ['Recurring', str(recurring_count), f'${recurring_amount:,.2f}']
    ]

def create_detailed_records_table(august_2025, matched, not_matched):
    """Create detailed records table."""
    matched_records = august_2025[august_2025['Order ID'].isin(matched)]
    not_matched_records = august_2025[august_2025['Order ID'].isin(not_matched)]
    
    table_data = [['Order ID', 'Amount', 'Type', 'Source', 'Status']]
    
    # Add matched records
    for _, row in matched_records.iterrows():
        table_data.append([
            row['Order ID'],
            f"${row['Amount']:,.2f}",
            row['Membership Type'],
            row['Source'],
            'MATCHED'
        ])
    
    # Add not matched records
    for _, row in not_matched_records.iterrows():
        table_data.append([
            row['Order ID'],
            f"${row['Amount']:,.2f}",
            row['Membership Type'],
            row['Source'],
            'NOT MATCHED'
        ])
    
    return table_data

def generate_pdf():
    """Generate the simple PDF report."""
    # Load data
    august_2025, matched, not_matched, new_count, recurring_count = load_and_process_data()
    
    # Create filename
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Simple_Membership_Summary_{timestamp}.pdf"
    
    print(f"📄 Generating PDF: {filename}")
    
    # Create PDF document
    doc = SimpleDocTemplate(filename, pagesize=A4, topMargin=1*inch)
    story = []
    
    # Define styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'Title',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=20,
        textColor=colors.darkblue,
        alignment=1
    )
    
    heading_style = ParagraphStyle(
        'Heading',
        parent=styles['Heading2'],
        fontSize=12,
        spaceAfter=10,
        textColor=colors.darkgreen
    )
    
    # Title
    story.append(Paragraph("🏪 Woodland Play Cafe", title_style))
    story.append(Paragraph("Membership Summary - August 2025", title_style))
    story.append(Spacer(1, 0.2*inch))
    
    # Total Summary
    total_amount = august_2025['Amount'].sum()
    story.append(Paragraph(f"📊 TOTAL MEMBERSHIP REVENUE: ${total_amount:,.2f}", heading_style))
    story.append(Spacer(1, 0.2*inch))
    
    # Match Status
    story.append(Paragraph("📊 MATCH STATUS", heading_style))
    match_data = create_match_status_table(august_2025, matched, not_matched)
    match_table = Table(match_data, colWidths=[1.5*inch, 1*inch, 1.5*inch])
    match_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(match_table)
    story.append(Spacer(1, 0.3*inch))
    
    # Membership Types
    story.append(Paragraph("🎯 MEMBERSHIP TYPES", heading_style))
    types_data = create_membership_types_table(august_2025, new_count, recurring_count)
    types_table = Table(types_data, colWidths=[1.5*inch, 1*inch, 1.5*inch])
    types_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    story.append(types_table)
    story.append(Spacer(1, 0.3*inch))
    
    # Detailed Records
    story.append(Paragraph("📋 DETAILED RECORDS", heading_style))
    detailed_data = create_detailed_records_table(august_2025, matched, not_matched)
    detailed_table = Table(detailed_data, colWidths=[1.5*inch, 1*inch, 1*inch, 1*inch, 1*inch])
    detailed_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.navy),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        # Color code matched vs not matched
        ('BACKGROUND', (4, 1), (4, len(matched)), colors.lightgreen),
        ('BACKGROUND', (4, len(matched)+1), (4, -1), colors.lightcoral)
    ]))
    story.append(detailed_table)
    
    # Build PDF
    doc.build(story)
    
    print(f"✅ PDF generated: {filename}")
    print(f"📁 Location: {os.path.abspath(filename)}")
    
    return filename

if __name__ == "__main__":
    try:
        generate_pdf()
        print("🎉 PDF generation completed successfully!")
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()
