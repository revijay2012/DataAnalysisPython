from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd
import numpy as np
import os

# Base path
base_path = "/Users/vijayaraghavandevaraj/Library/Mobile Documents/com~apple~CloudDocs/Common/WoodLandPlayCafeAnaysisv2/DataAnalysisPython/MemebershipData"
pdf_file = os.path.join(base_path, "Woodland_Play_Cafe_August_Membership_Tax.pdf")

# Load membership data
print("📁 Loading membership data...")
membership_file = f"{base_path}/Memebership.xlsx"
all_membership_data = pd.read_excel(membership_file)
print(f"✅ Membership data loaded: {len(all_membership_data)} records")

# Filter for August 2025 based on Transaction Date (updatedat field)
print("🔍 Filtering for August 2025 based on Transaction Date...")

# Convert updatedat to datetime for filtering
all_membership_data['updatedat'] = pd.to_datetime(all_membership_data['updatedat'], errors='coerce')

# Filter for August 2025 (2025-08)
august_2025 = all_membership_data[
    (all_membership_data['updatedat'].dt.year == 2025) & 
    (all_membership_data['updatedat'].dt.month == 8)
].copy()

print(f"✅ August 2025 records filtered: {len(august_2025)} records")

# If no records found with updatedat, try with purchasedon as fallback
if len(august_2025) == 0:
    print("⚠️ No records found with updatedat in August 2025, trying purchasedon...")
    all_membership_data['purchasedon'] = pd.to_datetime(all_membership_data['purchasedon'], errors='coerce')
    august_2025 = all_membership_data[
        (all_membership_data['purchasedon'].dt.year == 2025) & 
        (all_membership_data['purchasedon'].dt.month == 8)
    ].copy()
    print(f"✅ August 2025 records filtered by purchasedon: {len(august_2025)} records")

doc = SimpleDocTemplate(pdf_file, pagesize=A4)
styles = getSampleStyleSheet()
elements = []

# Title
title = Paragraph("Woodland Play Cafe - August 2025 Membership Tax Report (Filtered by Transaction Date)", styles["Title"])
elements.append(title)
elements.append(Spacer(1, 12))

# ---------------------------
# 📊 MEMBERSHIP SUMMARY
# ---------------------------
# Calculate actual summary from the loaded data in the exact format requested
if len(august_2025) > 0:
    # Group by membership type and calculate totals
    summary_stats = august_2025.groupby('membershiptypename').agg({
        'total': 'sum',
        'tax': 'sum'
    }).round(2)
    
    # Calculate grand totals
    grand_total = august_2025['total'].sum()
    grand_tax = august_2025['tax'].sum()
    effective_tax_rate = (grand_tax / grand_total * 100) if grand_total > 0 else 0
    
    # Build summary data in the exact format requested
    summary_data = [["Year", "Month", "Membership Type", "Total_Amount", "Total_Tax", "Effective_Tax_%"]]
    
    for membership_type, row in summary_stats.iterrows():
        tax_rate = (row['tax'] / row['total'] * 100) if row['total'] > 0 else 0
        summary_data.append([
            "2025",
            "8", 
            str(membership_type),
            f"{row['total']:.2f}",
            f"{row['tax']:.2f}",
            f"{tax_rate:.2f}"
        ])
    
    # Add grand total row
    summary_data.append([
        "All",
        "All",
        "Grand Total",
        f"{grand_total:.2f}",
        f"{grand_tax:.2f}",
        f"{effective_tax_rate:.2f}"
    ])
else:
    # Fallback data if no records
    summary_data = [
        ["Year", "Month", "Membership Type", "Total_Amount", "Total_Tax", "Effective_Tax_%"],
        ["2025", "8", "No Data", "0.00", "0.00", "0.00"]
    ]

summary_table = Table(summary_data)
summary_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
]))
elements.append(Paragraph("📊 MEMBERSHIP SUMMARY (Year, Month, Amount + Tax)", styles["Heading2"]))
elements.append(summary_table)
elements.append(Spacer(1, 20))

# ---------------------------
# 📋 AUGUST MEMBERSHIP RECORDS (All)
# ---------------------------
if len(august_2025) > 0:
    # Add page break before detailed records
    elements.append(PageBreak())
    
    # Create detailed records table in the exact format requested
    columns = ["Order ID", "Amount", "Tax", "Tax_Percentage", "Final_Tax", "Membership Type"]
    
    # Check which columns are available in the data
    available_columns = august_2025.columns.tolist()
    print(f"Available columns: {available_columns}")
    
    # Prepare data rows in the exact format requested
    detailed_data = [columns]
    
    for _, row in august_2025.iterrows():
        # Get data in the exact format requested
        order_id = row.get("membershipid", "N/A")
        amount = row.get("total", 0) if pd.notnull(row.get("total", 0)) else 0
        tax = row.get("tax", 0) if pd.notnull(row.get("tax", 0)) else 0
        
        # Calculate tax percentage
        tax_percentage = (tax / amount * 100) if amount > 0 else 0
        
        # Calculate final tax (use tax if available, otherwise calculate from amount)
        if pd.notnull(row.get("tax", 0)) and row.get("tax", 0) > 0:
            final_tax = tax
        else:
            # Calculate tax at 8.88% if not provided
            final_tax = amount * 0.0888
        
        membership_type = row.get("membershiptypename", "N/A")
        
        detailed_data.append([
            str(order_id),
            f"{amount:.2f}",
            "NaN" if tax == 0 else f"{tax:.2f}",
            f"{tax_percentage:.2f}",
            f"{final_tax:.2f}",
            str(membership_type)
        ])
    
    # Create table with optimized column widths for the requested format
    detail_table = Table(detailed_data, repeatRows=1, colWidths=[2.0*72, 1.0*72, 1.0*72, 1.2*72, 1.0*72, 1.5*72])
    detail_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lavender),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    
    elements.append(Paragraph("📋 SAMPLE DATA WITH FINAL_TAX", styles["Heading1"]))
    elements.append(Spacer(1, 12))
    
    # Add summary info
    summary_info = Paragraph(f"<b>Total Records:</b> {len(august_2025)} membership transactions (August 2025)", styles["Normal"])
    elements.append(summary_info)
    elements.append(Spacer(1, 12))
    
    elements.append(detail_table)
else:
    print("⚠️ No membership data found for August 2025")

# Build PDF
doc.build(elements)

print(f"✅ PDF generated at: {pdf_file}")
