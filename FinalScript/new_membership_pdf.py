#!/usr/bin/env python3
"""
🏪 Woodland Play Cafe - Clean Membership PDF Generator
Generates membership reports from Membership.ipynb data processing
"""

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import pandas as pd
import numpy as np
import os
from datetime import datetime

def load_membership_data():
    """Load and process membership data exactly as in Membership.ipynb"""
    print("📁 Loading membership data from Membership.ipynb processing...")
    
    # Base path
    base_path = "/Users/vijayaraghavandevaraj/Library/Mobile Documents/com~apple~CloudDocs/Common/WoodLandPlayCafeAnalysis/MemebershipData"
    
    # File paths
    trans_file = f"{base_path}/Transv1.xlsx"
    tax_file = f"{base_path}/Tax.xlsx"
    membership_file = f"{base_path}/Memebership.xlsx"
    
    # Load Excel files
    WoodTrans = pd.read_excel(trans_file)
    WoodTax = pd.read_excel(tax_file)
    SourceMembership = pd.read_excel(membership_file)
    
    print(f"✅ Transaction Report loaded: {len(WoodTrans)} records")
    print(f"✅ Tax Report loaded: {len(WoodTax)} records")
    print(f"✅ Membership Report loaded: {len(SourceMembership)} records")
    
    # Clean membership data
    Membership_df = SourceMembership[SourceMembership["membershipid"].notna()].copy()
    print(f"✅ Clean Membership Records: {len(Membership_df)} records")
    
    # Create Module column based on Order ID prefix
    WoodTrans["Module"] = np.where(
        WoodTrans["Order ID"].astype(str).str.startswith("MEM"), 
        "memberships", 
        WoodTrans["Source"]
    )
    
    # Filter only membership transactions
    WoodTrans_memberships = WoodTrans[WoodTrans["Module"] == "memberships"].copy()
    print(f"✅ Membership transactions: {len(WoodTrans_memberships)} records")
    
    # Filter for August 2025
    WoodTrans_memberships["Transaction Date"] = pd.to_datetime(WoodTrans_memberships["Transaction Date"])
    
    august_2025 = WoodTrans_memberships[
        (WoodTrans_memberships["Transaction Date"].dt.year == 2025) & 
        (WoodTrans_memberships["Transaction Date"].dt.month == 8)
    ].copy()
    
    print(f"✅ August 2025 membership transactions: {len(august_2025)} records")
    
    # Add startdate from membership data
    Membership_df["membershipid"] = Membership_df["membershipid"].astype(str)
    august_2025["Order ID"] = august_2025["Order ID"].astype(str)
    
    august_2025 = august_2025.merge(
        Membership_df[["membershipid", "startdate"]],
        left_on="Order ID",
        right_on="membershipid",
        how="left"
    )
    
    # Add Membership Type
    august_2025["startdate"] = pd.to_datetime(august_2025["startdate"], errors="coerce")
    august_2025["Membership Type"] = august_2025["startdate"].apply(
        lambda x: "New" if pd.notnull(x) and x >= pd.Timestamp("2025-08-01") else "Recurring"
    )
    
    # Merge Tax data
    WoodTax["Order ID"] = WoodTax["Order ID"].astype(str)
    august_2025 = august_2025.merge(
        WoodTax[["Order ID", "Tax"]],
        on="Order ID",
        how="left",
        suffixes=("", "_WoodTax")
    )
    
    # Calculate Tax_Percentage
    august_2025["Tax_Percentage"] = august_2025.apply(
        lambda row: (row["Tax"] / row["Amount"] * 100) if pd.notnull(row["Tax"]) and row["Amount"] > 0 else 0,
        axis=1
    )
    
    # Calculate Final_Tax
    common_tax_percentage = august_2025.loc[august_2025["Tax"].notnull(), "Tax_Percentage"].mode()
    if len(common_tax_percentage) > 0:
        tax_rate = common_tax_percentage.iloc[0] / 100
    else:
        tax_rate = 0.0888  # 8.88% fallback
    
    august_2025["Final_Tax"] = august_2025.apply(
        lambda row: row["Tax"] if pd.notnull(row["Tax"]) else row["Amount"] * tax_rate,
        axis=1
    )
    
    # Add Year and Month
    august_2025["Year"] = august_2025["Transaction Date"].dt.year
    august_2025["Month"] = august_2025["Transaction Date"].dt.month
    
    return august_2025

def create_membership_summary(august_2025):
    """Create membership summary table"""
    membership_summary = (
        august_2025.groupby(["Year", "Month", "Membership Type"])
        .agg({
            "Amount": "sum",
            "Final_Tax": "sum"
        })
        .round(2)
        .rename(columns={"Amount": "Total_Amount", "Final_Tax": "Total_Tax"})
        .reset_index()
    )
    
    # Calculate effective tax percentage
    membership_summary["Effective_Tax_%"] = (
        (membership_summary["Total_Tax"] / membership_summary["Total_Amount"]) * 100
    ).round(2)
    
    # Add grand total row
    grand_totals = pd.DataFrame({
        "Year": ["All"],
        "Month": ["All"],
        "Membership Type": ["Grand Total"],
        "Total_Amount": [membership_summary["Total_Amount"].sum()],
        "Total_Tax": [membership_summary["Total_Tax"].sum()],
        "Effective_Tax_%": [
            (membership_summary["Total_Tax"].sum() / membership_summary["Total_Amount"].sum()) * 100
        ]
    })
    
    grand_totals["Effective_Tax_%"] = grand_totals["Effective_Tax_%"].round(2)
    
    # Combine
    membership_summary = pd.concat([membership_summary, grand_totals], ignore_index=True)
    
    return membership_summary

def generate_clean_membership_pdf():
    """Generate clean membership PDF report"""
    
    # Load data
    august_2025 = load_membership_data()
    
    # Create membership summary
    membership_summary = create_membership_summary(august_2025)
    
    # Output file path
    output_path = "/Users/vijayaraghavandevaraj/Library/Mobile Documents/com~apple~CloudDocs/Common/WoodLandPlayCafeAnaysisv2/DataAnalysisPython/MemebershipData"
    pdf_file = os.path.join(output_path, f"Clean_Membership_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
    
    # Create PDF document
    doc = SimpleDocTemplate(pdf_file, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []
    
    # Title
    title = Paragraph("🏪 Woodland Play Cafe - Clean Membership Report (August 2025)", styles["Title"])
    elements.append(title)
    elements.append(Spacer(1, 12))
    
    # MEMBERSHIP SUMMARY
    elements.append(Paragraph("📊 MEMBERSHIP SUMMARY (Year, Month, Amount + Tax)", styles["Heading2"]))
    
    # Convert summary to table format (removed Effective_Tax_% column)
    summary_data = [["Year", "Month", "Membership Type", "Total_Amount", "Total_Tax"]]
    for _, row in membership_summary.iterrows():
        summary_data.append([
            str(row["Year"]),
            str(row["Month"]),
            str(row["Membership Type"]),
            f"{row['Total_Amount']:.2f}",
            f"{row['Total_Tax']:.2f}"
        ])
    
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
    
    elements.append(summary_table)
    elements.append(Spacer(1, 20))
    
    # SAMPLE DATA WITH FINAL_TAX
    if len(august_2025) > 0:
        elements.append(PageBreak())
        elements.append(Paragraph("📋 FINAL_TAX", styles["Heading1"]))
        elements.append(Spacer(1, 12))
        
        # Summary info
        summary_info = Paragraph(f"<b>Total Records:</b> {len(august_2025)} membership transactions (August 2025)", styles["Normal"])
        elements.append(summary_info)
        elements.append(Spacer(1, 12))
        
        # Create detailed records table with serial numbers and transaction date
        columns = ["#", "Order ID", "Transaction Date", "Amount", "Final_Tax", "Membership Type"]
        detailed_data = [columns]
        
        for idx, (_, row) in enumerate(august_2025.iterrows(), 1):
            # Format transaction date
            trans_date = str(row["Transaction Date"].date()) if pd.notnull(row["Transaction Date"]) else "N/A"
            detailed_data.append([
                str(idx),
                str(row["Order ID"]),
                trans_date,
                f"{row['Amount']:.2f}",
                f"{row['Final_Tax']:.2f}",
                str(row["Membership Type"])
            ])
        
        # Create table with optimized column widths (added Transaction Date column)
        detail_table = Table(detailed_data, repeatRows=1, colWidths=[0.5*72, 2.0*72, 1.2*72, 1.0*72, 1.0*72, 1.5*72])
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
        
        elements.append(detail_table)
    
    # Build PDF
    doc.build(elements)
    
    print(f"✅ Clean Membership PDF generated at: {pdf_file}")
    return pdf_file

if __name__ == "__main__":
    generate_clean_membership_pdf()
