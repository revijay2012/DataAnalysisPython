#!/usr/bin/env python3
"""
Woodland Play Cafe - PDF Report Generator
========================================

Generates comprehensive PDF reports for transaction and tax analysis
"""

import pandas as pd
import numpy as np
from datetime import datetime
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
import os

def load_and_analyze_data():
    """Load data and perform analysis."""
    print("📊 Loading and analyzing data...")
    
    # Load Excel files
    WoodTrans = pd.read_excel("/Users/vijayaraghavandevaraj/Downloads/Wood - TransReport.xlsx")
    WoodTax = pd.read_excel("/Users/vijayaraghavandevaraj/Downloads/Wood - Tax Report.xlsx")
    
    # Prepare transaction data
    trans_pos_df = WoodTrans[WoodTrans['Source'] == 'pos'].copy()
    trans_pos_df['Transaction Type'] = 'sale'
    trans_pos_df['Module'] = np.where(
        trans_pos_df['Order ID'].astype(str).str.startswith('MEM'),
        'memberships', 
        trans_pos_df['Source']
    )
    trans_pos_module_df = trans_pos_df[trans_pos_df['Module'] == 'pos'].copy()
    trans_pos_module_df['Transaction Date'] = pd.to_datetime(trans_pos_module_df['Transaction Date'])
    trans_pos_module_df['Month'] = trans_pos_module_df['Transaction Date'].dt.to_period('M')
    
    # Prepare tax data
    tax_pos_df = WoodTax[WoodTax['Module Name'] == 'pos'].copy()
    tax_pos_df['Date'] = pd.to_datetime(tax_pos_df['Date'])
    tax_pos_df['Month'] = tax_pos_df['Date'].dt.to_period('M')
    
    # Filter for August 2025
    target_month = "2025-08"
    trans_aug = trans_pos_module_df[trans_pos_module_df['Month'] == target_month]
    tax_aug = tax_pos_df[tax_pos_df['Month'] == target_month]
    
    # Group transaction data by Order ID
    trans_aug_grouped = trans_aug.groupby('Order ID').agg({
        'Amount': 'sum',
        'Transaction Date': 'first',
        'Location': 'first',
        'Payment Type': 'first',
        'Payment Gateway': 'first'
    }).reset_index()
    
    # Find missing records and amount differences
    missing_in_tax = trans_aug_grouped[~trans_aug_grouped['Order ID'].isin(tax_aug['Order ID'])]
    common_orders = trans_aug_grouped[trans_aug_grouped['Order ID'].isin(tax_aug['Order ID'])].copy()
    common_tax_orders = tax_aug[tax_aug['Order ID'].isin(trans_aug_grouped['Order ID'])].copy()
    
    merged_comparison = pd.merge(
        common_orders[['Order ID', 'Amount', 'Transaction Date', 'Location']], 
        common_tax_orders[['Order ID', 'Total Sum', 'Tip', 'Tax']], 
        on='Order ID', 
        how='inner'
    )
    
    merged_comparison['Amount_Diff'] = merged_comparison['Amount'] - merged_comparison['Total Sum']
    merged_comparison['Abs_Diff'] = abs(merged_comparison['Amount_Diff'])
    significant_mismatches = merged_comparison[merged_comparison['Abs_Diff'] > 0.01]
    
    return {
        'trans_aug_grouped': trans_aug_grouped,
        'tax_aug': tax_aug,
        'missing_in_tax': missing_in_tax,
        'merged_comparison': merged_comparison,
        'significant_mismatches': significant_mismatches
    }

def create_title_page(doc, styles):
    """Create title page for the report."""
    story = []
    
    # Title
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue
    )
    
    story.append(Paragraph("🏞️ WOODLAND PLAY CAFE", title_style))
    story.append(Paragraph("Data Analysis Report", title_style))
    story.append(Spacer(1, 0.5*inch))
    
    # Subtitle
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Normal'],
        fontSize=16,
        alignment=TA_CENTER,
        textColor=colors.grey
    )
    
    story.append(Paragraph("Transaction vs Tax Report Analysis", subtitle_style))
    story.append(Paragraph("August 2025", subtitle_style))
    story.append(Spacer(1, 1*inch))
    
    # Report details
    details_style = ParagraphStyle(
        'CustomDetails',
        parent=styles['Normal'],
        fontSize=12,
        alignment=TA_CENTER,
        leftIndent=2*inch,
        rightIndent=2*inch
    )
    
    report_date = datetime.now().strftime('%B %d, %Y')
    story.append(Paragraph(f"Report Generated: {report_date}", details_style))
    story.append(Paragraph("Analysis Period: August 2025", details_style))
    story.append(Paragraph("Purpose: Identify missing records and amount mismatches", details_style))
    
    story.append(PageBreak())
    return story

def create_tally_section(data, styles):
    """Create tally summary section with the specific data provided."""
    story = []
    
    # Section header
    header_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        textColor=colors.darkblue
    )
    
    story.append(Paragraph("📊 OVERALL TALLY", header_style))
    story.append(Spacer(1, 0.2*inch))
    
    # Overall Tally Table
    tally_data = [
        ['Metric', 'Value'],
        ['Total records in trans_aug', '699'],
        ['Total records in tax_aug', '693'],
        ['Total matching records', '693'],
        ['Records missing in tax_aug', '3'],
        ['Records with amount differences', '0']
    ]
    
    tally_table = Table(tally_data, colWidths=[3*inch, 2*inch])
    tally_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    story.append(tally_table)
    story.append(Spacer(1, 0.3*inch))
    
    # Amount Tally Section
    story.append(Paragraph("💰 AMOUNT TALLY", header_style))
    story.append(Spacer(1, 0.2*inch))
    
    amount_data = [
        ['Metric', 'Amount'],
        ['Total transaction amount (August)', '$9,841.10'],
        ['Total tax report amount (August)', '$9,441.11'],
        ['Difference (Trans - Tax)', '$399.99'],
        ['Total missing transaction value', '$399.99'],
        ['Total amount discrepancy', '$-0.00']
    ]
    
    amount_table = Table(amount_data, colWidths=[3*inch, 2*inch])
    amount_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    story.append(amount_table)
    story.append(Spacer(1, 0.3*inch))
    
    # Missing Record Tally Section
    story.append(Paragraph("📋 MISSING RECORD TALLY", header_style))
    story.append(Spacer(1, 0.2*inch))
    
    missing_records_data = [
        ['#', 'Order ID', 'Amount', 'Date', 'Payment'],
        ['1', '1754594250742', '$14.54', '08/07', 'physicalCard'],
        ['2', '1755643008336', '$353.63', '08/31', 'physicalCard'],
        ['3', '1756325127801', '$31.82', '08/27', 'physicalCard']
    ]
    
    missing_table = Table(missing_records_data, colWidths=[0.5*inch, 1.5*inch, 1*inch, 0.8*inch, 1.2*inch])
    missing_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkred),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
    ]))
    
    story.append(missing_table)
    story.append(PageBreak())
    return story

def create_summary_section(data, styles):
    """Create executive summary section."""
    story = []
    
    # Section header
    header_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        textColor=colors.darkblue
    )
    
    story.append(Paragraph("📊 EXECUTIVE SUMMARY", header_style))
    story.append(Spacer(1, 0.2*inch))
    
    # Summary data
    trans_total = data['trans_aug_grouped']['Amount'].sum()
    tax_total = data['tax_aug']['Total Sum'].sum()
    difference = trans_total - tax_total
    missing_count = len(data['missing_in_tax'])
    missing_amount = data['missing_in_tax']['Amount'].sum()
    diff_count = len(data['significant_mismatches'])
    diff_amount = data['merged_comparison']['Amount_Diff'].sum()
    
    # Create summary table
    summary_data = [
        ['Metric', 'Value'],
        ['Transaction Report Total', f'${trans_total:,.2f}'],
        ['Tax Report Total', f'${tax_total:,.2f}'],
        ['Total Difference', f'${difference:,.2f}'],
        ['', ''],
        ['Missing Records', f'{missing_count} orders'],
        ['Missing Amount', f'${missing_amount:,.2f}'],
        ['Amount Differences', f'{diff_count} orders'],
        ['Difference Amount', f'${diff_amount:,.2f}'],
        ['', ''],
        ['Total Orders (Transaction)', f'{len(data["trans_aug_grouped"]):,}'],
        ['Total Orders (Tax)', f'{len(data["tax_aug"]):,}'],
        ['Matching Orders', f'{len(data["merged_comparison"]):,}']
    ]
    
    summary_table = Table(summary_data, colWidths=[3*inch, 2*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    story.append(summary_table)
    story.append(PageBreak())
    return story

def create_missing_records_section(data, styles):
    """Create missing records detailed section."""
    story = []
    
    header_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        textColor=colors.darkred
    )
    
    story.append(Paragraph("🔍 MISSING RECORDS DETAILED ANALYSIS", header_style))
    story.append(Spacer(1, 0.2*inch))
    
    missing_data = data['missing_in_tax']
    
    if len(missing_data) > 0:
        # Summary
        summary_text = f"Total Missing Records: {len(missing_data)} orders worth ${missing_data['Amount'].sum():,.2f}"
        story.append(Paragraph(summary_text, styles['Normal']))
        story.append(Spacer(1, 0.3*inch))
        
        # COMPLETE DETAILED RECORDS - Each missing order with full details
        story.append(Paragraph("📋 COMPLETE MISSING RECORDS LIST:", styles['Heading2']))
        story.append(Spacer(1, 0.2*inch))
        
        missing_details = missing_data[['Order ID', 'Transaction Date', 'Amount', 'Location', 'Payment Type', 'Payment Gateway']].copy()
        missing_details = missing_details.sort_values('Transaction Date')
        
        for i, (_, row) in enumerate(missing_details.iterrows(), 1):
            # Individual record details
            record_text = f"""
            <b>Record {i}:</b><br/>
            • Order ID: {row['Order ID']}<br/>
            • Date: {row['Transaction Date'].strftime('%Y-%m-%d')}<br/>
            • Amount: ${row['Amount']:,.2f}<br/>
            • Location: {row['Location']}<br/>
            • Payment Type: {row['Payment Type']}<br/>
            • Payment Gateway: {row['Payment Gateway']}
            """
            story.append(Paragraph(record_text, styles['Normal']))
            story.append(Spacer(1, 0.1*inch))
        
        story.append(Spacer(1, 0.3*inch))
        
        # Detailed table for reference
        story.append(Paragraph("📊 MISSING RECORDS SUMMARY TABLE:", styles['Heading2']))
        story.append(Spacer(1, 0.2*inch))
        
        # Create table data
        table_data = [['#', 'Order ID', 'Date', 'Amount', 'Location', 'Payment Type', 'Gateway']]
        for i, (_, row) in enumerate(missing_details.iterrows(), 1):
            table_data.append([
                str(i),
                str(row['Order ID']),
                row['Transaction Date'].strftime('%m/%d'),
                f'${row["Amount"]:,.2f}',
                row['Location'],
                row['Payment Type'],
                row['Payment Gateway']
            ])
        
        missing_table = Table(table_data, colWidths=[0.5*inch, 1.2*inch, 0.8*inch, 1*inch, 1.5*inch, 1*inch, 1*inch])
        missing_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkred),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        
        story.append(missing_table)
        
        # Analysis by payment type
        story.append(Spacer(1, 0.3*inch))
        story.append(Paragraph("📈 ANALYSIS BY PAYMENT TYPE:", styles['Heading2']))
        
        payment_analysis = missing_data.groupby('Payment Type').agg({
            'Order ID': 'count',
            'Amount': 'sum'
        }).reset_index()
        payment_analysis.columns = ['Payment Type', 'Count', 'Total Amount']
        
        payment_data = [['Payment Type', 'Count', 'Total Amount']]
        for _, row in payment_analysis.iterrows():
            payment_data.append([
                row['Payment Type'],
                str(row['Count']),
                f'${row["Total Amount"]:,.2f}'
            ])
        
        payment_table = Table(payment_data, colWidths=[2*inch, 1*inch, 1.5*inch])
        payment_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(payment_table)
    else:
        story.append(Paragraph("✅ No missing records found!", styles['Normal']))
    
    story.append(PageBreak())
    return story

def create_amount_differences_section(data, styles):
    """Create amount differences detailed section."""
    story = []
    
    header_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        textColor=colors.darkgreen
    )
    
    story.append(Paragraph("💰 AMOUNT DIFFERENCES DETAILED ANALYSIS", header_style))
    story.append(Spacer(1, 0.2*inch))
    
    mismatches = data['significant_mismatches']
    comparison = data['merged_comparison']
    
    # Summary statistics
    summary_text = f"""
    Total Orders with Amount Differences: {len(mismatches)}
    Total Amount Difference: ${comparison['Amount_Diff'].sum():,.2f}
    Average Difference: ${comparison['Amount_Diff'].mean():.2f}
    Largest Positive Difference: ${comparison['Amount_Diff'].max():,.2f}
    Largest Negative Difference: ${comparison['Amount_Diff'].min():,.2f}
    """
    story.append(Paragraph(summary_text, styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    
    if len(mismatches) > 0:
        # COMPLETE DETAILED RECORDS - Each amount difference order with full details
        story.append(Paragraph("📋 COMPLETE AMOUNT DIFFERENCES LIST:", styles['Heading2']))
        story.append(Spacer(1, 0.2*inch))
        
        mismatch_details = mismatches[['Order ID', 'Amount', 'Total Sum', 'Amount_Diff', 'Tip', 'Tax']].copy()
        mismatch_details['Abs_Diff'] = abs(mismatch_details['Amount_Diff'])
        mismatch_details = mismatch_details.sort_values('Abs_Diff', ascending=False)
        
        for i, (_, row) in enumerate(mismatch_details.iterrows(), 1):
            # Individual record details
            record_text = f"""
            <b>Record {i}:</b><br/>
            • Order ID: {row['Order ID']}<br/>
            • Transaction Amount: ${row['Amount']:,.2f}<br/>
            • Tax Report Amount: ${row['Total Sum']:,.2f}<br/>
            • Difference: ${row['Amount_Diff']:+,.2f}<br/>
            • Tip Amount: ${row['Tip']:,.2f}<br/>
            • Tax Amount: ${row['Tax']:,.2f}
            """
            story.append(Paragraph(record_text, styles['Normal']))
            story.append(Spacer(1, 0.1*inch))
        
        story.append(Spacer(1, 0.3*inch))
        
        # Detailed table for reference
        story.append(Paragraph("📊 AMOUNT DIFFERENCES SUMMARY TABLE:", styles['Heading2']))
        story.append(Spacer(1, 0.2*inch))
        
        # Create table data
        table_data = [['#', 'Order ID', 'Trans Amount', 'Tax Amount', 'Difference', 'Tip', 'Tax']]
        for i, (_, row) in enumerate(mismatch_details.iterrows(), 1):
            table_data.append([
                str(i),
                str(row['Order ID']),
                f'${row["Amount"]:,.2f}',
                f'${row["Total Sum"]:,.2f}',
                f'${row["Amount_Diff"]:+,.2f}',
                f'${row["Tip"]:,.2f}',
                f'${row["Tax"]:,.2f}'
            ])
        
        diff_table = Table(table_data, colWidths=[0.5*inch, 1.2*inch, 1*inch, 1*inch, 1*inch, 0.8*inch, 0.8*inch])
        diff_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        
        story.append(diff_table)
        
        # Analysis by difference magnitude
        story.append(Spacer(1, 0.3*inch))
        story.append(Paragraph("📈 DIFFERENCE MAGNITUDE ANALYSIS:", styles['Heading2']))
        
        # Categorize differences
        positive_diffs = mismatch_details[mismatch_details['Amount_Diff'] > 0]
        negative_diffs = mismatch_details[mismatch_details['Amount_Diff'] < 0]
        
        analysis_text = f"""
        • Positive Differences: {len(positive_diffs)} orders (Tax Report lower than Transaction)
        • Negative Differences: {len(negative_diffs)} orders (Tax Report higher than Transaction)
        • Total Positive Impact: ${positive_diffs['Amount_Diff'].sum():,.2f}
        • Total Negative Impact: ${negative_diffs['Amount_Diff'].sum():,.2f}
        """
        story.append(Paragraph(analysis_text, styles['Normal']))
    else:
        story.append(Paragraph("✅ No significant amount differences found!", styles['Normal']))
    
    story.append(PageBreak())
    return story

def create_all_transactions_section(data, styles):
    """Create section with ALL transaction records."""
    story = []
    
    header_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        textColor=colors.navy
    )
    
    story.append(Paragraph("📋 ALL TRANSACTION RECORDS (August 2025)", header_style))
    story.append(Spacer(1, 0.2*inch))
    
    trans_data = data['trans_aug_grouped'].copy()
    trans_data = trans_data.sort_values('Transaction Date')
    
    summary_text = f"Total Transaction Records: {len(trans_data)} orders worth ${trans_data['Amount'].sum():,.2f}"
    story.append(Paragraph(summary_text, styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    
    # Create table with all transaction records
    table_data = [['#', 'Order ID', 'Date', 'Amount', 'Location', 'Payment Type', 'Gateway']]
    for i, (_, row) in enumerate(trans_data.iterrows(), 1):
        table_data.append([
            str(i),
            str(row['Order ID']),
            row['Transaction Date'].strftime('%m/%d'),
            f'${row["Amount"]:,.2f}',
            row['Location'],
            row['Payment Type'],
            row['Payment Gateway']
        ])
    
    # Split into multiple pages if too many records
    records_per_page = 30
    total_records = len(trans_data)
    
    for page_start in range(0, total_records, records_per_page):
        page_end = min(page_start + records_per_page, total_records)
        page_data = [table_data[0]] + table_data[page_start + 1:page_end + 1]
        
        if page_start > 0:
            story.append(Paragraph(f"📋 ALL TRANSACTION RECORDS (Continued - Records {page_start + 1} to {page_end})", styles['Heading2']))
            story.append(Spacer(1, 0.2*inch))
        
        trans_table = Table(page_data, colWidths=[0.5*inch, 1.2*inch, 0.8*inch, 1*inch, 1.5*inch, 1*inch, 1*inch])
        trans_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.navy),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        
        story.append(trans_table)
        
        if page_end < total_records:
            story.append(PageBreak())
    
    story.append(PageBreak())
    return story

def create_all_tax_records_section(data, styles):
    """Create section with ALL tax records."""
    story = []
    
    header_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        textColor=colors.purple
    )
    
    story.append(Paragraph("📋 ALL TAX RECORDS (August 2025)", header_style))
    story.append(Spacer(1, 0.2*inch))
    
    tax_data = data['tax_aug'].copy()
    tax_data = tax_data.sort_values('Date')
    
    summary_text = f"Total Tax Records: {len(tax_data)} orders worth ${tax_data['Total Sum'].sum():,.2f}"
    story.append(Paragraph(summary_text, styles['Normal']))
    story.append(Spacer(1, 0.3*inch))
    
    # Create table with all tax records
    table_data = [['#', 'Order ID', 'Date', 'Total Sum', 'Tip', 'Tax', 'Order Status', 'Payment Status']]
    for i, (_, row) in enumerate(tax_data.iterrows(), 1):
        table_data.append([
            str(i),
            str(row['Order ID']),
            row['Date'].strftime('%m/%d'),
            f'${row["Total Sum"]:,.2f}',
            f'${row["Tip"]:,.2f}',
            f'${row["Tax"]:,.2f}',
            row.get('Order Status', 'N/A'),
            row.get('Payment Status', 'N/A')
        ])
    
    # Split into multiple pages if too many records
    records_per_page = 30
    total_records = len(tax_data)
    
    for page_start in range(0, total_records, records_per_page):
        page_end = min(page_start + records_per_page, total_records)
        page_data = [table_data[0]] + table_data[page_start + 1:page_end + 1]
        
        if page_start > 0:
            story.append(Paragraph(f"📋 ALL TAX RECORDS (Continued - Records {page_start + 1} to {page_end})", styles['Heading2']))
            story.append(Spacer(1, 0.2*inch))
        
        tax_table = Table(page_data, colWidths=[0.4*inch, 1*inch, 0.6*inch, 0.8*inch, 0.6*inch, 0.6*inch, 1*inch, 1*inch])
        tax_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.purple),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('FONTSIZE', (0, 1), (-1, -1), 7),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')
        ]))
        
        story.append(tax_table)
        
        if page_end < total_records:
            story.append(PageBreak())
    
    story.append(PageBreak())
    return story

def create_recommendations_section(data, styles):
    """Create recommendations section."""
    story = []
    
    header_style = ParagraphStyle(
        'SectionHeader',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=12,
        textColor=colors.purple
    )
    
    story.append(Paragraph("📋 RECOMMENDATIONS & NEXT STEPS", header_style))
    story.append(Spacer(1, 0.2*inch))
    
    missing_count = len(data['missing_in_tax'])
    diff_count = len(data['significant_mismatches'])
    
    recommendations = []
    
    if missing_count > 0:
        recommendations.extend([
            f"🔍 Investigate {missing_count} missing Order IDs in Tax Report system",
            "💳 Check if cash transactions are properly recorded in tax system",
            "⚙️ Verify tax calculation process for missing orders",
            "📞 Contact tax system administrator to resolve missing records"
        ])
    
    if diff_count > 0:
        recommendations.extend([
            f"💰 Review {diff_count} orders with amount discrepancies",
            "🧮 Verify tip and tax calculations in Tax Report",
            "📊 Cross-check amount calculations with source systems",
            "🔄 Update tax calculation formulas if needed"
        ])
    
    if missing_count == 0 and diff_count == 0:
        recommendations.append("✅ All data is consistent - no action required!")
    
    recommendations.extend([
        "",
        "📈 MONITORING RECOMMENDATIONS:",
        "• Run this analysis monthly to catch discrepancies early",
        "• Set up automated alerts for missing records",
        "• Implement data validation checks in both systems",
        "• Create reconciliation procedures for future reporting"
    ])
    
    for rec in recommendations:
        if rec.startswith("📈"):
            story.append(Paragraph(rec, styles['Heading2']))
        elif rec == "":
            story.append(Spacer(1, 0.1*inch))
        else:
            story.append(Paragraph(f"• {rec}", styles['Normal']))
    
    return story

def generate_pdf_report():
    """Generate the complete PDF report."""
    print("📊 Generating PDF report...")
    
    # Load and analyze data
    data = load_and_analyze_data()
    
    # Create PDF document
    filename = f"Woodland_Analysis_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    filepath = os.path.join(os.getcwd(), filename)
    
    doc = SimpleDocTemplate(filepath, pagesize=A4)
    styles = getSampleStyleSheet()
    
    # Build story
    story = []
    
    # Title page
    story.extend(create_title_page(doc, styles))
    
    # Summary section
    story.extend(create_summary_section(data, styles))
    
    # All transaction records section
    story.extend(create_all_transactions_section(data, styles))
    
    # All tax records section
    story.extend(create_all_tax_records_section(data, styles))
    
    # Recommendations section
    story.extend(create_recommendations_section(data, styles))
    
    # Build PDF
    doc.build(story)
    
    print(f"✅ PDF report generated: {filename}")
    print(f"📁 Location: {filepath}")
    
    return filepath

if __name__ == "__main__":
    try:
        pdf_path = generate_pdf_report()
        print(f"\n🎉 Report generation completed successfully!")
        print(f"📄 Report saved as: {pdf_path}")
    except Exception as e:
        print(f"❌ Error generating report: {e}")
        import traceback
        traceback.print_exc()
