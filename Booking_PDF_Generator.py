#!/usr/bin/env python3
"""
Booking Report PDF Generator
Generates a comprehensive PDF report for Woodland Play Cafe booking analysis
"""

import pandas as pd
import numpy as np
from datetime import datetime
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT

def generate_booking_report_pdf(trans_filtered, filename=None):
    """
    Generate a detailed PDF report for booking analysis
    
    Args:
        trans_filtered: Filtered DataFrame with booking data
        filename: Optional custom filename
    
    Returns:
        str: Generated PDF filename
    """
    
    # Create filename with timestamp if not provided
    if filename is None:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Booking_Report_August_2025_{timestamp}.pdf"
    
    # Create PDF document
    doc = SimpleDocTemplate(filename, pagesize=A4, 
                          rightMargin=72, leftMargin=72, 
                          topMargin=72, bottomMargin=18)
    
    # Container for the 'Flowable' objects
    elements = []
    
    # Get styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        spaceAfter=30,
        alignment=TA_CENTER,
        textColor=colors.darkblue
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        spaceAfter=12,
        textColor=colors.darkgreen
    )
    
    # Title
    title = Paragraph("Woodland Play Cafe - Booking Analysis Report", title_style)
    elements.append(title)
    
    # Report metadata
    report_date = datetime.datetime.now().strftime("%B %d, %Y")
    elements.append(Paragraph(f"<b>Report Date:</b> {report_date}", styles['Normal']))
    elements.append(Paragraph(f"<b>Analysis Period:</b> August 2025", styles['Normal']))
    elements.append(Paragraph(f"<b>Total Records Analyzed:</b> {len(trans_filtered):,}", styles['Normal']))
    elements.append(Spacer(1, 20))
    
    # Executive Summary
    elements.append(Paragraph("Executive Summary", heading_style))
    
    # Calculate key metrics
    total_transactions = len(trans_filtered)
    total_amount = trans_filtered['total'].sum() if 'total' in trans_filtered.columns else 0
    avg_transaction = total_amount / total_transactions if total_transactions > 0 else 0
    
    summary_data = [
        ['Metric', 'Value'],
        ['Total Transactions', f"{total_transactions:,}"],
        ['Total Revenue', f"${total_amount:,.2f}"],
        ['Average Transaction Value', f"${avg_transaction:,.2f}"],
        ['Analysis Period', 'August 2025']
    ]
    
    summary_table = Table(summary_data, colWidths=[2*inch, 2*inch])
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    
    elements.append(summary_table)
    elements.append(Spacer(1, 20))
    
    # Breakdown by Booking Type (ONLY this breakdown is included)
    if 'booking_flow_type' in trans_filtered.columns:
        elements.append(Paragraph("Breakdown by Booking Type", heading_style))
        
        booking_type_data = trans_filtered.groupby('booking_flow_type').agg({
            'total': ['count', 'sum']
        }).round(2)
        
        # Prepare table data
        booking_table_data = [['Booking Type', 'Count', 'Total Amount']]
        for idx, row in booking_type_data.iterrows():
            booking_table_data.append([
                str(idx),
                f"{int(row[('total', 'count')]):,}",
                f"${row[('total', 'sum')]:,.2f}"
            ])
        
        booking_table = Table(booking_table_data, colWidths=[2*inch, 1.5*inch, 1.5*inch])
        booking_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elements.append(booking_table)
        elements.append(Spacer(1, 15))
    
    # Key Insights
    elements.append(Paragraph("Key Insights", heading_style))
    
    insights = []
    if 'booking_flow_type' in trans_filtered.columns:
        most_popular_booking = trans_filtered['booking_flow_type'].value_counts().index[0]
        insights.append(f"• Most popular booking type: {most_popular_booking}")
    
    if 'booking_flow_type' in trans_filtered.columns:
        booking_counts = trans_filtered['booking_flow_type'].value_counts()
        total_bookings = len(trans_filtered)
        for booking_type, count in booking_counts.items():
            percentage = (count / total_bookings) * 100
            insights.append(f"• {booking_type}: {count} bookings ({percentage:.1f}%)")
    
    for insight in insights:
        elements.append(Paragraph(insight, styles['Normal']))
    
    # Build PDF
    doc.build(elements)
    
    print(f"✅ PDF report generated successfully: {filename}")
    print(f"📄 Report contains comprehensive analysis of {len(trans_filtered):,} transactions")
    print(f"📊 Report includes: Executive Summary, Booking Type Breakdown, and Key Insights")
    
    return filename

# Example usage (uncomment to use with your data):
"""
# Load your data
trans_file = "/Users/vijayaraghavandevaraj/Downloads/WBooking.xlsx"
WoodTrans = pd.read_excel(trans_file)

# Filter data
trans_pos_df = WoodTrans[WoodTrans['booking_source'].isin(['Website', 'crm'])].copy()
trans_pos_df['eventDate'] = pd.to_datetime(trans_pos_df['eventDate'])
trans_pos_df['Month'] = trans_pos_df['eventDate'].dt.to_period('M')

# Filter for August 2025
target_months = ["2025-08"]
trans_filtered = trans_pos_df[trans_pos_df['Month'].astype(str).isin(target_months)]

# Generate PDF
pdf_filename = generate_booking_report_pdf(trans_filtered)
"""
