#!/usr/bin/env python3
"""
Woodland Play Cafe - Clean Data Analysis Script
===============================================

This script analyzes transaction and tax data to find:
1. Missing records in tax report
2. Amount mismatches between transaction and tax reports
3. Comprehensive summary and recommendations

Author: Data Analysis
Date: 2025
"""

import pandas as pd
import numpy as np
from datetime import datetime

def load_data():
    """Load and prepare the Excel data files."""
    print("📊 Loading data files...")
    
    # Load Excel files
    WoodTrans = pd.read_excel("/Users/vijayaraghavandevaraj/Downloads/Wood - TransReport.xlsx")
    WoodTax = pd.read_excel("/Users/vijayaraghavandevaraj/Downloads/Wood - Tax Report.xlsx")
    
    print(f"   • Transaction Report: {len(WoodTrans):,} records")
    print(f"   • Tax Report: {len(WoodTax):,} records")
    
    return WoodTrans, WoodTax

def prepare_transaction_data(WoodTrans):
    """Prepare transaction data by filtering and grouping."""
    print("\n🔧 Preparing transaction data...")
    
    # Step 1: Filter for 'pos' source and add transaction type
    trans_pos_df = WoodTrans[WoodTrans['Source'] == 'pos'].copy()
    trans_pos_df['Transaction Type'] = 'sale'
    
    # Step 2: Add module classification
    trans_pos_df['Module'] = np.where(
        trans_pos_df['Order ID'].astype(str).str.startswith('MEM'),
        'memberships', 
        trans_pos_df['Source']
    )
    
    # Step 3: Filter for pos module only
    trans_pos_module_df = trans_pos_df[trans_pos_df['Module'] == 'pos'].copy()
    
    # Step 4: Prepare date columns
    trans_pos_module_df['Transaction Date'] = pd.to_datetime(trans_pos_module_df['Transaction Date'])
    trans_pos_module_df['Month'] = trans_pos_module_df['Transaction Date'].dt.to_period('M')
    
    print(f"   • POS transactions: {len(trans_pos_module_df):,} records")
    
    return trans_pos_module_df

def prepare_tax_data(WoodTax):
    """Prepare tax data by filtering and date processing."""
    print("\n🔧 Preparing tax data...")
    
    # Filter for 'pos' module and prepare dates
    tax_pos_df = WoodTax[WoodTax['Module Name'] == 'pos'].copy()
    tax_pos_df['Date'] = pd.to_datetime(tax_pos_df['Date'])
    tax_pos_df['Month'] = tax_pos_df['Date'].dt.to_period('M')
    
    print(f"   • POS tax records: {len(tax_pos_df):,} records")
    
    return tax_pos_df

def filter_august_data(trans_pos_module_df, tax_pos_df):
    """Filter data for August 2025."""
    print("\n📅 Filtering for August 2025 data...")
    
    target_month = "2025-08"
    
    trans_aug = trans_pos_module_df[trans_pos_module_df['Month'] == target_month]
    tax_aug = tax_pos_df[tax_pos_df['Month'] == target_month]
    
    print(f"   • August transactions: {len(trans_aug):,} records")
    print(f"   • August tax records: {len(tax_aug):,} records")
    
    return trans_aug, tax_aug

def group_transaction_data(trans_aug):
    """Group transaction data by Order ID to match tax report structure."""
    print("\n🔄 Grouping transaction data by Order ID...")
    
    trans_aug_grouped = trans_aug.groupby('Order ID').agg({
        'Amount': 'sum',
        'Transaction Date': 'first',
        'Location': 'first',
        'Payment Type': 'first',
        'Payment Gateway': 'first'
    }).reset_index()
    
    print(f"   • Original transaction records: {len(trans_aug):,}")
    print(f"   • Unique Order IDs (grouped): {len(trans_aug_grouped):,}")
    
    return trans_aug_grouped

def find_missing_records(trans_aug_grouped, tax_aug):
    """Find records missing in tax report."""
    print("\n🔍 Finding missing records in tax report...")
    
    missing_in_tax = trans_aug_grouped[~trans_aug_grouped['Order ID'].isin(tax_aug['Order ID'])]
    
    if len(missing_in_tax) > 0:
        print(f"   • Missing records: {len(missing_in_tax)}")
        print(f"   • Missing amount: ${missing_in_tax['Amount'].sum():.2f}")
        
        # Detailed breakdown
        print(f"\n   📋 Missing Order IDs:")
        missing_details = missing_in_tax[['Order ID', 'Transaction Date', 'Amount', 'Location', 'Payment Type']].copy()
        missing_details = missing_details.sort_values('Transaction Date')
        
        for i, (_, row) in enumerate(missing_details.iterrows(), 1):
            print(f"      {i:2d}. {row['Order ID']} | ${row['Amount']:6.2f} | {row['Transaction Date'].strftime('%m/%d')} | {row['Payment Type']}")
    else:
        print("   ✅ No missing records found!")
    
    return missing_in_tax

def find_amount_mismatches(trans_aug_grouped, tax_aug):
    """Find amount mismatches between transaction and tax reports."""
    print("\n💰 Finding amount mismatches...")
    
    # Find common orders
    common_orders = trans_aug_grouped[trans_aug_grouped['Order ID'].isin(tax_aug['Order ID'])].copy()
    common_tax_orders = tax_aug[tax_aug['Order ID'].isin(trans_aug_grouped['Order ID'])].copy()
    
    # Merge for comparison
    merged_comparison = pd.merge(
        common_orders[['Order ID', 'Amount', 'Transaction Date', 'Location']], 
        common_tax_orders[['Order ID', 'Total Sum', 'Tip', 'Tax']], 
        on='Order ID', 
        how='inner'
    )
    
    # Calculate differences
    merged_comparison['Amount_Diff'] = merged_comparison['Amount'] - merged_comparison['Total Sum']
    merged_comparison['Abs_Diff'] = abs(merged_comparison['Amount_Diff'])
    
    # Find significant mismatches
    significant_mismatches = merged_comparison[merged_comparison['Abs_Diff'] > 0.01]
    
    print(f"   • Total matching orders: {len(merged_comparison):,}")
    print(f"   • Orders with amount differences: {len(significant_mismatches)}")
    
    if len(significant_mismatches) > 0:
        print(f"   • Total discrepancy: ${merged_comparison['Amount_Diff'].sum():.2f}")
        
        print(f"\n   🔍 Amount Differences:")
        mismatch_details = significant_mismatches[['Order ID', 'Amount', 'Total Sum', 'Amount_Diff', 'Abs_Diff']].copy()
        mismatch_details = mismatch_details.sort_values('Abs_Diff', ascending=False)
        
        for i, (_, row) in enumerate(mismatch_details.iterrows(), 1):
            print(f"      {i:2d}. {row['Order ID']} | Trans: ${row['Amount']:6.2f} | Tax: ${row['Total Sum']:6.2f} | Diff: ${row['Amount_Diff']:+6.2f}")
    else:
        print("   ✅ No significant amount mismatches found!")
    
    return merged_comparison, significant_mismatches

def generate_summary_report(trans_aug_grouped, tax_aug, missing_in_tax, merged_comparison, significant_mismatches):
    """Generate comprehensive summary report."""
    print("\n" + "="*80)
    print("📊 COMPREHENSIVE ANALYSIS SUMMARY")
    print("="*80)
    
    print(f"\n📈 DATASET OVERVIEW:")
    print(f"   • Transaction Report (August): {len(trans_aug_grouped):,} unique orders")
    print(f"   • Tax Report (August): {len(tax_aug):,} records")
    print(f"   • Matching orders: {len(merged_comparison):,}")
    
    print(f"\n🔍 MISSING DATA ANALYSIS:")
    print(f"   • Missing in Tax Report: {len(missing_in_tax)} orders")
    print(f"   • Missing transaction value: ${missing_in_tax['Amount'].sum():.2f}")
    
    print(f"\n💰 AMOUNT MISMATCH ANALYSIS:")
    print(f"   • Orders with differences: {len(significant_mismatches)}")
    print(f"   • Total amount discrepancy: ${merged_comparison['Amount_Diff'].sum():.2f}")
    
    print(f"\n🎯 KEY FINDINGS:")
    if len(missing_in_tax) > 0:
        print("   1. ❌ Some transactions are missing from Tax Report")
        print("   2. 🔍 Need to investigate missing Order IDs")
    else:
        print("   1. ✅ All transactions present in Tax Report")
    
    if len(significant_mismatches) > 0:
        print("   3. ❌ Some amount mismatches found")
        print("   4. 🔍 Need to verify tax calculations")
    else:
        print("   3. ✅ All amounts match correctly")
    
    print(f"\n📋 RECOMMENDED ACTIONS:")
    if len(missing_in_tax) > 0:
        print("   1. Investigate missing Order IDs in Tax Report system")
        print("   2. Check if cash transactions are properly recorded")
        print("   3. Verify tax calculation process for missing orders")
    
    if len(significant_mismatches) > 0:
        print("   4. Review amount discrepancies in Tax Report")
        print("   5. Verify tip and tax calculations")
    
    if len(missing_in_tax) == 0 and len(significant_mismatches) == 0:
        print("   ✅ All data is consistent - no action required!")
    
    print("\n" + "="*80)

def main():
    """Main analysis function."""
    print("🏞️ WOODLAND PLAY CAFE - DATA ANALYSIS")
    print("=" * 50)
    print(f"Analysis started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    try:
        # Load and prepare data
        WoodTrans, WoodTax = load_data()
        trans_pos_module_df = prepare_transaction_data(WoodTrans)
        tax_pos_df = prepare_tax_data(WoodTax)
        
        # Filter for August 2025
        trans_aug, tax_aug = filter_august_data(trans_pos_module_df, tax_pos_df)
        
        # Group transaction data
        trans_aug_grouped = group_transaction_data(trans_aug)
        
        # Perform analysis
        missing_in_tax = find_missing_records(trans_aug_grouped, tax_aug)
        merged_comparison, significant_mismatches = find_amount_mismatches(trans_aug_grouped, tax_aug)
        
        # Generate summary
        generate_summary_report(trans_aug_grouped, tax_aug, missing_in_tax, merged_comparison, significant_mismatches)
        
        print(f"\n✅ Analysis completed successfully!")
        print(f"Completed: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
    except Exception as e:
        print(f"\n❌ Error during analysis: {e}")
        raise

if __name__ == "__main__":
    main()
