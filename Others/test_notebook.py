#!/usr/bin/env python3
"""
Test script to run the clean notebook code and identify any errors
"""

import pandas as pd
import numpy as np
from datetime import datetime

def test_notebook():
    print("🧪 Testing clean notebook code...")
    
    try:
        # Cell 1: Import libraries
        print("\n1. Testing imports...")
        import pandas as pd
        import numpy as np
        from datetime import datetime
        print("✅ Libraries imported successfully!")

        # Cell 2: Load data files
        print("\n2. Testing data loading...")
        WoodTrans = pd.read_excel("/Users/vijayaraghavandevaraj/Downloads/Wood - TransReport.xlsx")
        WoodTax = pd.read_excel("/Users/vijayaraghavandevaraj/Downloads/Wood - Tax Report.xlsx")
        print(f"✅ Transaction Report: {len(WoodTrans):,} records")
        print(f"✅ Tax Report: {len(WoodTax):,} records")

        # Cell 3: Prepare and filter data
        print("\n3. Testing data preparation...")
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

        tax_pos_df = WoodTax[WoodTax['Module Name'] == 'pos'].copy()
        tax_pos_df['Date'] = pd.to_datetime(tax_pos_df['Date'])
        tax_pos_df['Month'] = tax_pos_df['Date'].dt.to_period('M')

        target_month = "2025-08"
        trans_aug = trans_pos_module_df[trans_pos_module_df['Month'] == target_month]
        tax_aug = tax_pos_df[tax_pos_df['Month'] == target_month]
        
        print(f"✅ August transactions: {len(trans_aug):,} records")
        print(f"✅ August tax records: {len(tax_aug):,} records")

        # Cell 4: Group transaction data
        print("\n4. Testing grouping...")
        trans_aug_grouped = trans_aug.groupby('Order ID').agg({
            'Amount': 'sum',
            'Transaction Date': 'first',
            'Location': 'first',
            'Payment Type': 'first',
            'Payment Gateway': 'first'
        }).reset_index()
        
        print(f"✅ Original records: {len(trans_aug):,}")
        print(f"✅ Grouped records: {len(trans_aug_grouped):,}")

        # Cell 5: Find missing records
        print("\n5. Testing missing records analysis...")
        missing_in_tax = trans_aug_grouped[~trans_aug_grouped['Order ID'].isin(tax_aug['Order ID'])]
        print(f"✅ Missing records: {len(missing_in_tax)}")

        # Cell 6: Find amount mismatches
        print("\n6. Testing amount mismatch analysis...")
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
        
        print(f"✅ Matching orders: {len(merged_comparison):,}")
        print(f"✅ Significant mismatches: {len(significant_mismatches)}")

        # Cell 7: Generate summary
        print("\n7. Testing summary generation...")
        print("="*50)
        print("📊 ANALYSIS SUMMARY")
        print("="*50)
        print(f"Transaction Report: {len(trans_aug_grouped):,} unique orders")
        print(f"Tax Report: {len(tax_aug):,} records")
        print(f"Missing records: {len(missing_in_tax)}")
        print(f"Amount mismatches: {len(significant_mismatches)}")
        print(f"Missing value: ${missing_in_tax['Amount'].sum():.2f}")
        print(f"Total discrepancy: ${merged_comparison['Amount_Diff'].sum():.2f}")
        
        print("\n✅ All tests passed successfully!")
        return True
        
    except Exception as e:
        print(f"\n❌ Error occurred: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    success = test_notebook()
    if success:
        print("\n🎉 Notebook code is working correctly!")
    else:
        print("\n💥 There's an error in the notebook code!")
