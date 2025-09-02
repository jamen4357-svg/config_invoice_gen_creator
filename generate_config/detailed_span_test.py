"""
Detailed test to see what HeaderSpanAnalyzer detects
"""

import sys
import os
sys.path.append(os.path.dirname(__file__))

from config_generator.header_span_analyzer import HeaderSpanAnalyzer


def detailed_test():
    """Run a detailed test to see exactly what's being detected."""
    
    # Find Excel file
    possible_files = [
        "../CT&INV&PL JF25042 FCA COMBINE.xlsx",
        "../config_data_extractor/CT&INV&PL MT2-25005E DAP.xlsx"
    ]
    
    excel_file = None
    for file_path in possible_files:
        if os.path.exists(file_path):
            excel_file = file_path
            print(f"✅ Found Excel file: {file_path}")
            break
    
    if not excel_file:
        print("❌ No Excel file found")
        return
    
    print(f"\n🔍 Testing HeaderSpanAnalyzer with: {excel_file}")
    print("="*60)
    
    try:
        # Initialize analyzer
        analyzer = HeaderSpanAnalyzer(excel_file)
        
        # Test each sheet individually
        print("\n📊 ANALYZING EACH SHEET...")
        spans = analyzer.analyze_header_spans()
        
        for sheet_name, sheet_spans in spans.items():
            print(f"\n📋 SHEET: {sheet_name}")
            print(f"   Found {len(sheet_spans)} header spans")
            
            for i, span in enumerate(sheet_spans, 1):
                print(f"   [{i}] Text: '{span['text']}'")
                print(f"       Col ID: {span['col_id']}")
                print(f"       Rowspan: {span['rowspan']}, Colspan: {span['colspan']}")
                print()
        
        # Test single sheet method
        print("\n🎯 TESTING SINGLE SHEET METHOD...")
        contract_spans = analyzer.get_spans_for_sheet("Contract")
        print(f"Contract sheet has {len(contract_spans)} spans:")
        for span in contract_spans:
            print(f"  - '{span['text']}' spans {span['colspan']} columns")
        
        # Summary
        total_spans = sum(len(spans) for spans in spans.values())
        print(f"\n📈 SUMMARY:")
        print(f"   Total sheets analyzed: {len(spans)}")
        print(f"   Total header spans found: {total_spans}")
        
        # Show what types of spans we found
        all_colspans = []
        all_rowspans = []
        for sheet_spans in spans.values():
            for span in sheet_spans:
                all_colspans.append(span['colspan'])
                all_rowspans.append(span['rowspan'])
        
        print(f"   Colspan range: {min(all_colspans)} to {max(all_colspans)}")
        print(f"   Rowspan range: {min(all_rowspans)} to {max(all_rowspans)}")
        
    except Exception as e:
        print(f"❌ Test failed: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    detailed_test()
