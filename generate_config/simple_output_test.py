"""
Simple test to see what HeaderSpanAnalyzer produces
"""

import sys
import os
sys.path.append(os.path.dirname(__file__))

from config_generator.header_span_analyzer import HeaderSpanAnalyzer
import json


def simple_output_test():
    """Just show the clean output from the module."""
    
    excel_file = "../config_data_extractor/CT&INV&PL MT2-25005E DAP.xlsx"
    
    if not os.path.exists(excel_file):
        print("Excel file not found")
        return
    
    # Initialize and get results
    analyzer = HeaderSpanAnalyzer(excel_file)
    result = analyzer.analyze_header_spans()
    
    # Show clean JSON output
    print("HeaderSpanAnalyzer Output:")
    print("=" * 40)
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    simple_output_test()
