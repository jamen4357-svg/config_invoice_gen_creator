# KB25019 Excel File Analysis - Table Structure and Column Spans

## Overview
Analysis of `CT&INV&PL KB25019 DAP(1)(1).xlsx` file to identify table structures, column headers, and column span configurations for automatic detection and configuration.

## File Structure
- **Filename**: CT&INV&PL KB25019 DAP(1)(1).xlsx
- **Document Type**: Contract, Invoice, and Packing List (CT&INV&PL)
- **Reference Number**: KB25019
- **Term**: DAP (Delivered at Place)

## Sheet Analysis

### 1. Contract Sheet (Range: A1:J36)

#### CT/Contract References Found:
- **G1**: "SALES CONTRACT" - Main document title
- **A5**: "Sales Contract Number" - Header field
- **A6**: "KB25019" - Contract number value

#### Table Structure:
**Main Data Table (Rows 11-26)**
- **Header Row (Row 11)**:
  - A11: "Description of Goods."
  - G11: "Unit Price"
  - H11: "Quantity(SF)"
  - I11: "Amount(USD)"

**Column Spans (Merged Cells)**:
- Description column spans A-F (6 columns wide)
- Unit Price: Single column G
- Quantity: Single column H  
- Amount: Single column I
- Total row spans A-F for "TOTAL OF:" label

#### Key Column Span Patterns:
1. **Wide Description Area**: A1:F26 (6-column span for product descriptions)
2. **Financial Columns**: G:I (3 separate single columns for price, quantity, amount)
3. **Header Spans**: Various administrative fields use 3-column spans (A:C, D:F, G:J)

### 2. Invoice Sheet (Range: A1:G45)

#### Table Structure:
**Main Data Table (Rows 20-35)**
- **Header Row (Row 20)**:
  - A20: "Mark & Nº"
  - B20: "P.O. Nº"  
  - C20: "ITEM Nº"
  - D20: "Description"
  - E20: "Quantity"
  - F20: "Unit price (USD)"
  - G20: "Amount (USD)"

**Column Spans (Merged Cells)**:
- Company name spans full width A1:G1
- Address lines span A2:G2, A3:G3, etc.
- Description column (D21:D34) merged vertically for multiple rows
- Each data column is single-width

#### Key Column Span Patterns:
1. **Header Spans**: Full-width spans (A:G) for company information
2. **Data Table**: 7 distinct single-width columns
3. **Vertical Merges**: Description field uses vertical merging for related items

### 3. Packing List Sheet (Range: A1:L200, actual data to L30)

#### Table Structure:
**Main Data Table (Rows 19-30)**
- **Header Row (Row 19)**:
  - A19: "Mark & Nº"
  - B19: "P.O Nº"
  - C19: "ITEM Nº"  
  - D19: "Description"
  - E19: "Quantity"
  - F19: [merged with E for quantity]
  - G19: "N.W (kgs)"
  - H19: "G.W (kgs)"
  - I19: "CBM"

- **Sub-header Row (Row 20)**:
  - E20: "PCS"
  - F20: "SF"

**Column Spans (Merged Cells)**:
- Quantity spans E19:F19 with sub-headers E20:F20 (PCS and SF)
- Description is single column D
- Weight and volume columns are single-width

#### Key Column Span Patterns:
1. **Header Spans**: Company info spans A1:L1 (12 columns)
2. **Quantity Group**: E19:F19 merged header with E20, F20 sub-headers
3. **Data Columns**: 9 distinct columns (A,B,C,D,E,F,G,H,I)

## Column Span Detection Rules

### Pattern 1: Administrative Header Spans
- **Full-width company headers**: Span entire sheet width
- **3-column groupings**: Common for related fields (A:C, D:F, G:J)
- **Field labels**: Often span 1-3 columns

### Pattern 2: Data Table Spans  
- **Description columns**: Wide spans (4-6 columns) for product details
- **Financial columns**: Typically single-width for precise alignment
- **Quantity groups**: 2-column spans with sub-headers for different units

### Pattern 3: Vertical Merging
- **Related items**: Same description across multiple product variants
- **Continuation fields**: Multi-row spans for lengthy text

## Configuration Recommendations

### Auto-Detection Algorithm:
1. **Identify table start**: Look for header patterns like "Description", "Quantity", "Amount"
2. **Detect column spans**: 
   - Scan merged cell ranges
   - Group columns by content type (text vs numbers)
   - Identify repeating patterns
3. **Configure spans**:
   - Description areas: 4-6 column spans
   - Financial data: Single columns
   - Quantity groups: 2-column spans with sub-headers

### Key Indicators for CT/Contract Detection:
- Keywords: "CONTRACT", "SALES CONTRACT", "Contract Number"
- Typical location: Header area (rows 1-10)
- Pattern: Often in merged cells spanning multiple columns
- Associated with: Date fields, party information, terms

### Suggested Span Categories:
1. **HEADER_FULL**: Full sheet width (company info)
2. **HEADER_SECTION**: 3-4 column groupings (field groups)  
3. **DATA_DESCRIPTION**: 4-6 columns (product details)
4. **DATA_FINANCIAL**: Single columns (prices, amounts)
5. **DATA_QUANTITY**: 2 columns with sub-headers (units)

## Implementation Notes
- Monitor for patterns across multiple similar documents
- Consider document type variations (CT, INV, PL)
- Account for different column counts per sheet type
- Implement flexible detection for various merge patterns
