"""
PositionUpdater component for updating start_row and column positions in configuration templates.

This module provides functionality to update start_row values and column positions in
header_to_write sections while preserving rowspan/colspan attributes and template structure.
"""

from typing import Dict, List, Any, Optional
import copy
from .models import QuantityAnalysisData, SheetData, HeaderPosition
from .mapping_manager import MappingManager, MappingManagerError


class PositionUpdaterError(Exception):
    """Custom exception for PositionUpdater errors."""
    pass


class PositionUpdater:
    """Updates start_row and column positions in configuration templates."""
    
    def __init__(self, mapping_config_path: str = "mapping_config.json"):
        """Initialize PositionUpdater with mapping manager."""
        # Start row mappings based on analysis data (Contract: 18, Invoice: 21, Packing list: 22)
        self.start_row_mappings = {
            'Contract': 18,
            'Invoice': 21,
            'Packing list': 22
        }
        
        # Initialize mapping manager
        try:
            self.mapping_manager = MappingManager(mapping_config_path)
        except MappingManagerError as e:
            print(f"Warning: Could not load mapping config: {e}")
            # Fallback to default mappings
            self.mapping_manager = None
    
    def update_positions(self, template: Dict[str, Any], quantity_data: QuantityAnalysisData) -> Dict[str, Any]:
        """
        Update both start_row values and row heights using analysis data.
        
        Args:
            template: Configuration template dictionary
            quantity_data: Quantity analysis data containing position information
            
        Returns:
            Updated template with positions and row heights updated
            
        Raises:
            PositionUpdaterError: If template structure is invalid or update fails
        """
        # First update start rows
        updated_template = self.update_start_rows(template, quantity_data)
        
        # Then update row heights
        updated_template = self.update_row_heights(updated_template, quantity_data)
        
        return updated_template
    
    def update_start_rows(self, template: Dict[str, Any], quantity_data: QuantityAnalysisData) -> Dict[str, Any]:
        """
        Update start_row values using analysis data while preserving template structure.
        
        Args:
            template: Configuration template dictionary
            quantity_data: Quantity analysis data containing start row information
            
        Returns:
            Updated template with start_row values replaced
            
        Raises:
            PositionUpdaterError: If template structure is invalid or update fails
        """
        try:
            if not isinstance(template, dict):
                raise PositionUpdaterError("Template must be a dictionary")
            
            if not isinstance(quantity_data, QuantityAnalysisData):
                raise PositionUpdaterError("Quantity data must be QuantityAnalysisData instance")
            
            # Validate template structure
            self._validate_template_structure(template)
            
            # Create deep copy to avoid modifying original template
            updated_template = copy.deepcopy(template)
            
            # Process each sheet in the template
            data_mapping = updated_template.get('data_mapping', {})
            
            # Track sheets with missing start row data
            missing_start_row_sheets = []
            
            for sheet_data in quantity_data.sheets:
                quantity_sheet_name = sheet_data.sheet_name
                mapped_sheet_name = self._map_sheet_name(quantity_sheet_name)
                
                if mapped_sheet_name not in data_mapping:
                    missing_start_row_sheets.append(f"{quantity_sheet_name} -> {mapped_sheet_name}")
                    continue
                
                # Validate start row data
                self._validate_start_row_data(sheet_data, quantity_sheet_name)
                    
                sheet_config = data_mapping[mapped_sheet_name]
                
                # Validate sheet config structure
                if not isinstance(sheet_config, dict):
                    raise PositionUpdaterError(f"Sheet config for '{mapped_sheet_name}' must be a dictionary")
                
                # Update start_row using analysis data
                sheet_config['start_row'] = sheet_data.start_row
            
            # Apply fallback strategies for missing start row data
            if missing_start_row_sheets:
                self._apply_start_row_fallback_strategies(updated_template, missing_start_row_sheets)
            
            return updated_template
            
        except Exception as e:
            if isinstance(e, PositionUpdaterError):
                raise
            raise PositionUpdaterError(f"Start row update failed: {str(e)}") from e
    
    def update_column_positions(self, template: Dict[str, Any], quantity_data: QuantityAnalysisData) -> Dict[str, Any]:
        """
        Update column positions in header_to_write sections while preserving rowspan/colspan.
        
        Args:
            template: Configuration template dictionary
            quantity_data: Quantity analysis data containing header positions
            
        Returns:
            Updated template with column positions adjusted
            
        Raises:
            PositionUpdaterError: If template structure is invalid or update fails
        """
        try:
            if not isinstance(template, dict):
                raise PositionUpdaterError("Template must be a dictionary")
            
            if not isinstance(quantity_data, QuantityAnalysisData):
                raise PositionUpdaterError("Quantity data must be QuantityAnalysisData instance")
            
            # Validate template structure
            self._validate_template_structure(template)
            
            # Create deep copy to avoid modifying original template
            updated_template = copy.deepcopy(template)
            
            # Process each sheet in the template
            data_mapping = updated_template.get('data_mapping', {})
            
            # Track sheets with missing position data
            missing_position_sheets = []
            
            for sheet_data in quantity_data.sheets:
                quantity_sheet_name = sheet_data.sheet_name
                mapped_sheet_name = self._map_sheet_name(quantity_sheet_name)
                
                if mapped_sheet_name not in data_mapping:
                    missing_position_sheets.append(f"{quantity_sheet_name} -> {mapped_sheet_name}")
                    continue
                
                # Validate header positions data
                self._validate_header_positions_data(sheet_data, quantity_sheet_name)
                    
                sheet_config = data_mapping[mapped_sheet_name]
                header_to_write = sheet_config.get('header_to_write', [])
                
                # Update column positions for this sheet
                self._update_sheet_column_positions_with_validation(header_to_write, sheet_data.header_positions, mapped_sheet_name)
            
            # Apply fallback strategies for missing position data
            if missing_position_sheets:
                self._apply_position_fallback_strategies(updated_template, missing_position_sheets)
            
            return updated_template
            
        except Exception as e:
            if isinstance(e, PositionUpdaterError):
                raise
            raise PositionUpdaterError(f"Column position update failed: {str(e)}") from e
    

    
    def _map_sheet_name(self, quantity_sheet_name: str) -> str:
        """
        Map quantity data sheet name to template config sheet name.
        
        Args:
            quantity_sheet_name: Sheet name from quantity data
            
        Returns:
            Mapped sheet name for template config, or original name if no mapping found
        """
        if self.mapping_manager:
            return self.mapping_manager.map_sheet_name(quantity_sheet_name)
        
        # Fallback to hardcoded mappings if mapping manager is not available
        fallback_mappings = {
            'INV': 'Invoice',
            'PAK': 'Packing list',
            'CON': 'Contract',
            'CONTRACT': 'Contract',
            'INVOICE': 'Invoice',
            'PACKING': 'Packing list',
            'PACKING LIST': 'Packing list'
        }
        
        return fallback_mappings.get(quantity_sheet_name.upper(), quantity_sheet_name)
    
    def _normalize_header_text(self, text: str) -> str:
        """
        Normalize header text for comparison by removing special characters and whitespace.
        
        Args:
            text: Header text to normalize
            
        Returns:
            Normalized text string
        """
        if not isinstance(text, str):
            return ''
        
        # Remove common variations and normalize
        normalized = text.lower().strip()
        
        # Replace common character variations
        normalized = normalized.replace('º', '°').replace('nº', 'n°')
        normalized = normalized.replace('\n', ' ').replace('  ', ' ')
        
        # Remove punctuation for better matching
        normalized = normalized.replace('.', '').replace('&', 'and')
        
        return normalized
    

    
    def _validate_template_structure(self, template: Dict[str, Any]) -> None:
        """
        Validate template structure for position updates.
        
        Args:
            template: Template dictionary to validate
            
        Raises:
            PositionUpdaterError: If template structure is invalid
        """
        if 'data_mapping' not in template:
            raise PositionUpdaterError("Template missing 'data_mapping' section")
        
        data_mapping = template['data_mapping']
        if not isinstance(data_mapping, dict):
            raise PositionUpdaterError("Template 'data_mapping' must be a dictionary")
        
        # Validate each sheet configuration
        for sheet_name, sheet_config in data_mapping.items():
            if not isinstance(sheet_config, dict):
                raise PositionUpdaterError(f"Sheet config for '{sheet_name}' must be a dictionary")
            
            if 'header_to_write' not in sheet_config:
                raise PositionUpdaterError(f"Sheet '{sheet_name}' missing 'header_to_write' section")
            
            header_to_write = sheet_config['header_to_write']
            if not isinstance(header_to_write, list):
                raise PositionUpdaterError(f"'header_to_write' for sheet '{sheet_name}' must be a list")
    
    def _validate_start_row_data(self, sheet_data, sheet_name: str) -> None:
        """
        Validate start row data from quantity analysis.
        
        Args:
            sheet_data: Sheet data containing start row information
            sheet_name: Name of the sheet for error messages
            
        Raises:
            PositionUpdaterError: If start row data is invalid
        """
        if not hasattr(sheet_data, 'start_row'):
            raise PositionUpdaterError(f"Sheet '{sheet_name}' missing start_row data")
        
        if not isinstance(sheet_data.start_row, int) or sheet_data.start_row < 0:
            raise PositionUpdaterError(f"Sheet '{sheet_name}' start_row must be a non-negative integer")
    
    def _validate_header_positions_data(self, sheet_data, sheet_name: str) -> None:
        """
        Validate header positions data from quantity analysis.
        
        Args:
            sheet_data: Sheet data containing header positions
            sheet_name: Name of the sheet for error messages
            
        Raises:
            PositionUpdaterError: If header positions data is invalid
        """
        if not hasattr(sheet_data, 'header_positions'):
            raise PositionUpdaterError(f"Sheet '{sheet_name}' missing header_positions data")
        
        if not isinstance(sheet_data.header_positions, list):
            raise PositionUpdaterError(f"Sheet '{sheet_name}' header_positions must be a list")
        
        for i, header_pos in enumerate(sheet_data.header_positions):
            if not isinstance(header_pos, HeaderPosition):
                raise PositionUpdaterError(f"Sheet '{sheet_name}' header_position {i} must be HeaderPosition instance")
    
    def _update_sheet_column_positions_with_validation(self, header_to_write: List[Dict[str, Any]], 
                                                     header_positions: List[HeaderPosition], sheet_name: str) -> None:
        """
        Update column positions for a single sheet while preserving spans with validation.
        
        Args:
            header_to_write: List of header entries to update
            header_positions: Header positions from quantity analysis
            sheet_name: Name of the sheet for error messages
            
        Raises:
            PositionUpdaterError: If update fails
        """
        try:
            # Validate inputs
            if not isinstance(header_to_write, list):
                raise PositionUpdaterError(f"header_to_write for sheet '{sheet_name}' must be a list")
            
            if not isinstance(header_positions, list):
                raise PositionUpdaterError(f"header_positions for sheet '{sheet_name}' must be a list")
            
            # Create mapping of header text to column positions from analysis data
            text_to_position = {}
            
            for i, header_pos in enumerate(header_positions):
                if not isinstance(header_pos, HeaderPosition):
                    raise PositionUpdaterError(f"header_position {i} for sheet '{sheet_name}' must be HeaderPosition instance")
                
                if not hasattr(header_pos, 'keyword') or not header_pos.keyword:
                    continue  # Skip invalid keywords
                
                if not hasattr(header_pos, 'column') or not isinstance(header_pos.column, int):
                    continue  # Skip invalid column data
                
                text_to_position[header_pos.keyword] = {
                    'row': getattr(header_pos, 'row', 0),
                    'column': header_pos.column
                }
            
            # Update header_to_write entries with actual column positions
            for i, header_entry in enumerate(header_to_write):
                if not isinstance(header_entry, dict):
                    raise PositionUpdaterError(f"header_entry {i} for sheet '{sheet_name}' must be a dictionary")
                
                if 'text' not in header_entry:
                    continue  # Skip entries without text
                
                header_text = header_entry['text']
                
                # Look for exact match first
                if header_text in text_to_position:
                    position = text_to_position[header_text]
                    # Validate column position before updating
                    if isinstance(position['column'], int) and position['column'] >= 0:
                        header_entry['col'] = position['column'] - 1  # Convert to 0-based indexing
                else:
                    # Try to find similar header text (case-insensitive, normalized)
                    normalized_text = self._normalize_header_text(header_text)
                    
                    for analysis_text, position in text_to_position.items():
                        normalized_analysis = self._normalize_header_text(analysis_text)
                        
                        if normalized_text == normalized_analysis:
                            if isinstance(position['column'], int) and position['column'] >= 0:
                                header_entry['col'] = position['column'] - 1  # Convert to 0-based indexing
                            break
                            
        except Exception as e:
            if isinstance(e, PositionUpdaterError):
                raise
            raise PositionUpdaterError(f"Failed to update column positions for sheet '{sheet_name}': {str(e)}") from e
    
    def _apply_start_row_fallback_strategies(self, template: Dict[str, Any], missing_sheets: List[str]) -> None:
        """
        Apply fallback strategies for sheets with missing start row data.
        
        Args:
            template: Template dictionary to update
            missing_sheets: List of sheet names missing start row data
        """
        if not missing_sheets:
            return
        
        # Log missing start row data for manual review
        print(f"Warning: Missing start row data for sheets: {missing_sheets}")
        
        # Use default start row mappings as fallback
        data_mapping = template.get('data_mapping', {})
        
        for sheet_name in missing_sheets:
            if sheet_name in data_mapping and sheet_name in self.start_row_mappings:
                data_mapping[sheet_name]['start_row'] = self.start_row_mappings[sheet_name]
    
    def _apply_position_fallback_strategies(self, template: Dict[str, Any], missing_sheets: List[str]) -> None:
        """
        Apply fallback strategies for sheets with missing position data.
        
        Args:
            template: Template dictionary to update
            missing_sheets: List of sheet names missing position data
        """
        if not missing_sheets:
            return
        
        # Log missing position data for manual review
        print(f"Warning: Missing position data for sheets: {missing_sheets}")
        
        # Keep existing column positions as fallback (no changes needed) 
   
    def update_row_heights(self, template: Dict[str, Any], quantity_data: QuantityAnalysisData) -> Dict[str, Any]:
        """
        Update row_heights in styling sections using actual Excel row heights.
        
        Args:
            template: Configuration template dictionary
            quantity_data: Quantity analysis data containing row height information
            
        Returns:
            Updated template with row_heights extracted from Excel
            
        Raises:
            PositionUpdaterError: If template structure is invalid or update fails
        """
        print("🔍 [DEBUG] update_row_heights() method called!")
        
        try:
            if not isinstance(template, dict):
                raise PositionUpdaterError("Template must be a dictionary")
            
            if not isinstance(quantity_data, QuantityAnalysisData):
                raise PositionUpdaterError("Quantity data must be QuantityAnalysisData instance")
            
            print(f"🔍 [DEBUG] Processing {len(quantity_data.sheets)} sheets for row heights")
            
            # Store Excel file path for height extraction
            self._excel_file_path = quantity_data.file_path
            print(f"📏 [HEIGHT_SETUP] Excel file path: {self._excel_file_path}")
            
            # Create deep copy to avoid modifying original template
            updated_template = copy.deepcopy(template)
            
            # Update footer configurations with correct total text column IDs
            self.update_footer_configurations(updated_template, quantity_data)
            
            # Process each sheet in the template
            data_mapping = updated_template.get('data_mapping', {})
            
            for sheet_data in quantity_data.sheets:
                quantity_sheet_name = sheet_data.sheet_name
                mapped_sheet_name = self._map_sheet_name(quantity_sheet_name)
                
                print(f"🔍 [DEBUG] Processing sheet: {quantity_sheet_name} -> {mapped_sheet_name}")
                
                if mapped_sheet_name not in data_mapping:
                    print(f"🔍 [DEBUG] Sheet {mapped_sheet_name} not found in data_mapping, skipping")
                    continue
                    
                sheet_config = data_mapping[mapped_sheet_name]
                
                print(f"🔍 [DEBUG] Extracting row heights for {mapped_sheet_name}")
                
                # Extract row heights from the sheet data
                row_heights = self._extract_row_heights_from_sheet(sheet_data, mapped_sheet_name)
                
                # Update the styling section with extracted row heights
                if 'styling' not in sheet_config:
                    sheet_config['styling'] = {}
                    print(f"🔍 [DEBUG] Created styling section for {mapped_sheet_name}")
                
                sheet_config['styling']['row_heights'] = row_heights
                
                print(f"✅ [ROW_HEIGHTS] Updated {mapped_sheet_name}: {row_heights}")
            
            return updated_template
            
        except Exception as e:
            if isinstance(e, PositionUpdaterError):
                raise
            raise PositionUpdaterError(f"Row height update failed: {str(e)}") from e
    
    def _extract_row_heights_from_sheet(self, sheet_data: SheetData, sheet_name: str) -> Dict[str, float]:
        """
        Extract row heights from sheet data using actual Excel row heights when possible.
        
        Args:
            sheet_data: Sheet data containing row height information
            sheet_name: Name of the sheet for logging
            
        Returns:
            Dictionary with row height information
        """
        row_heights = {}
        
        # Try to use actual Excel heights first
        excel_heights = self._get_actual_excel_heights(sheet_data, sheet_name)
        if excel_heights:
            row_heights.update(excel_heights)
            print(f"📏 [HEIGHT_EXCEL] Using actual Excel heights for {sheet_name}: {excel_heights}")
        else:
            # Fallback to font-based estimation
            print(f"📏 [HEIGHT_FALLBACK] Using font-based estimation for {sheet_name}")
            row_heights = self._get_font_based_heights(sheet_data, sheet_name)
        
        # Special case for Packing list - add before_footer height
        if sheet_name == 'Packing list':
            row_heights['before_footer'] = row_heights['data_default']
        
        print(f"[ROW_HEIGHTS] Extracted for {sheet_name}: header={row_heights['header']}, data={row_heights['data_default']}, footer={row_heights['footer']}")
        
        return row_heights
    
    def _get_actual_excel_heights(self, sheet_data: SheetData, sheet_name: str) -> Optional[Dict[str, float]]:
        """
        Get actual Excel row heights using ExcelHeightAnalyzer.
        
        Args:
            sheet_data: Sheet data containing row information
            sheet_name: Name of the sheet
            
        Returns:
            Dictionary with actual heights or None if Excel access fails
        """
        try:
            # Import here to avoid circular imports
            from .excel_height_analyzer import ExcelHeightAnalyzer
            
            # Get the Excel file path from quantity_data if available
            excel_file_path = getattr(sheet_data, 'excel_file_path', None)
            if not excel_file_path and hasattr(self, '_excel_file_path'):
                excel_file_path = self._excel_file_path
            
            if not excel_file_path:
                print(f"📏 [HEIGHT_EXCEL] No Excel file path available for {sheet_name}")
                return None
            
            analyzer = ExcelHeightAnalyzer(excel_file_path)
            
            # Get sheet structure
            structure = analyzer.analyze_sheet_structure(sheet_name)
            
            if structure['heights']:
                actual_heights = structure['heights']
                
                # Validate heights are within reasonable ranges
                validated_heights = self._validate_height_ranges(actual_heights, sheet_name)
                return validated_heights
            else:
                print(f"📏 [HEIGHT_EXCEL] Could not extract structure for {sheet_name}")
                return None
                
        except Exception as e:
            print(f"📏 [HEIGHT_EXCEL] Error accessing Excel heights for {sheet_name}: {e}")
            return None
    
    def _validate_height_ranges(self, heights: Dict[str, float], sheet_name: str) -> Dict[str, float]:
        """
        Validate and adjust height values to be within reasonable ranges.
        
        Args:
            heights: Dictionary of height values
            sheet_name: Name of the sheet for logging
            
        Returns:
            Dictionary with validated height values
        """
        validated = {}
        
        # Height validation ranges (expanded to accommodate real Excel files)
        ranges = {
            'header': (10, 80),      # Headers can be quite tall
            'data_default': (10, 60), # Data rows typically smaller
            'footer': (10, 70),      # Footers can be medium height
            'before_footer': (10, 60)
        }
        
        for key, value in heights.items():
            if key in ranges:
                min_val, max_val = ranges[key]
                if min_val <= value <= max_val:
                    validated[key] = value
                    print(f"📏 [HEIGHT_VALIDATION] {sheet_name} {key}: {value}pt ✅ VALID")
                else:
                    # Use the closest valid value
                    validated[key] = max(min_val, min(value, max_val))
                    print(f"📏 [HEIGHT_VALIDATION] {sheet_name} {key}: {value}pt -> {validated[key]}pt (clamped to range {min_val}-{max_val})")
            else:
                validated[key] = value
        
        return validated
    
    def _get_font_based_heights(self, sheet_data: SheetData, sheet_name: str) -> Dict[str, float]:
        """
        Get row heights based on font sizes (fallback method).
        
        Args:
            sheet_data: Sheet data containing font information
            sheet_name: Name of the sheet for logging
            
        Returns:
            Dictionary with font-based height estimates
        """
        row_heights = {}
        
        # Extract header row height (from the header font)
        if hasattr(sheet_data, 'header_font') and sheet_data.header_font:
            header_font_size = sheet_data.header_font.size
            # Rule of thumb: row height ≈ font size * 1.8 + padding for headers
            estimated_header_height = max(header_font_size * 1.8, 25)
            row_heights['header'] = round(estimated_header_height)
        else:
            row_heights['header'] = 30  # Default fallback
        
        # Extract data row height (from the data font)
        if hasattr(sheet_data, 'data_font') and sheet_data.data_font:
            data_font_size = sheet_data.data_font.size
            # Rule of thumb: data row height ≈ font size * 1.6 + padding for data
            estimated_data_height = max(data_font_size * 1.6, 20)
            row_heights['data_default'] = round(estimated_data_height)
        else:
            row_heights['data_default'] = 25  # Default fallback
        
        # Extract footer height (from actual footer font if available)
        if hasattr(sheet_data, 'footer_info') and sheet_data.footer_info:
            footer_font_size = sheet_data.footer_info.font.size
            # Footer with formulas (SUM) may need more height for readability
            if sheet_data.footer_info.has_formulas:
                # Formula rows need extra height for better visibility
                estimated_footer_height = max(footer_font_size * 2.0, 30)
                print(f"[FOOTER_DETECTION] {sheet_name}: Formula footer detected at row {sheet_data.footer_info.row}")
            else:
                # Regular footer height
                estimated_footer_height = max(footer_font_size * 1.8, 25)
            
            row_heights['footer'] = round(estimated_footer_height)
            print(f"[FOOTER_DETECTION] {sheet_name}: Footer font size {footer_font_size}, height {row_heights['footer']}")
        else:
            # Fallback: Use header height for footer (common pattern)
            row_heights['footer'] = row_heights['header']
            print(f"[FOOTER_DETECTION] {sheet_name}: No footer info found, using header height {row_heights['footer']}")
        
        return row_heights
    
    def update_footer_configurations(self, template: Dict[str, Any], quantity_data: QuantityAnalysisData) -> None:
        """
        Update footer configurations with correct total_text_column_id based on actual Excel footer analysis.
        
        Args:
            template: Template dictionary to update
            quantity_data: Processed quantity analysis data containing footer information
        """
        print("🔍 [FOOTER_CONFIG] Starting footer configuration updates...")
        
        data_mapping = template.get('data_mapping', {})
        
        for sheet_name, sheet_config in data_mapping.items():
            print(f"🔍 [FOOTER_CONFIG] Processing sheet: {sheet_name}")
            
            # Find corresponding sheet data
            sheet_data = None
            for data_sheet in quantity_data.sheets:
                if data_sheet.sheet_name == sheet_name:
                    sheet_data = data_sheet
                    break
            
            if not sheet_data:
                print(f"🔍 [FOOTER_CONFIG] Sheet '{sheet_name}' not found in quantity data, skipping")
                continue
            
            # Check if sheet has footer info with total text column or pallet count column
            footer_updates = {}
            
            # Handle total text column
            if hasattr(sheet_data, 'footer_info') and sheet_data.footer_info and sheet_data.footer_info.total_text_column:
                total_text_column = sheet_data.footer_info.total_text_column
                total_text_value = sheet_data.footer_info.total_text_value
                
                print(f"📊 [FOOTER_CONFIG] Found total text '{total_text_value}' in column {total_text_column} for {sheet_name}")
                
                # Map Excel column to column ID
                correct_column_id = self._map_excel_column_to_column_id(sheet_config, total_text_column, sheet_name)
                
                if correct_column_id:
                    footer_updates['total_text_column_id'] = correct_column_id
                    footer_updates['total_text'] = total_text_value
                    print(f"✅ [FOOTER_CONFIG] Will update {sheet_name}: total_text_column_id → '{correct_column_id}'")
                else:
                    print(f"❌ [FOOTER_CONFIG] Could not map total text column {total_text_column} to column ID for {sheet_name}")
            
            # Handle pallet count column
            if hasattr(sheet_data, 'footer_info') and sheet_data.footer_info and sheet_data.footer_info.pallet_count_column:
                pallet_count_column = sheet_data.footer_info.pallet_count_column
                pallet_count_value = sheet_data.footer_info.pallet_count_value
                
                print(f"📦 [FOOTER_CONFIG] Found pallet count '{pallet_count_value}' in column {pallet_count_column} for {sheet_name}")
                
                # Map Excel column to column ID
                correct_pallet_column_id = self._map_excel_column_to_column_id(sheet_config, pallet_count_column, sheet_name)
                
                if correct_pallet_column_id:
                    footer_updates['pallet_count_column_id'] = correct_pallet_column_id
                    print(f"✅ [FOOTER_CONFIG] Will update {sheet_name}: pallet_count_column_id → '{correct_pallet_column_id}'")
                else:
                    print(f"❌ [FOOTER_CONFIG] Could not map pallet count column {pallet_count_column} to column ID for {sheet_name}")
            
            # Apply updates if any were found
            if footer_updates:
                footer_config = sheet_config.get('footer_configurations', {})
                
                for key, value in footer_updates.items():
                    old_value = footer_config.get(key, 'none')
                    footer_config[key] = value
                    print(f"✅ [FOOTER_CONFIG] Updated {sheet_name}: {key}: '{old_value}' → '{value}'")
            else:
                print(f"📊 [FOOTER_CONFIG] No footer updates needed for {sheet_name}")
    
    def _map_excel_column_to_column_id(self, sheet_config: Dict[str, Any], excel_column: int, sheet_name: str) -> Optional[str]:
        """
        Map Excel column number to corresponding column ID based on header_to_write configuration.
        
        Args:
            sheet_config: Sheet configuration containing header_to_write
            excel_column: Excel column number (1-based)
            sheet_name: Sheet name for logging
            
        Returns:
            Column ID string or None if mapping not found
        """
        header_to_write = sheet_config.get('header_to_write', [])
        
        # Convert Excel column (1-based) to config column (0-based)
        config_column = excel_column - 1
        
        print(f"🔍 [COLUMN_MAPPING] Looking for column {config_column} (Excel col {excel_column}) in {sheet_name} headers")
        
        # Search through header_to_write entries
        for header_entry in header_to_write:
            if isinstance(header_entry, dict) and 'col' in header_entry and 'id' in header_entry:
                if header_entry['col'] == config_column:
                    column_id = header_entry['id']
                    print(f"✅ [COLUMN_MAPPING] Found mapping: Excel col {excel_column} → {column_id}")
                    return column_id
        
        # If exact match not found, try to find closest match or fallback
        print(f"⚠️ [COLUMN_MAPPING] No exact match for column {config_column}, trying fallbacks...")
        
        # Common fallbacks based on column position
        if config_column == 0:  # First column often col_static
            return 'col_static'
        elif config_column == 1:  # Second column often col_po
            return 'col_po'
        elif config_column == 2:  # Third column often col_item
            return 'col_item'
        
        print(f"❌ [COLUMN_MAPPING] No mapping found for Excel column {excel_column} in {sheet_name}")
        return None