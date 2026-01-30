"""
BOM Processor - Extract and compress hierarchical BOM data from Excel files
Author: GitHub Copilot
Date: 2026-01-24

This script processes hierarchical BOM (Bill of Materials) Excel files by:
1. Extracting Level 2 items and their children into separate sheets
2. Creating compressed versions with deduplicated parts and aggregated quantities
"""

import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from collections import defaultdict
import datetime
import os
from typing import Dict, List, Tuple, Any


class BOMProcessor:
    """Process hierarchical BOM Excel files"""
    
    def __init__(self, input_file: str = None, output_file: str = None, summary_file: str = None, log_file: str = None, log_callback=None):
        self.input_file = input_file
        self.output_file = output_file
        self.summary_file = summary_file
        self.log_file = log_file
        self.log_callback = log_callback  # Callback function for real-time log updates
        self.log_messages = []
        self.start_time = datetime.datetime.now()
        
        # Statistics
        self.stats = {
            'level1_count': 0,
            'level2_count': 0,
            'sheets_created': [],
            'rows_per_sheet': {},
            'dedup_stats': {}
        }
    
    def log(self, message: str, level: str = "INFO"):
        """Add message to log"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}"
        self.log_messages.append(log_entry)
        
        # Call callback if provided (for Streamlit real-time updates)
        if self.log_callback:
            self.log_callback(log_entry)
        else:
            # Default: print to console (for CLI usage)
            print(log_entry)
    
    def write_log_file(self):
        """Write all log messages to file"""
        with open(self.log_file, 'w', encoding='utf-8') as f:
            f.write("="*80 + "\n")
            f.write("BOM PROCESSING LOG\n")
            f.write("="*80 + "\n\n")
            f.write(f"Start Time: {self.start_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Input File: {self.input_file}\n")
            f.write(f"Output File: {self.output_file}\n")
            f.write("\n" + "="*80 + "\n")
            f.write("PROCESSING LOG\n")
            f.write("="*80 + "\n\n")
            
            for msg in self.log_messages:
                f.write(msg + "\n")
            
            f.write("\n" + "="*80 + "\n")
            f.write("STATISTICS SUMMARY\n")
            f.write("="*80 + "\n\n")
            f.write(f"Level 1 items found: {self.stats['level1_count']}\n")
            f.write(f"Level 2 items found: {self.stats['level2_count']}\n")
            f.write(f"Sheets created: {len(self.stats['sheets_created'])}\n")
            f.write(f"Sheet names: {', '.join(self.stats['sheets_created'])}\n\n")
            
            f.write("Rows per sheet:\n")
            for sheet_name, count in self.stats['rows_per_sheet'].items():
                f.write(f"  {sheet_name}: {count} rows\n")
            
            if self.stats['dedup_stats']:
                f.write("\nDeduplication statistics:\n")
                for sheet_name, stats in self.stats['dedup_stats'].items():
                    f.write(f"  {sheet_name}:\n")
                    f.write(f"    Original rows: {stats['original']}\n")
                    f.write(f"    After dedup: {stats['deduplicated']}\n")
                    f.write(f"    Removed: {stats['removed']}\n")
            
            end_time = datetime.datetime.now()
            duration = end_time - self.start_time
            f.write(f"\nEnd Time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"Total Duration: {duration.total_seconds():.2f} seconds\n")
            f.write("\n" + "="*80 + "\n")
            f.write("PROCESSING COMPLETED SUCCESSFULLY\n")
            f.write("="*80 + "\n")
    
    def load_workbook_data(self) -> Tuple[Any, Any, Dict[str, int], List[Dict]]:
        """Load workbook and extract hierarchical data with outline levels"""
        self.log(f"Loading workbook: {self.input_file}")
        
        try:
            wb = load_workbook(self.input_file, data_only=True, keep_vba=False)
            self.log(f"Workbook loaded successfully. Sheets: {wb.sheetnames}")
        except Exception as e:
            self.log(f"Error loading workbook: {e}", "ERROR")
            raise
        
        # Get the main sheet (first sheet)
        main_sheet = wb.worksheets[0]
        sheet_name = main_sheet.title
        self.log(f"Processing main sheet: '{sheet_name}' ({main_sheet.max_row} rows, {main_sheet.max_column} cols)")
        
        # Extract headers from row 1
        headers = {}
        for col_idx in range(1, main_sheet.max_column + 1):
            cell_value = main_sheet.cell(1, col_idx).value
            if cell_value:
                headers[str(cell_value).strip()] = col_idx
        
        self.log(f"Found {len(headers)} column headers")
        
        # Required columns
        required_cols = ['BOM Line', 'CAD OEM Part Number', 'CAD OEM Rev', 'Quantity']
        missing_cols = [col for col in required_cols if col not in headers]
        if missing_cols:
            self.log(f"WARNING: Missing required columns: {missing_cols}", "WARNING")
        
        # Extract all data rows with outline levels
        data_rows = []
        for row_idx in range(2, main_sheet.max_row + 1):
            row_dim = main_sheet.row_dimensions[row_idx]
            outline_level = row_dim.outline_level if row_dim.outline_level else 0
            
            row_data = {'_row_num': row_idx, '_outline_level': outline_level}
            
            for col_name, col_idx in headers.items():
                cell_value = main_sheet.cell(row_idx, col_idx).value
                row_data[col_name] = cell_value
            
            data_rows.append(row_data)
        
        self.log(f"Extracted {len(data_rows)} data rows")
        
        return wb, main_sheet, headers, data_rows
    
    def analyze_hierarchy(self, data_rows: List[Dict]) -> Tuple[List[Dict], List[Dict]]:
        """Analyze and extract Level 1 and Level 2 items"""
        self.log("Analyzing hierarchy levels...")
        
        level1_items = [row for row in data_rows if row['_outline_level'] == 0]
        level2_items = [row for row in data_rows if row['_outline_level'] == 1]
        
        self.stats['level1_count'] = len(level1_items)
        self.stats['level2_count'] = len(level2_items)
        
        self.log(f"Level 1 items (outline_level=0): {len(level1_items)}")
        if level1_items:
            for item in level1_items[:3]:
                bom_line = item.get('BOM Line', 'N/A')
                self.log(f"  - Row {item['_row_num']}: {bom_line}")
        
        self.log(f"Level 2 items (outline_level=1): {len(level2_items)}")
        if level2_items:
            for item in level2_items[:5]:
                bom_line = item.get('BOM Line', 'N/A')
                self.log(f"  - Row {item['_row_num']}: {bom_line}")
        
        # Verify Level 1 has only one item
        if len(level1_items) != 1:
            self.log(f"WARNING: Expected 1 Level 1 item, found {len(level1_items)}", "WARNING")
        else:
            self.log("✓ Verified: Level 1 has exactly 1 item")
        
        return level1_items, level2_items
    
    def extract_level2_branches(self, data_rows: List[Dict], level2_items: List[Dict]) -> Dict[str, List[Dict]]:
        """Extract Level 2 items and their children into separate groups"""
        self.log("Extracting Level 2 branches...")
        
        branches = {}
        
        for i, level2_item in enumerate(level2_items):
            bom_line = level2_item.get('BOM Line', '')
            
            # Extract sheet name from text after the last '-' in BOM Line
            if bom_line and '-' in bom_line:
                # Find the last '-' and get everything after it
                sheet_name = bom_line.split('-')[-1].strip().upper()
                if not sheet_name:  # If nothing after last '-'
                    sheet_name = f"L2_{i+1}"
                    self.log(f"WARNING: Empty text after last '-' in BOM Line '{bom_line}', using {sheet_name}", "WARNING")
            else:
                sheet_name = f"L2_{i+1}"
                self.log(f"WARNING: No '-' found in BOM Line '{bom_line}', using {sheet_name}", "WARNING")

            self.log(f"  Extracted sheet name '{sheet_name}' from BOM Line: {bom_line}")

            # Handle duplicate sheet names
            original_name = sheet_name
            counter = 1
            while sheet_name in branches:
                sheet_name = f"{original_name}_{counter}"
                counter += 1
            
            # Find all children of this Level 2 item
            level2_row_num = level2_item['_row_num']
            
            # Determine the range of children (until next Level 2 or end)
            next_level2_row = None
            if i + 1 < len(level2_items):
                next_level2_row = level2_items[i + 1]['_row_num']
            
            # Collect Level 2 item and all its children
            branch_rows = [level2_item]
            
            for row in data_rows:
                row_num = row['_row_num']
                
                # Skip if before current Level 2 item
                if row_num <= level2_row_num:
                    continue
                
                # Stop if reached next Level 2 item
                if next_level2_row and row_num >= next_level2_row:
                    break
                
                # Add if it's a child (outline_level >= 2)
                if row['_outline_level'] >= 2:
                    branch_rows.append(row)
            
            branches[sheet_name] = branch_rows
            self.log(f"Branch '{sheet_name}': {len(branch_rows)} rows (Level 2 + children)")
            self.stats['rows_per_sheet'][sheet_name] = len(branch_rows)
        
        return branches
    
    def create_output_workbook(self, headers: Dict[str, int], branches: Dict[str, List[Dict]]) -> Any:
        """Create output workbook with extracted sheets"""
        self.log("Creating output workbook...")
        
        output_wb = openpyxl.Workbook()
        output_wb.remove(output_wb.active)  # Remove default sheet
        
        # Sort headers by column index
        sorted_headers = sorted(headers.items(), key=lambda x: x[1])
        header_names = [h[0] for h in sorted_headers]
        
        for sheet_name, rows in branches.items():
            self.log(f"Creating sheet '{sheet_name}' with {len(rows)} rows")
            
            ws = output_wb.create_sheet(title=sheet_name)
            
            # Write headers
            for col_idx, header_name in enumerate(header_names, start=1):
                ws.cell(1, col_idx, header_name)
            
            # Write data rows
            for row_idx, row_data in enumerate(rows, start=2):
                for col_idx, header_name in enumerate(header_names, start=1):
                    value = row_data.get(header_name)
                    ws.cell(row_idx, col_idx, value)
            
            self.stats['sheets_created'].append(sheet_name)
        
        self.log(f"Created {len(branches)} extraction sheets")
        return output_wb
    
    def create_compressed_sheets(self, output_wb: Any, branches: Dict[str, List[Dict]]):
        """Create compressed sheets with deduplication and quantity aggregation"""
        self.log("Creating compressed sheets...")
        
        compress_columns = [
            'CAD OEM Part Number',
            'CAD OEM Rev',
            'Material Spec',
            'CAD Oem Name',
            'Thickness',
            'Quantity'
        ]
        
        for sheet_name, rows in branches.items():
            compress_sheet_name = f"{sheet_name}_Compress"
            self.log(f"Creating compressed sheet '{compress_sheet_name}'")
            
            ws = output_wb.create_sheet(title=compress_sheet_name)
            
            # Write headers
            for col_idx, col_name in enumerate(compress_columns, start=1):
                ws.cell(1, col_idx, col_name)
            
            # Group by (CAD OEM Part Number, CAD OEM Rev)
            grouped_data = defaultdict(lambda: {
                'Material Spec': None,
                'CAD Oem Name': None,
                'Thickness': None,
                'Quantity': 0
            })
            
            original_count = 0
            for row_data in rows:
                part_num = row_data.get('CAD OEM Part Number')
                rev = row_data.get('CAD OEM Rev')
                
                # Skip rows without part number
                if not part_num:
                    continue
                
                original_count += 1
                
                key = (part_num, rev)
                
                # Get quantity and normalize (empty/0 -> 1)
                qty = row_data.get('Quantity')
                if qty is None or qty == '' or qty == 0 or qty == '0' or qty == 0.0:
                    qty = 1
                else:
                    try:
                        qty = float(qty)
                    except (ValueError, TypeError):
                        self.log(f"WARNING: Invalid quantity '{qty}' in {sheet_name}, treating as 1", "WARNING")
                        qty = 1
                
                # First occurrence: store all fields
                if grouped_data[key]['Material Spec'] is None:
                    grouped_data[key]['Material Spec'] = row_data.get('Material Spec')
                    grouped_data[key]['CAD Oem Name'] = row_data.get('CAD Oem Name')
                    grouped_data[key]['Thickness'] = row_data.get('Thickness')
                
                # Accumulate quantity
                grouped_data[key]['Quantity'] += qty
            
            # Write deduplicated data
            row_idx = 2
            for (part_num, rev), data in sorted(grouped_data.items()):
                ws.cell(row_idx, 1, part_num)
                ws.cell(row_idx, 2, rev)
                ws.cell(row_idx, 3, data['Material Spec'])
                ws.cell(row_idx, 4, data['CAD Oem Name'])
                ws.cell(row_idx, 5, data['Thickness'])
                ws.cell(row_idx, 6, data['Quantity'])
                row_idx += 1
            
            deduplicated_count = len(grouped_data)
            removed_count = original_count - deduplicated_count
            
            self.stats['dedup_stats'][compress_sheet_name] = {
                'original': original_count,
                'deduplicated': deduplicated_count,
                'removed': removed_count
            }
            
            self.log(f"  Original rows: {original_count}, After dedup: {deduplicated_count}, Removed: {removed_count}")
            self.stats['sheets_created'].append(compress_sheet_name)
            self.stats['rows_per_sheet'][compress_sheet_name] = deduplicated_count
        
        self.log(f"Created {len(branches)} compressed sheets")
    

    def create_unified_summary(self, output_wb: Any) -> openpyxl.Workbook:
        """Create unified BOM summary sheet consolidating all compress sheets.
        Returns a new workbook containing only the BOM_Summary sheet.
        """
        self.log("Creating unified BOM summary...")
        
        # Collect all compress sheets
        compress_sheets = {}
        for sheet_name in output_wb.sheetnames:
            if sheet_name.endswith('_Compress'):
                base_name = sheet_name.replace('_Compress', '')
                compress_sheets[base_name] = output_wb[sheet_name]
        
        if not compress_sheets:
            self.log("WARNING: No compress sheets found to create summary", "WARNING")
            return
        
        self.log(f"Found {len(compress_sheets)} compress sheets to consolidate")
        
        # Data structure: {(part_num, rev): {metadata, {sheet_type: quantity}}}
        unified_data = defaultdict(lambda: {
            'Material Spec': None,
            'CAD Oem Name': None,
            'Thickness': None,
            'quantities': defaultdict(float),
            'metadata_sources': set()
        })
        
        # Collect data from all compress sheets
        for sheet_type, ws in compress_sheets.items():
            self.log(f"  Processing compress sheet: {sheet_type}")
            
            # Read data from compress sheet (skip header row)
            for row_idx in range(2, ws.max_row + 1):
                part_num = ws.cell(row_idx, 1).value
                rev = ws.cell(row_idx, 2).value
                material_spec = ws.cell(row_idx, 3).value
                cad_oem_name = ws.cell(row_idx, 4).value
                thickness = ws.cell(row_idx, 5).value
                quantity = ws.cell(row_idx, 6).value
                
                if not part_num:
                    continue
                
                key = (part_num, rev)
                
                # Store or verify metadata
                if unified_data[key]['Material Spec'] is None:
                    unified_data[key]['Material Spec'] = material_spec
                    unified_data[key]['CAD Oem Name'] = cad_oem_name
                    unified_data[key]['Thickness'] = thickness
                    unified_data[key]['metadata_sources'].add(sheet_type)
                else:
                    # Check for metadata conflicts
                    if unified_data[key]['Material Spec'] != material_spec:
                        self.log(f"WARNING: Part {part_num}/{rev} has different Material Spec: '{unified_data[key]['Material Spec']}' vs '{material_spec}' in sheet {sheet_type}", "WARNING")
                    if unified_data[key]['CAD Oem Name'] != cad_oem_name:
                        self.log(f"WARNING: Part {part_num}/{rev} has different CAD Oem Name: '{unified_data[key]['CAD Oem Name']}' vs '{cad_oem_name}' in sheet {sheet_type}", "WARNING")
                    if unified_data[key]['Thickness'] != thickness:
                        self.log(f"WARNING: Part {part_num}/{rev} has different Thickness: '{unified_data[key]['Thickness']}' vs '{thickness}' in sheet {sheet_type}", "WARNING")
                
                # Store quantity for this sheet type
                unified_data[key]['quantities'][sheet_type] = quantity if quantity else 0
        
        # Create new workbook for BOM_Summary
        summary_wb = openpyxl.Workbook()
        summary_ws = summary_wb.active
        summary_ws.title = 'BOM_Summary'
        self.log("Creating BOM_Summary sheet")
        
        # Get sorted sheet types (maintain original order from workbook)
        sheet_types = list(compress_sheets.keys())
        
        # Write headers
        headers = ['CAD OEM Part Number', 'CAD OEM Rev', 'Material Spec', 'CAD Oem Name', 'Thickness']
        headers.extend(sheet_types)
        
        for col_idx, header in enumerate(headers, start=1):
            summary_ws.cell(1, col_idx, header)
        
        # Write data rows
        row_idx = 2
        for (part_num, rev), data in sorted(unified_data.items()):
            summary_ws.cell(row_idx, 1, part_num)
            summary_ws.cell(row_idx, 2, rev)
            summary_ws.cell(row_idx, 3, data['Material Spec'])
            summary_ws.cell(row_idx, 4, data['CAD Oem Name'])
            summary_ws.cell(row_idx, 5, data['Thickness'])
            
            # Write quantities for each sheet type
            for col_idx, sheet_type in enumerate(sheet_types, start=6):
                qty = data['quantities'].get(sheet_type, 0)
                summary_ws.cell(row_idx, col_idx, qty)
            
            row_idx += 1
        
        unique_parts = len(unified_data)
        self.log(f"✓ Created BOM_Summary with {unique_parts} unique parts across {len(sheet_types)} sheet types")
        self.stats['sheets_created'].append('BOM_Summary')
        self.stats['rows_per_sheet']['BOM_Summary'] = unique_parts
        
        return summary_wb

    def process_from_buffer(self, uploaded_file_buffer, filename="uploaded_file.xlsm"):
        """
        Process BOM from uploaded file buffer (for Web interface)
        
        Args:
            uploaded_file_buffer: File-like object (BytesIO) from file upload
            filename: Original filename for logging
            
        Returns:
            tuple: (output_buffer, summary_buffer, log_text, stats_dict)
                - output_buffer: BytesIO containing extraction workbook
                - summary_buffer: BytesIO containing BOM summary
                - log_text: Complete log as text string
                - stats_dict: Dictionary with processing statistics
        """
        from io import BytesIO
        
        try:
            self.log("="*80)
            self.log("BOM PROCESSING STARTED (Web Mode)")
            self.log("="*80)
            self.log(f"Processing uploaded file: {filename}")
            
            # Temporarily set input_file for logging purposes
            self.input_file = filename
            
            # Step 1: Load workbook from buffer
            self.log(f"Loading workbook from uploaded file...")
            wb = load_workbook(uploaded_file_buffer, data_only=True, keep_vba=False)
            self.log(f"Workbook loaded successfully. Sheets: {wb.sheetnames}")
            
            # Get the main sheet (first sheet)
            main_sheet = wb.worksheets[0]
            sheet_name = main_sheet.title
            self.log(f"Processing main sheet: '{sheet_name}' ({main_sheet.max_row} rows, {main_sheet.max_column} cols)")
            
            # Extract headers from row 1
            headers = {}
            for col_idx in range(1, main_sheet.max_column + 1):
                cell_value = main_sheet.cell(1, col_idx).value
                if cell_value:
                    headers[str(cell_value).strip()] = col_idx
            
            self.log(f"Found {len(headers)} column headers")
            
            # Required columns validation
            required_cols = ['BOM Line', 'CAD OEM Part Number', 'CAD OEM Rev', 'Quantity']
            missing_cols = [col for col in required_cols if col not in headers]
            if missing_cols:
                self.log(f"WARNING: Missing required columns: {missing_cols}", "WARNING")
            
            # Extract all data rows with outline levels
            data_rows = []
            for row_idx in range(2, main_sheet.max_row + 1):
                row_dim = main_sheet.row_dimensions[row_idx]
                outline_level = row_dim.outline_level if row_dim.outline_level else 0
                
                row_data = {'_row_num': row_idx, '_outline_level': outline_level}
                
                for col_name, col_idx in headers.items():
                    cell_value = main_sheet.cell(row_idx, col_idx).value
                    row_data[col_name] = cell_value
                
                data_rows.append(row_data)
            
            self.log(f"Extracted {len(data_rows)} data rows")
            
            # Step 2: Analyze hierarchy
            level1_items, level2_items = self.analyze_hierarchy(data_rows)
            
            # Step 3: Extract Level 2 branches
            branches = self.extract_level2_branches(data_rows, level2_items)
            
            # Step 4: Create output workbook with extracted sheets
            output_wb = self.create_output_workbook(headers, branches)
            
            # Step 5: Create compressed sheets
            self.create_compressed_sheets(output_wb, branches)
            
            # Step 6: Create unified BOM summary in separate file
            summary_wb = self.create_unified_summary(output_wb)
            
            # Step 7: Save to BytesIO buffers instead of files
            self.log("Preparing output files for download...")
            
            output_buffer = BytesIO()
            output_wb.save(output_buffer)
            output_buffer.seek(0)
            self.log("✓ Output workbook prepared")
            
            summary_buffer = BytesIO()
            summary_wb.save(summary_buffer)
            summary_buffer.seek(0)
            self.log("✓ BOM summary prepared")
            
            # Prepare log text
            log_text = self.get_log_content()
            self.log("✓ Log content prepared")
            
            self.log("="*80)
            self.log("BOM PROCESSING COMPLETED SUCCESSFULLY")
            self.log("="*80)
            
            return output_buffer, summary_buffer, log_text, self.stats
            
        except Exception as e:
            self.log(f"FATAL ERROR: {str(e)}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")
            raise
    
    def get_log_content(self):
        """Generate complete log content as string (for Web downloads)"""
        lines = []
        lines.append("="*80)
        lines.append("BOM PROCESSING LOG")
        lines.append("="*80)
        lines.append("")
        lines.append(f"Start Time: {self.start_time.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"Input File: {self.input_file}")
        lines.append("")
        lines.append("="*80)
        lines.append("PROCESSING LOG")
        lines.append("="*80)
        lines.append("")
        
        for msg in self.log_messages:
            lines.append(msg)
        
        lines.append("")
        lines.append("="*80)
        lines.append("STATISTICS SUMMARY")
        lines.append("="*80)
        lines.append("")
        lines.append(f"Level 1 items found: {self.stats['level1_count']}")
        lines.append(f"Level 2 items found: {self.stats['level2_count']}")
        lines.append(f"Sheets created: {len(self.stats['sheets_created'])}")
        lines.append(f"Sheet names: {', '.join(self.stats['sheets_created'])}")
        lines.append("")
        
        lines.append("Rows per sheet:")
        for sheet_name, count in self.stats['rows_per_sheet'].items():
            lines.append(f"  {sheet_name}: {count} rows")
        
        if self.stats['dedup_stats']:
            lines.append("")
            lines.append("Deduplication statistics:")
            for sheet_name, stats in self.stats['dedup_stats'].items():
                lines.append(f"  {sheet_name}:")
                lines.append(f"    Original rows: {stats['original']}")
                lines.append(f"    After dedup: {stats['deduplicated']}")
                lines.append(f"    Removed: {stats['removed']}")
        
        end_time = datetime.datetime.now()
        duration = end_time - self.start_time
        lines.append("")
        lines.append(f"End Time: {end_time.strftime('%Y-%m-%d %H:%M:%S')}")
        lines.append(f"Total Duration: {duration.total_seconds():.2f} seconds")
        lines.append("")
        lines.append("="*80)
        lines.append("PROCESSING COMPLETED SUCCESSFULLY")
        lines.append("="*80)
        
        return "\n".join(lines)

    def process(self):
        """Main processing pipeline (for CLI usage)"""
        try:
            self.log("="*80)
            self.log("BOM PROCESSING STARTED")
            self.log("="*80)
            
            # Step 1: Load workbook
            wb, main_sheet, headers, data_rows = self.load_workbook_data()
            
            # Step 2: Analyze hierarchy
            level1_items, level2_items = self.analyze_hierarchy(data_rows)
            
            # Step 3: Extract Level 2 branches
            branches = self.extract_level2_branches(data_rows, level2_items)
            
            # Step 4: Create output workbook with extracted sheets
            output_wb = self.create_output_workbook(headers, branches)
            
            # Step 5: Create compressed sheets
            self.create_compressed_sheets(output_wb, branches)
            
            # Step 6: Create unified BOM summary in separate file
            summary_wb = self.create_unified_summary(output_wb)
            
            # Step 7: Save output workbook
            self.log(f"Saving output workbook to: {self.output_file}")
            output_wb.save(self.output_file)
            self.log("✓ Output workbook saved successfully")
            
            # Step 8: Save BOM summary workbook
            self.log(f"Saving BOM summary to: {self.summary_file}")
            summary_wb.save(self.summary_file)
            self.log("✓ BOM summary saved successfully")
            
            # Step 9: Write log file
            self.log(f"Writing log file to: {self.log_file}")
            self.write_log_file()
            
            self.log("="*80)
            self.log("BOM PROCESSING COMPLETED SUCCESSFULLY")
            self.log("="*80)
            
            return True
            
        except Exception as e:
            self.log(f"FATAL ERROR: {str(e)}", "ERROR")
            import traceback
            self.log(traceback.format_exc(), "ERROR")
            
            # Try to write log even on failure
            try:
                self.write_log_file()
            except:
                pass
            
            raise


def main():
    """Main entry point"""
    # Generate date stamp for output files
    current_date = datetime.datetime.now().strftime("%Y%m%d")
    
    # File paths with date stamp
    input_file = r"c:\Users\zongyue.liu\Desktop\AdvSimulation\BOM_Summary\data\BOM-COS1000334701-BA.xlsm"
    file_name_wo_ext = os.path.splitext(os.path.basename(input_file))[0]
    current_time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    output_file = rf"c:\Users\zongyue.liu\Desktop\AdvSimulation\BOM_Summary\data\{file_name_wo_ext}_processed_{current_time_str}.xlsx"
    summary_file = rf"c:\Users\zongyue.liu\Desktop\AdvSimulation\BOM_Summary\data\{file_name_wo_ext}_BOM_Sum_{current_time_str}.xlsx"
    log_file = rf"c:\Users\zongyue.liu\Desktop\AdvSimulation\BOM_Summary\data\{file_name_wo_ext}_process_log_{current_time_str}.txt"
    
    # Verify input file exists
    if not os.path.exists(input_file):
        print(f"ERROR: Input file not found: {input_file}")
        return
    
    print(f"Input file: {input_file}")
    print(f"Output file: {output_file}")
    print(f"Summary file: {summary_file}")
    print(f"Log file: {log_file}")
    print()
    
    # Create processor and run
    processor = BOMProcessor(input_file, output_file, summary_file, log_file)
    processor.process()
    
    print()
    print(f"✓ Processing complete!")
    print(f"  Output: {output_file}")
    print(f"  Summary: {summary_file}")
    print(f"  Log: {log_file}")


if __name__ == "__main__":
    main()

