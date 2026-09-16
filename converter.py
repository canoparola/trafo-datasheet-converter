"""
Transformer Datasheet Converter
Converts project-specific CSV to generic Excel format
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os
from config import (
    INPUT_FILES, OUTPUT_FILE, NOTES_PATH, HOLDS_PATH,
    DATASHEET_SECTIONS, CSV_COLUMNS, EXCEL_FORMAT, CSV_ENCODING,
    CLEANUP
)

class TransformerDatasheetConverter:
    def __init__(self):
        self.workbook = Workbook()
        self.workbook.remove(self.workbook.active)  # Remove default sheet
        self.project_data = {}
        self.datasheet_data = []
        
    def read_csv_files(self):
        """Read and parse CSV input files"""
        print("[1/6] Reading CSV files...")
        
        all_data = []
        for file_path in INPUT_FILES:
            if os.path.exists(file_path):
                print(f"  Reading: {file_path}")
                df = pd.read_csv(file_path, encoding=CSV_ENCODING)
                all_data.append(df)
            else:
                print(f"  WARNING: File not found: {file_path}")
        
        if all_data:
            self.datasheet_data = pd.concat(all_data, ignore_index=True)
            print(f"  Total rows read: {len(self.datasheet_data)}")
        else:
            print("  ERROR: No CSV files found!")
            
    def clean_datasheet_data(self):
        """Clean and process datasheet data"""
        print("[2/6] Cleaning data...")
        
        # Remove completely empty rows
        self.datasheet_data = self.datasheet_data.dropna(how='all')
        
        # Fill NaN values in key columns
        self.datasheet_data.fillna('', inplace=True)
        
        # Remove leading/trailing whitespace
        if CLEANUP['remove_trailing_whitespace']:
            for col in self.datasheet_data.columns:
                if self.datasheet_data[col].dtype == 'object':
                    self.datasheet_data[col] = self.datasheet_data[col].str.strip()
        
        print(f"  Rows after cleaning: {len(self.datasheet_data)}")
        
    def read_notes_and_holds(self):
        """Read project notes and hold notes"""
        print("[3/6] Reading Notes and Holds...")
        
        self.project_notes = ""
        self.hold_notes = ""
        
        if os.path.exists(NOTES_PATH):
            with open(NOTES_PATH, 'r', encoding=CSV_ENCODING) as f:
                self.project_notes = f.read()
            print(f"  Project Notes loaded ({len(self.project_notes)} chars)")
        
        if os.path.exists(HOLDS_PATH):
            with open(HOLDS_PATH, 'r', encoding=CSV_ENCODING) as f:
                self.hold_notes = f.read()
            print(f"  Hold Notes loaded ({len(self.hold_notes)} chars)")
    
    def create_sheet2_notes_and_scope(self):
        """Create Sheet 2: Notes, Reference Docs, Scope"""
        print("[4/6] Creating Sheet 2 (Notes and Scope)...")
        
        ws = self.workbook.create_sheet(title="Notes & Scope")
        
        # Header styling
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, size=12, color="FFFFFF")
        
        # Title
        ws['A1'] = "PROJECT NOTES"
        ws['A1'].font = header_font
        ws['A1'].fill = header_fill
        ws.merge_cells('A1:D1')
        
        # Notes content
        row = 2
        for line in self.project_notes.split('\n'):
            if line.strip():
                ws[f'A{row}'] = line
                row += 1
            elif CLEANUP['keep_single_blank_lines']:
                row += 1
        
        # Reference docs title
        row += 2
        ws[f'A{row}'] = "REFERENCE DOCUMENTS"
        ws[f'A{row}'].font = header_font
        ws[f'A{row}'].fill = header_fill
        ws.merge_cells(f'A{row}:D{row}')
        
        ws.column_dimensions['A'].width = 80
        
    def create_sheet3_holds_and_legends(self):
        """Create Sheet 3: Holds and Legends"""
        print("[5/6] Creating Sheet 3 (Holds and Legends)...")
        
        ws = self.workbook.create_sheet(title="Holds & Legends")
        
        # Header styling
        header_fill = PatternFill(start_color="C55A11", end_color="C55A11", fill_type="solid")
        header_font = Font(bold=True, size=12, color="FFFFFF")
        
        # Title
        ws['A1'] = "HOLD NOTES"
        ws['A1'].font = header_font
        ws['A1'].fill = header_fill
        ws.merge_cells('A1:D1')
        
        # Hold notes content
        row = 2
        for line in self.hold_notes.split('\n'):
            if line.strip():
                ws[f'A{row}'] = line
                row += 1
            elif CLEANUP['keep_single_blank_lines']:
                row += 1
        
        # Legends
        row += 2
        ws[f'A{row}'] = "LEGENDS & ABBREVIATIONS"
        ws[f'A{row}'].font = header_font
        ws[f'A{row}'].fill = header_fill
        ws.merge_cells(f'A{row}:D{row}')
        
        row += 1
        legends = [
            "VTA: Vendor To Answer",
            "TBD: To Be Determined",
            "HV: High Voltage (66 kV)",
            "LV: Low Voltage (11 kV)",
            "kVA: Kilovolt-Amperes",
            "CT: Current Transformer",
            "DGA: Dissolved Gas Analysis",
        ]
        
        for legend in legends:
            ws[f'A{row}'] = legend
            row += 1
        
        ws.column_dimensions['A'].width = 80
        
    def create_sheet4_datasheet_part1(self):
        """Create Sheet 4: Datasheet Part 1"""
        print("[6a/6] Creating Sheet 4 (Datasheet Part 1)...")
        
        ws = self.workbook.create_sheet(title="Datasheet - Part 1")
        self._write_datasheet_to_sheet(ws, is_part2=False)
        
    def create_sheet5_datasheet_part2(self):
        """Create Sheet 5: Datasheet Part 2"""
        print("[6b/6] Creating Sheet 5 (Datasheet Part 2)...")
        
        ws = self.workbook.create_sheet(title="Datasheet - Part 2")
        self._write_datasheet_to_sheet(ws, is_part2=True)
        
    def _write_datasheet_to_sheet(self, ws, is_part2=False):
        """Write datasheet content to sheet"""
        
        # Header styling
        header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        header_font = Font(bold=True, size=11, color="FFFFFF")
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Column headers
        headers = ['SL. NO', 'REV', 'TECHNICAL PARTICULARS', "PURCHASER'S REQUIREMENTS", 'UNIT', "VENDOR'S RESPONSE"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col)
            cell.value = header
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(wrap_text=True)
            cell.border = border
        
        # Set column widths
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 8
        ws.column_dimensions['C'].width = 35
        ws.column_dimensions['D'].width = 25
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 25
        
        # Write data
        current_row = 2
        for idx, row_data in self.datasheet_data.iterrows():
            for col, value in enumerate(row_data, 1):
                cell = ws.cell(row=current_row, column=col)
                cell.value = value
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                cell.border = border
            current_row += 1
        
    def save_workbook(self):
        """Save workbook to Excel file"""
        print(f"Saving to {OUTPUT_FILE}...")
        
        # Create output directory if it doesn't exist
        output_dir = os.path.dirname(OUTPUT_FILE)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        self.workbook.save(OUTPUT_FILE)
        print(f"✓ Excel file created: {OUTPUT_FILE}")
        
    def run(self):
        """Run the complete conversion process"""
        print("\n" + "="*60)
        print("Transformer Datasheet Converter")
        print("="*60 + "\n")
        
        self.read_csv_files()
        self.clean_datasheet_data()
        self.read_notes_and_holds()
        self.create_sheet2_notes_and_scope()
        self.create_sheet3_holds_and_legends()
        self.create_sheet4_datasheet_part1()
        self.create_sheet5_datasheet_part2()
        self.save_workbook()
        
        print("\n" + "="*60)
        print("Conversion completed successfully!")
        print("="*60 + "\n")


if __name__ == "__main__":
    converter = TransformerDatasheetConverter()
    converter.run()
