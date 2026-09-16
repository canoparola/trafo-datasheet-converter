"""
Configuration for Transformer Datasheet Converter
Maps project-specific CSV format to generic Excel output
"""

# Sheet definitions
SHEETS = {
    'Sheet1': 'HEADER',           # Antet (to be added later)
    'Sheet2': 'NOTES_AND_SCOPE',  # Notes, Reference Docs, Scope
    'Sheet3': 'HOLDS_AND_LEGENDS', # Holds and Legends
    'Sheet4': 'DATASHEET_PART1',  # Datasheet Part 1
    'Sheet5': 'DATASHEET_PART2',  # Datasheet Part 2
}

# CSV Input files
INPUT_FILES = [
    'input/1.csv',  # Temel bilgiler
    'input/2.csv',  # Detaylı bilgiler
]

# Notes and documentation paths
NOTES_PATH = 'notes/project_notes.txt'
HOLDS_PATH = 'notes/hold_notes.txt'

# Output file
OUTPUT_FILE = 'output/transformer_datasheet.xlsx'

# CSV column names
CSV_COLUMNS = {
    'sl_no': 'SL. NO',
    'rev': 'REV',
    'technical_particulars': 'TECHNICAL\nPARTICULARS',
    'purchaser_requirements': 'PURCHASER\'S\nREQUIREMENTS',
    'unit': 'UNIT',
    'vendor_response': 'VENDOR\'S\nRESPONSE',
}

# Datasheet sections (to split between Sheet4 and Sheet5)
DATASHEET_SECTIONS = {
    'PART1': [
        'General',
        'Service Conditions',
        'Power System Data',
        'Basic Performance',
        'Terminations',
    ],
    'PART2': [
        'Protection Device Wiring',
        'Accessories',
        'Current Transformers',
        'Condition Monitoring',
        'Electrical Performance',
        'Construction Data',
        'Inspection and Testing',
    ],
}

# Clean up rules
CLEANUP = {
    'remove_empty_rows_in_datasheet': True,
    'keep_single_blank_lines': True,  # Keep one blank line between sections
    'remove_trailing_whitespace': True,
}

# Excel formatting
EXCEL_FORMAT = {
    'font_name': 'Calibri',
    'font_size': 11,
    'header_font_size': 12,
    'header_bold': True,
    'column_widths': {
        'sl_no': 8,
        'rev': 8,
        'technical_particulars': 35,
        'purchaser_requirements': 25,
        'unit': 15,
        'vendor_response': 25,
    },
}

# Encoding
CSV_ENCODING = 'utf-8'
