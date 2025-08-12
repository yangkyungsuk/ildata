# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**ildata** is a Korean construction estimate Excel parsing system that converts various Excel formats from Korean construction companies into standardized JSON format. The system handles multiple non-standardized Excel formats used in the construction industry, preserving complete data integrity including empty rows and original ordering.

## Commands

### Running Individual Parsers
```bash
# Parse specific file types (v4 is the current version)
python parser_sangcul_test1_v4.py          # For test1.xlsx
python parser_sangcul_sgs_v4.py            # For sgs.xls
python parser_sangcul_est_v4.py            # For est.xlsx
python parser_sangcul_ebs_v4.py            # For ebs.xls
python parser_sangcul_construction_v4.py   # For 건축구조내역.xlsx
python parser_sangcul_stmate_v4.py         # For stmate.xlsx 단가산출
python parser_ilwidae_stmate_v4.py         # For stmate.xlsx 일위대가
```

### Smart Parser (Auto-detection)
```bash
# Parse a single file with automatic type detection
python smart_sangcul_parser.py test1.xlsx

# Parse all supported files in the directory
python smart_sangcul_parser.py --all
```

### Batch Operations
```bash
# Run all parsers
python run_all_parsers.py

# Verify all parsing results
python verify_all_results.py
python verify_all_sangcul_parsers.py
```

### Installation
```bash
pip install pandas xlrd openpyxl xlwings
```

## Architecture

### Core Design Principles

1. **Complete Order Preservation**: All rows between hopyo (工票) are preserved in original order
2. **Data Integrity**: Even empty rows are preserved with `has_content: false` markers
3. **Pattern Flexibility**: Supports various hopyo patterns (제N호표, #N, No.N, 산근N, (산근N))

### File Structure Pattern

Each parser follows this structure:
1. **Hopyo Detection**: Identify section markers using regex patterns
2. **Complete Row Extraction**: Extract ALL rows between hopyo markers (including empty ones)
3. **JSON Output**: Generate structured JSON with validation metadata

### Standard JSON Output Structure
```json
{
  "file": "filename.xlsx",
  "sheet": "sheet_name",
  "total_hopyo_count": 12,
  "validation": {
    "list_count": 12,
    "parsed_count": 12,
    "match": true
  },
  "hopyo_data": {
    "호표1": {
      "호표번호": "1",
      "작업명": "work_name",
      "시작행": 2,
      "종료행": 50,
      "모든_행_데이터": [
        {
          "row_number": 3,
          "columns": [...],
          "has_content": true
        }
      ]
    }
  }
}
```

### Parser Selection Logic (smart_sangcul_parser.py)

The smart parser automatically detects file type based on:
1. Filename patterns (test1, sgs, est, ebs, stmate, construction)
2. Sheet structure analysis
3. Hopyo pattern detection

### Critical Implementation Notes

1. **UTF-8 Encoding**: All parsers must set `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')` for Korean text
2. **xlwings Requirement**: stmate.xlsx files require xlwings due to special Excel properties
3. **Row Preservation**: Never skip or filter rows - all data between hopyo markers must be preserved
4. **Validation**: Always compare parsed count with list count to ensure completeness

## Key Files

- `smart_sangcul_parser.py`: Main entry point with auto-detection
- `parser_sangcul_*_v4.py`: Individual parsers for each file type (v4 is current)
- `run_all_parsers.py`: Batch processing script
- `verify_all_*.py`: Validation and verification scripts
- `analyze_*.py`: Data structure analysis tools
- `check_*.py`: Data integrity checking tools

## Development Guidelines

When modifying parsers:
1. Preserve ALL rows between hopyo markers (including empty ones)
2. Maintain exact column order from source Excel
3. Include row numbers for traceability
4. Generate validation metadata comparing list vs parsed counts
5. Use UTF-8 encoding throughout
6. Test with actual construction company files before deployment