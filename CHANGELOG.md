# Changelog

## [Unreleased] - 2026-02-11

### Fixed
- **Merged Cell Recognition**: Fixed issue where signatures were incorrectly placed in merged cells
  - Signatures are now only placed in the top-left cell of merged ranges
  - Auto-matching logic now properly detects and skips non-top-left cells in merged ranges
  - Added comprehensive merged cell detection functions

- **Print Area Detection**: Fixed print area settings not being properly detected and preserved
  - Enhanced print area parsing to support multiple formats:
    - Simple ranges: `A1:C10`
    - Sheet-qualified ranges: `Sheet1!A1:C10`
    - Absolute references: `$A$1:$C$10`
    - Combined formats: `Sheet1!$A$1:$C$10`
  - Added validation for print area ranges
  - Improved error handling and fallback to entire sheet when print area is not set

### Added
- **New Type Definitions**:
  - Added `mergedCells` field to `SheetData` interface
  - Added `printArea` field to `SheetData` interface

- **New Helper Functions**:
  - `columnNumberToLetter`: Convert column numbers to letters (e.g., 1 → "A", 27 → "AA")
  - `parseCellAddress`: Parse cell addresses like "A1" into row/column numbers
  - `isCellInMergedRange`: Check if a cell is within a merged cell range
  - `isTopLeftOfMergedCell`: Check if a cell is the top-left cell of a merged range
  - `isValidPrintAreaRange`: Validate print area range parameters

- **Enhanced Logging**:
  - Added detailed logging for merged cell detection
  - Added print area parsing status logs
  - Added validation logs for signature placement
  - Added final verification logs before saving

### Changed
- **Excel Parsing**: `parseExcelFile` now extracts and returns merged cells and print area information
- **Auto-Matching**: Enhanced to respect merged cells and only match in valid locations
- **Signature Placement**: Improved to check both print area and merged cell constraints
- **File Preservation**: Original file structure is now better preserved by avoiding manipulation of merged cells and print areas

### Technical Details
- All changes maintain backward compatibility
- Build process verified successfully
- No security vulnerabilities introduced (verified with CodeQL)
- Code quality improvements based on review feedback

## Usage Notes

### For Developers
The improved Excel handling now properly:
1. Detects merged cells during file parsing
2. Validates signature placement locations
3. Preserves original file structure (merged cells, print areas, etc.)
4. Provides comprehensive logging for debugging

### For Users
- Excel files with merged cells now work correctly
- Print area settings are now properly preserved
- Signatures are placed only in valid locations
- Better error messages when issues occur
