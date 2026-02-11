# Changelog

## [Unreleased] - 2026-02-11

### Fixed
- **CRITICAL: Merged Cells Lost in Export** ⚠️ [NEW]
  - Fixed issue where merged cells were being lost when exporting files with signatures
  - Resolved ExcelJS library limitation by explicitly re-applying merged cells after adding images
  - Files now open without errors in Excel
  - Content formatting and structure are preserved correctly
  - Issue: 최종 작성한 파일을 내보내기를 한 파일을 열면 오류파일이 뜨며 병합된 셀들도 전부 풀어짐

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
- **Excel Export Handling**: [IMPROVED]
  - Changed strategy to explicitly re-apply merged cells after adding images (ExcelJS bug workaround)
  - Enhanced logging to track merge cell restoration process
  - Better error handling for merge operations
  - Success/failure tracking for each merge operation
  
- **Excel Parsing**: `parseExcelFile` now extracts and returns merged cells and print area information
- **Auto-Matching**: Enhanced to respect merged cells and only match in valid locations
- **Signature Placement**: Improved to check both print area and merged cell constraints
- **File Preservation**: Original file structure is now better preserved with explicit merge restoration

### Technical Details
- Addresses known ExcelJS library limitation (GitHub issues #2641, #2146, #2755)
- Merged cells are now explicitly re-applied after adding images to prevent loss
- All changes maintain backward compatibility
- Build process verified successfully
- No security vulnerabilities introduced (verified with CodeQL)
- Code quality improvements based on review feedback

## Usage Notes

### For Developers
The improved Excel handling now properly:
1. Detects merged cells during file parsing
2. Preserves merged cell information during processing
3. Re-applies merged cells after adding images (workaround for ExcelJS bug)
4. Validates signature placement locations
5. Preserves original file structure (merged cells, print areas, etc.)
6. Provides comprehensive logging for debugging

### For Users
- **✅ Files now open without errors** - No more corrupted file warnings
- Excel files with merged cells now work correctly and stay merged
- Content formatting and structure are preserved
- Print area settings are now properly preserved
- Signatures are placed only in valid locations
- Better error messages when issues occur
