# Changelog

## [Unreleased] - 2026-02-19

### Fixed
- **Random Signature Selection Improvements** 🎯 [NEW]
  - Fixed issue where the same signature variant could be used multiple times in the same row
  - Improved natural variation by preferring unused variants within each row
  - Added safety checks for edge cases with invalid signature variants
  - Enhanced code readability with new `randomInt` helper function
  - Issue: 엑셀파일 내보내기를 할때에나 원본양식 그대로에 사인을 무작위 랜덤으로 넣는것에 대한 오류사항이 있는지 검증 및 분석하여 개선

### Added
- **New Utility Function** (`services/excelUtils.ts`):
  - `randomInt(min, max)`: Generate random integers in a cleaner, more readable way
  - Replaces complex `Math.floor(Math.random() * range) + offset` patterns
  - Makes random value generation more maintainable

### Changed
- **Signature Matching Algorithm** (`autoMatchSignatures`):
  - Now tracks used signature variants per row to avoid immediate reuse
  - When multiple placeholders exist in a row, different variants are preferred
  - Automatically cycles through variants when more placeholders than variants exist
  - Better logging for debugging signature selection process

## [Unreleased] - 2026-02-13

### Fixed
- **Print Area Export Issues** ⚠️ [NEW]
  - Fixed issue where exported Excel files had broken formatting and extra rows
  - Added logic to clear rows and columns outside the print area before saving
  - Prevents worksheet expansion beyond print area bounds
  - Reduces file size and prevents format corruption
  - Issue: 작성완료한 엑셀파일을 열었을때 에러가 뜨며 최대복구를 하여 열어보았으나 기존 양식의 틀이 다 깨지고, 행수도 많이 늘어나있음

### Added
- **Alternative Export Formats** 🎉 [NEW]
  - Added PDF export functionality using jsPDF and html2canvas
  - Added PNG image export functionality using html2canvas
  - Export format selector in UI (Excel/PDF/PNG)
  - All export formats respect print area settings
  - Issue: 엑셀파일로 내보내기가 어렵다면 PDF파일이나 이미지파일로도 내보내기 해줬으면 좋겠어

- **New Utility Module** (`services/excelUtils.ts`):
  - Shared utility functions for Excel operations
  - `columnLetterToNumber`: Convert Excel column letters to numbers
  - `columnNumberToLetter`: Convert numbers to Excel column letters
  - `parseCellAddress`: Parse cell addresses like "A1" into coordinates
  - `SIGNATURE_PLACEHOLDERS`: Constant array of signature placeholder values
  - `isSignaturePlaceholder`: Helper to check if value is a placeholder

- **New Export Service** (`services/alternativeExportService.ts`):
  - `exportToPDF`: Generate PDF documents from Excel sheets
  - `exportToPNG`: Generate PNG images from Excel sheets
  - Renders Excel sheets as HTML tables for conversion
  - Supports signature placement and formatting

### Changed
- **UI Improvements**:
  - Added export format selection buttons in preview toolbar
  - Visual indicators for selected export format
  - Format-specific file naming (with .xlsx, .pdf, or .png extension)
  - Enhanced user feedback for different export types

### Security
- **Dependency Updates**:
  - Updated `jspdf` from v2.5.2 to v4.1.0 (fixes 5 CVEs):
    - CVE: PDF Injection in AcroFormChoiceField
    - CVE: DoS via Unvalidated BMP Dimensions
    - CVE: Denial of Service (DoS)
    - CVE: ReDoS Bypass
    - CVE: Local File Inclusion/Path Traversal
  - Added `html2canvas` v1.4.1 (no known vulnerabilities)
  - CodeQL security scan: 0 alerts

### Technical Details
- Removed code duplication by extracting shared utilities
- Improved type safety with explicit null handling
- Enhanced error handling in async operations
- Added comprehensive documentation and constants
- All builds successful with TypeScript compilation clean

## [Unreleased] - 2026-02-11

### Fixed
- **TypeScript Compilation Errors** ⚠️ [NEW]
  - Fixed incorrect property access for merged cells in ExcelJS
  - Changed `worksheet.merged` to `worksheet.model.merges` (correct ExcelJS API)
  - Fixed null safety issue in auto-matching flow
  - Fixed blob verification code type errors
  - All TypeScript errors resolved, project now compiles cleanly

- **CRITICAL: Merged Cells Lost in Export** ⚠️
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
