# Excel Export Error Fixes - Testing Guide

## What Was Fixed

This update resolves critical issues with Excel file export and adds alternative export formats:

### 1. Merged Cells Lost in Export ✅ [CRITICAL FIX]
**Problem**: When exporting files with signatures, merged cells were being lost, causing:
- "파일 오류" (File Error) messages when opening in Excel
- Merged cells being completely separated/unmerged
- Content formatting being distorted
- Files requiring recovery mode to open

**Root Cause**: ExcelJS library bug - when images are added to a workbook, merged cell information can be lost during save operation (documented in ExcelJS issues #2641, #2146, #2755).

**Solution**: 
- Explicitly preserve merged cell information from original file
- After adding all signature images, re-apply merged cells to worksheet
- This ensures merged cells are properly written to the output file
- Comprehensive logging to track restoration process

### 2. Merged Cell Recognition Problem ✅
**Problem**: Signatures were being placed in the middle of merged cells, causing Excel errors and corrupted files.

**Solution**: 
- The system now detects merged cells during file parsing
- Signatures are only placed in the top-left cell of merged ranges
- Cells within merged ranges (but not the top-left) are automatically skipped

### 3. Print Area Detection ✅
**Problem**: Print area settings were not being properly detected, and Excel files with print areas would lose these settings.

**Solution**:
- Enhanced print area parser supports all Excel formats
- Print area settings are now preserved in the output file
- Better handling when print area is not set

### 4. Print Area Export Issues ✅ [NEW FIX]
**Problem**: Exported Excel files had broken formatting and extra rows beyond the print area, causing:
- Files to become very large
- Formatting to break when opened
- Extra empty rows extending far beyond intended content
- 작성완료한 엑셀파일을 열었을때 에러가 뜨며 최대복구를 하여 열어보았으나 기존 양식의 틀이 다 깨지고, 행수도 많이 늘어나있음

**Solution**:
- Added logic to clear all rows and columns outside the print area before saving
- Prevents worksheet expansion beyond print area boundaries
- Reduces file size significantly
- Maintains proper formatting structure

### 5. Alternative Export Formats 🎉 [NEW FEATURE]
**Problem**: Users needed alternative export formats in case Excel export was problematic.
- 엑셀파일로 내보내기가 어렵다면 PDF파일이나 이미지파일로도 내보내기 해줬으면 좋겠어

**Solution**:
- Added PDF export using jsPDF and html2canvas
- Added PNG image export using html2canvas
- Format selector UI in the preview toolbar
- All formats respect print area settings
- Signatures are properly rendered in all formats

## How to Test

### Test Case 1: Excel File with Merged Cells

1. Create or use an Excel file with:
   - Some merged cells in the data area
   - A column with names (성명/이름)
   - Signature markers (1, o, ○) in some cells

2. Upload the Excel file to the application
3. Upload signature images
4. Process the file and download the result
5. Open the result in Excel

**Expected Result**:
- File opens without errors
- Merged cells remain merged
- Signatures appear only in the top-left cell of merged ranges
- No signatures in the middle of merged cells

### Test Case 2: Excel File with Print Area

1. Create an Excel file:
   - Set a print area (Page Layout → Print Area → Set Print Area)
   - Add names and signature markers
   - Save the file

2. Upload to the application and process
3. Download the result
4. Open in Excel and check Print Preview (Ctrl+P)

**Expected Result**:
- Print area is preserved
- Only the designated print area is included when printing
- Signatures are placed only within the print area

### Test Case 3: Complex File (Most Important!)

1. Create an Excel file with:
   - Multiple merged cell ranges in various locations
   - Print area set to specific range (e.g., A1:H50)
   - Names in various cells
   - Signature markers in various locations

2. Process through the application
3. Download and open the result file

**Expected Result**:
- ✅ **File opens without any errors or warnings**
- ✅ All merged cells preserved exactly as in original
- ✅ Print area preserved
- ✅ Signatures correctly placed in valid cells
- ✅ No need for Excel's "recovery mode"
- ✅ Content formatting matches original

### Test Case 4: Previously Failing Files

If you have Excel files that previously failed with errors:
1. Try processing them again with this fix
2. They should now work correctly
3. Check console logs for merge restoration messages

## Console Logging

The application now provides detailed logging. Open browser console (F12) to see:

- `[parseExcelFile]` - Information about merged cells and print area
- `[인쇄영역파싱]` - Print area parsing details
- `[autoMatch]` - Auto-matching information
- `[병합셀]` - Merged cell information from original file
- `[병합셀 복원]` - **NEW** Merged cell restoration process
- `[병합셀 복원 완료]` - **NEW** Summary of merge restoration (how many restored/kept/failed)
- `[최종확인]` - Final validation before saving with success indicators

Key messages to look for:
- `✅ 병합된 셀이 성공적으로 보존되었습니다!` - Merged cells preserved successfully
- `⚠️ 경고: 병합된 셀 수가 여전히 다릅니다!` - Warning if merge count differs
- `✓ 병합 복원: A1:B2` - Individual merge cell restoration

These logs help debug any issues.

## Known Limitations

1. **ExcelJS Library**: The underlying ExcelJS library has some limitations with very complex Excel features, but the workaround we implemented addresses the most common issue
2. **Large Files**: Very large Excel files (>5MB) may cause memory issues
3. **Complex Formulas**: Some complex formulas might not be preserved exactly

## Troubleshooting

### Issue: Excel file won't open after processing (SHOULD BE FIXED NOW!)
This was the primary issue that has been addressed. If you still encounter this:
**Check**:
- Look at browser console for `[병합셀 복원]` messages
- Ensure you see `✅ 병합된 셀이 성공적으로 보존되었습니다!` message
- If you see errors during merge restoration, note which cells failed
- Ensure original file is valid XLSX format
- Try with a simpler test file first

### Issue: Signatures in wrong locations
**Check**:
- Console logs show merged cell detection
- Ensure print area is properly set in original file
- Check that signature markers (1, o, ○) are in valid cells

### Issue: Merged cells are not preserved
**Check**:
- Console for warnings: `병합된 셀 수가 변경되었습니다`
- This might be an ExcelJS limitation with specific merge types
- Try simplifying the merged cell structure

## Technical Details

### Files Modified/Added
- `types.ts` - Added merged cells and print area to data types
- `services/excelService.ts` - Core improvements (300+ lines of changes)
  - Print area restriction logic
  - Improved cell clearing
- `services/excelUtils.ts` - **NEW** Shared utility functions
- `services/alternativeExportService.ts` - **NEW** PDF and PNG export
- `App.tsx` - Added export format selector UI
- `package.json` - Added html2canvas and jspdf dependencies
- `.gitignore` - Added to prevent committing dependencies
- `CHANGELOG.md` - Documentation of changes

### New Functions
- `isCellInMergedRange` - Check if cell is in merged range
- `isTopLeftOfMergedCell` - Check if cell is top-left of merge
- `parseCellAddress` - Parse cell addresses
- `isValidPrintAreaRange` - Validate print area ranges
- `canPlaceSignature` - Validate signature placement
- `columnLetterToNumber` - Convert Excel columns to numbers
- `columnNumberToLetter` - Convert numbers to Excel columns
- `isSignaturePlaceholder` - Check if value is a placeholder
- `exportToPDF` - **NEW** Export sheet as PDF
- `exportToPNG` - **NEW** Export sheet as PNG image

### Quality Assurance
- ✅ Build successful
- ✅ TypeScript compilation clean
- ✅ Security scan passed (CodeQL - 0 alerts)
- ✅ No vulnerabilities in dependencies
- ✅ Code review completed (all feedback addressed)

## New Test Cases for Alternative Exports

### Test Case 5: PDF Export ✨

1. Create an Excel file with print area set
2. Upload to application and process
3. Select "PDF" format in the toolbar
4. Click download button
5. Open the PDF file

**Expected Result**:
- PDF opens without errors
- Shows only content within print area
- Signatures are properly rendered
- Layout matches Excel preview
- File size is reasonable

### Test Case 6: PNG Image Export ✨

1. Upload and process an Excel file
2. Select "PNG" format in the toolbar
3. Click download button
4. Open the PNG image

**Expected Result**:
- Image opens without errors
- High quality rendering (2x scale)
- Shows only content within print area
- Signatures are visible and clear
- Background is white

### Test Case 7: Format Switching

1. Upload and process an Excel file
2. Try exporting in each format:
   - Excel (.xlsx)
   - PDF (.pdf)
   - PNG (.png)
3. Verify each download works

**Expected Result**:
- All three formats download successfully
- Filenames include correct extension
- Content is consistent across formats
- Format selector shows correct selection

## Reporting Issues

If you encounter problems:

1. Check browser console logs (F12 → Console tab)
2. Try with a simpler test file
3. Note the specific file characteristics (merged cells, print area, etc.)
4. For PDF/PNG exports, check if the issue occurs in Excel export too
5. Provide console error messages if available

## Next Steps

1. Test with your actual Excel files
2. Try all three export formats (Excel, PDF, PNG)
3. Report any issues found
4. Verify all edge cases work correctly
5. Consider additional improvements if needed
