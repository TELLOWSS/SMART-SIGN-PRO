# Excel Export Error Fixes - Testing Guide

## What Was Fixed

This update resolves two critical issues with Excel file export:

### 1. Merged Cell Recognition Problem ✅
**Problem**: Signatures were being placed in the middle of merged cells, causing Excel errors and corrupted files.

**Solution**: 
- The system now detects merged cells during file parsing
- Signatures are only placed in the top-left cell of merged ranges
- Cells within merged ranges (but not the top-left) are automatically skipped

### 2. Print Area Detection ✅
**Problem**: Print area settings were not being properly detected, and Excel files with print areas would lose these settings.

**Solution**:
- Enhanced print area parser supports all Excel formats
- Print area settings are now preserved in the output file
- Better handling when print area is not set

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

### Test Case 3: Complex File

1. Create an Excel file with:
   - Multiple merged cell ranges
   - Print area set to specific range (e.g., A1:H50)
   - Names in various cells
   - Signature markers in various locations

2. Process through the application

**Expected Result**:
- All merged cells preserved
- Print area preserved
- Signatures correctly placed
- No Excel errors when opening

## Console Logging

The application now provides detailed logging. Open browser console (F12) to see:

- `[parseExcelFile]` - Information about merged cells and print area
- `[인쇄영역파싱]` - Print area parsing details
- `[autoMatch]` - Auto-matching information
- `[병합셀]` - Merged cell information
- `[최종확인]` - Final validation before saving

These logs help debug any issues.

## Known Limitations

1. **ExcelJS Library**: The underlying ExcelJS library has some limitations with very complex Excel features
2. **Large Files**: Very large Excel files (>5MB) may cause memory issues
3. **Complex Formulas**: Some complex formulas might not be preserved exactly

## Troubleshooting

### Issue: Excel file won't open after processing
**Check**:
- Look at browser console for error messages
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

### Files Modified
- `types.ts` - Added merged cells and print area to data types
- `services/excelService.ts` - Core improvements (200+ lines of changes)
- `.gitignore` - Added to prevent committing dependencies
- `CHANGELOG.md` - Documentation of changes

### New Functions
- `isCellInMergedRange` - Check if cell is in merged range
- `isTopLeftOfMergedCell` - Check if cell is top-left of merge
- `parseCellAddress` - Parse cell addresses
- `isValidPrintAreaRange` - Validate print area ranges
- `canPlaceSignature` - Validate signature placement

### Quality Assurance
- ✅ Build successful
- ✅ TypeScript compilation clean
- ✅ Security scan passed (CodeQL)
- ✅ Code review completed

## Reporting Issues

If you encounter problems:

1. Check browser console logs
2. Try with a simpler test file
3. Note the specific file characteristics (merged cells, print area, etc.)
4. Provide console error messages if available

## Next Steps

1. Test with your actual Excel files
2. Report any issues found
3. Verify all edge cases work correctly
4. Consider additional improvements if needed
