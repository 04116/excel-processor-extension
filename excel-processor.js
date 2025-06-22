// Excel Processing Module
// Handles Excel file reading, writing, and data manipulation operations
// Preserves original file format (xlsb stays xlsb, xlsx stays xlsx)

// Get reference to XLSX library
function getXLSX() {
  return window.XLSX;
}

// Convert column number to letter (0 = A, 1 = B, etc.)
function numberToColumn(num) {
  let result = '';
  while (num >= 0) {
    result = String.fromCharCode((num % 26) + 'A'.charCodeAt(0)) + result;
    num = Math.floor(num / 26) - 1;
  }
  return result;
}

// Convert column letter to number (A = 0, B = 1, etc.)
function columnToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return result - 1;
}

// Read Excel file buffer and return workbook
function readExcelFile(buffer, filename, requiredSheets = null) {
  const XLSX = getXLSX();
  const extension = filename.toLowerCase().split('.').pop();

  try {
    // Use SheetJS library with options optimized for different file types
    const options = {
      type: 'array',
      cellDates: true,
      cellNF: false,
      cellStyles: false
    };

    // For output files, only read required sheets to optimize performance
    if (requiredSheets && Array.isArray(requiredSheets)) {
      options.sheets = requiredSheets;
    }

    return XLSX.read(buffer, options);
  } catch (error) {
    console.warn(`Failed to read ${extension} file with array buffer, trying fallback:`, error);
    try {
      // Try with minimal options
      return XLSX.read(buffer, { type: 'array' });
    } catch (fallbackError) {
      console.warn('Array buffer read failed, trying binary string:', fallbackError);
      // Convert to binary string as last resort
      const binaryString = Array.from(buffer, byte => String.fromCharCode(byte)).join('');
      return XLSX.read(binaryString, { type: 'binary' });
    }
  }
}

// Detect column range with data in a workbook (first sheet)
function detectColumnRange(workbook) {
  const XLSX = getXLSX();

  try {
    const sheetNames = workbook.SheetNames;
    if (sheetNames.length === 0) return null;

    const sheetName = sheetNames[0];
    let worksheet = workbook.Sheets[sheetName];

    // Unmerge cells first to get accurate data range
    worksheet = unmergeSheet(worksheet);

    if (!worksheet['!ref']) return null;

    // Convert to array data for analysis
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    let minCol = null;
    let maxCol = null;

    // Find the actual data range by scanning all non-empty cells
    for (let rowIndex = 0; rowIndex < jsonData.length; rowIndex++) {
      const row = jsonData[rowIndex];
      for (let colIndex = 0; colIndex < row.length; colIndex++) {
        const cellValue = row[colIndex];
        if (cellValue !== undefined && cellValue !== null && cellValue.toString().trim() !== '') {
          if (minCol === null || colIndex < minCol) minCol = colIndex;
          if (maxCol === null || colIndex > maxCol) maxCol = colIndex;
        }
      }
    }

    if (minCol !== null && maxCol !== null) {
      return {
        start: numberToColumn(minCol),
        end: numberToColumn(maxCol),
        startNum: minCol,
        endNum: maxCol
      };
    }

    return null;
  } catch (error) {
    console.error('Error detecting column range:', error);
    return null;
  }
}

// Unmerge all merged cells in a worksheet
function unmergeSheet(worksheet) {
  const XLSX = getXLSX();

  if (!worksheet['!merges']) return worksheet;

  const merges = worksheet['!merges'];
  const newWorksheet = { ...worksheet };

  // For each merged range, copy the top-left value to all cells in the range
  merges.forEach(merge => {
    const startRow = merge.s.r;
    const endRow = merge.e.r;
    const startCol = merge.s.c;
    const endCol = merge.e.c;

    const topLeftCell = XLSX.utils.encode_cell({ r: startRow, c: startCol });
    const topLeftValue = worksheet[topLeftCell];

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (topLeftValue) {
          newWorksheet[cellAddress] = { ...topLeftValue };
        }
      }
    }
  });

  // Remove merge information since we've unmerged everything
  delete newWorksheet['!merges'];
  return newWorksheet;
}

// Clear specified column range in a worksheet (only the detected columns)
function clearColumnRange(worksheet, startCol, endCol) {
  const XLSX = getXLSX();

  if (!worksheet['!ref']) return worksheet;

  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const startColNum = columnToNumber(startCol.toUpperCase());
  const endColNum = columnToNumber(endCol.toUpperCase());

  console.log(`Clearing columns ${startCol} to ${endCol} (${startColNum} to ${endColNum}) in worksheet`);

  // Only delete cells in the specified column range, preserving other columns
  for (let row = range.s.r; row <= range.e.r; row++) {
    for (let col = startColNum; col <= endColNum; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
      delete worksheet[cellAddress];
    }
  }

  return worksheet;
}

// Copy data from specific source sheet to specific target sheet (only detected columns)
function copyDataToSheetSpecific(sourceWorkbook, targetWorkbook, sourceSheetName, targetSheetName, columnRange) {
  const XLSX = getXLSX();

  try {
    if (!sourceWorkbook.Sheets[sourceSheetName]) {
      throw new Error(`Source sheet "${sourceSheetName}" not found`);
    }

    let sourceSheet = sourceWorkbook.Sheets[sourceSheetName];
    // Unmerge source sheet
    sourceSheet = unmergeSheet(sourceSheet);

    if (!targetWorkbook.Sheets[targetSheetName]) {
      throw new Error(`Target sheet "${targetSheetName}" not found in output file. Available sheets: ${targetWorkbook.SheetNames.join(', ')}`);
    }

    let targetSheet = targetWorkbook.Sheets[targetSheetName];
    // Clear ONLY the detected column range, preserving other columns
    targetSheet = clearColumnRange(targetSheet, columnRange.start, columnRange.end);

    // Get source data
    const sourceData = XLSX.utils.sheet_to_json(sourceSheet, { header: 1, defval: '' });

    console.log(`Replacing data in columns ${columnRange.start}-${columnRange.end}, preserving other columns`);

    // Copy data to target sheet within detected column range only
    for (let rowIndex = 0; rowIndex < sourceData.length; rowIndex++) {
      const row = sourceData[rowIndex];
      for (let colIndex = 0; colIndex < row.length; colIndex++) {
        // Only copy within the detected column range
        if (colIndex <= (columnRange.endNum - columnRange.startNum)) {
          const cellValue = row[colIndex];
          if (cellValue !== undefined && cellValue !== null && cellValue.toString().trim() !== '') {
            const targetColIndex = columnRange.startNum + colIndex;
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: targetColIndex });
            targetSheet[cellAddress] = {
              v: cellValue,
              t: typeof cellValue === 'number' ? 'n' : 's'
            };
          }
        }
      }
    }

    // Preserve existing sheet range, only extend if necessary
    const existingRange = targetSheet['!ref'] ?
      XLSX.utils.decode_range(targetSheet['!ref']) :
      { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };

    if (sourceData.length > 0) {
      // Only extend range if new data goes beyond existing bounds
      existingRange.e.r = Math.max(existingRange.e.r, sourceData.length - 1);
      existingRange.e.c = Math.max(existingRange.e.c, columnRange.endNum);
      targetSheet['!ref'] = XLSX.utils.encode_range(existingRange);
    }

    targetWorkbook.Sheets[targetSheetName] = targetSheet;
    return targetWorkbook;
  } catch (error) {
    console.error('Error copying data:', error);
    throw error;
  }
}

// Download Excel file preserving original format
function downloadExcelFile(workbook, filename) {
  const XLSX = getXLSX();

  // Determine file type from extension to preserve original format
  let bookType = 'xlsx'; // default

  if (filename.toLowerCase().endsWith('.xlsb')) {
    bookType = 'xlsb';
  } else if (filename.toLowerCase().endsWith('.xls')) {
    bookType = 'xls';
  }

  console.log(`Downloading file as ${bookType} format`);

  // Write workbook in the same format as input
  const wbout = XLSX.write(workbook, {
    bookType: bookType,
    type: 'array'
  });

  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);

  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// Validate required sheets exist in workbook
function validateRequiredSheets(workbook, requiredSheetNames) {
  const availableSheets = workbook.SheetNames;
  const missingSheets = [];

  requiredSheetNames.forEach(sheetName => {
    if (!availableSheets.includes(sheetName)) {
      missingSheets.push(sheetName);
    }
  });

  return {
    isValid: missingSheets.length === 0,
    missingSheets: missingSheets,
    availableSheets: availableSheets,
    requiredSheets: requiredSheetNames
  };
}

// Read file as array buffer
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function (e) {
      resolve(new Uint8Array(e.target.result));
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// Detect columns from downloaded input files, these will be the ONLY columns replaced in output
function detectColumnsFromBIFiles(biWorkbooks) {
  const detectedColumns = {};

  for (const [sheetName, workbook] of Object.entries(biWorkbooks)) {
    const columnRange = detectColumnRange(workbook);
    if (columnRange) {
      detectedColumns[sheetName] = columnRange;
      console.log(`Detected columns in input file ${sheetName}: ${columnRange.start}-${columnRange.end} (will replace ONLY these columns)`);
    } else {
      console.warn(`No data columns detected in input file ${sheetName}`);
    }
  }

  return detectedColumns;
}

// Replace ONLY detected columns in output file, preserving other columns
function replaceMatchingColumns(outputWorkbook, biWorkbooks, detectedColumns) {
  const results = [];

  console.log(`Replacing columns in ${Object.keys(detectedColumns).length} sheets, preserving other columns`);

  for (const [sheetName, columnRange] of Object.entries(detectedColumns)) {
    if (biWorkbooks[sheetName] && outputWorkbook.Sheets[sheetName]) {
      try {
        console.log(`Processing sheet ${sheetName}: replacing columns ${columnRange.start}-${columnRange.end}`);

        copyDataToSheetSpecific(
          biWorkbooks[sheetName],
          outputWorkbook,
          biWorkbooks[sheetName].SheetNames[0], // First sheet from input file
          sheetName, // Target sheet in output file
          columnRange // Only replace these specific columns
        );

        results.push({
          sheetName,
          success: true,
          columnRange: `${columnRange.start}-${columnRange.end}`,
          message: `Replaced columns ${columnRange.start}-${columnRange.end}, other columns preserved`
        });
      } catch (error) {
        console.error(`Failed to replace columns in ${sheetName}:`, error);
        results.push({
          sheetName,
          success: false,
          error: error.message
        });
      }
    } else {
      console.warn(`Skipping sheet ${sheetName}: missing in input or output file`);
      results.push({
        sheetName,
        success: false,
        error: 'Sheet not found in input or output file'
      });
    }
  }

  return results;
}

// Process PowerBI downloaded files and replace sheets in destination Excel file
function processPowerBIFilesWithSheetMapping(outputWorkbook, downloadedWorkbooks, sheetMapping) {
  const results = [];
  const processedSheets = new Set();

  console.log('Processing PowerBI files with sheet mapping:', sheetMapping);

  // Validate that all target sheets exist
  const sheetValidation = validateSheetMappingInWorkbook(outputWorkbook, sheetMapping);
  const invalidMappings = Object.entries(sheetValidation).filter(([_, validation]) => !validation.exists);
  
  if (invalidMappings.length > 0) {
    const missingSheets = invalidMappings.map(([fileKey, validation]) => 
      `${fileKey} → ${validation.targetSheet}`
    ).join(', ');
    
    throw new Error(`Target sheets not found in destination file: ${missingSheets}. Available sheets: ${outputWorkbook.SheetNames.join(', ')}`);
  }

  // Process each downloaded file according to the sheet mapping
  for (const [fileKey, targetSheetName] of Object.entries(sheetMapping)) {
    if (downloadedWorkbooks[fileKey]) {
      try {
        const sourceWorkbook = downloadedWorkbooks[fileKey];
        const sourceSheetName = sourceWorkbook.SheetNames[0]; // Use first sheet from downloaded file
        
        console.log(`Processing ${fileKey}: ${sourceSheetName} → ${targetSheetName}`);

        // Detect column range in the source file
        const columnRange = detectColumnRange(sourceWorkbook);
        if (!columnRange) {
          throw new Error(`No data columns detected in ${fileKey}`);
        }

        // Copy data to the target sheet, replacing only the detected columns
        copyDataToSheetSpecific(
          sourceWorkbook,
          outputWorkbook,
          sourceSheetName,
          targetSheetName,
          columnRange
        );

        results.push({
          fileKey,
          sourceSheet: sourceSheetName,
          targetSheet: targetSheetName,
          success: true,
          columnRange: `${columnRange.start}-${columnRange.end}`,
          message: `Successfully replaced columns ${columnRange.start}-${columnRange.end} in sheet ${targetSheetName}`
        });

        processedSheets.add(targetSheetName);

      } catch (error) {
        console.error(`Failed to process ${fileKey}:`, error);
        results.push({
          fileKey,
          targetSheet: sheetMapping[fileKey],
          success: false,
          error: error.message
        });
      }
    } else {
      console.warn(`Downloaded file ${fileKey} not found`);
      results.push({
        fileKey,
        targetSheet: sheetMapping[fileKey],
        success: false,
        error: `Downloaded file ${fileKey} not available`
      });
    }
  }

  console.log(`Processing completed. ${processedSheets.size} sheets updated:`, Array.from(processedSheets));
  return results;
}

// Validate sheet mapping against destination workbook
function validateSheetMappingInWorkbook(outputWorkbook, sheetMapping) {
  const availableSheets = outputWorkbook.SheetNames;
  const validation = {};
  
  Object.entries(sheetMapping).forEach(([fileKey, targetSheet]) => {
    validation[fileKey] = {
      targetSheet: targetSheet,
      exists: availableSheets.includes(targetSheet),
      available: availableSheets
    };
  });
  
  return validation;
}

// Complete processing workflow for PowerBI integration
async function processDownloadedPowerBIFiles(outputWorkbook, downloadResults, sheetMapping) {
  try {
    // Convert download results to workbooks if needed
    let downloadedWorkbooks = downloadResults;
    
    if (downloadResults.workbooks) {
      // Results from downloadAndProcessPowerBIFiles
      downloadedWorkbooks = downloadResults.workbooks;
    }

    // Validate sheet mapping
    const validation = validateSheetMappingInWorkbook(outputWorkbook, sheetMapping);
    console.log('Sheet mapping validation:', validation);

    // Process files with sheet mapping
    const processingResults = processPowerBIFilesWithSheetMapping(
      outputWorkbook, 
      downloadedWorkbooks, 
      sheetMapping
    );

    return {
      success: true,
      results: processingResults,
      validation: validation,
      processedSheets: processingResults.filter(r => r.success).map(r => r.targetSheet)
    };

  } catch (error) {
    console.error('Error processing PowerBI files:', error);
    return {
      success: false,
      error: error.message,
      results: []
    };
  }
}

// Export functions for use in other modules
window.ExcelProcessor = {
  numberToColumn,
  columnToNumber,
  readExcelFile,
  detectColumnRange,
  unmergeSheet,
  clearColumnRange,
  copyDataToSheetSpecific,
  downloadExcelFile,
  validateRequiredSheets,
  readFileAsArrayBuffer,
  detectColumnsFromBIFiles,
  replaceMatchingColumns,
  processPowerBIFilesWithSheetMapping,
  validateSheetMappingInWorkbook,
  processDownloadedPowerBIFiles
};