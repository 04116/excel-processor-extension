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
    console.log(`🔄 Starting data copy: ${sourceSheetName} → ${targetSheetName}`);
    console.log(`  Column range: ${columnRange.start}-${columnRange.end} (${columnRange.startNum}-${columnRange.endNum})`);

    if (!sourceWorkbook.Sheets[sourceSheetName]) {
      throw new Error(`Source sheet "${sourceSheetName}" not found`);
    }

    let sourceSheet = sourceWorkbook.Sheets[sourceSheetName];
    console.log(`  Source sheet range: ${sourceSheet['!ref']}`);

    if (!targetWorkbook.Sheets[targetSheetName]) {
      throw new Error(`Target sheet "${targetSheetName}" not found in output file. Available sheets: ${targetWorkbook.SheetNames.join(', ')}`);
    }

    let targetSheet = targetWorkbook.Sheets[targetSheetName];
    console.log(`  Target sheet range before clear: ${targetSheet['!ref']}`);

    // Clear ONLY the detected column range, preserving other columns
    targetSheet = clearColumnRange(targetSheet, columnRange.start, columnRange.end);
    console.log(`  Target sheet range after clear: ${targetSheet['!ref']}`);

    // Get source data
    const sourceData = XLSX.utils.sheet_to_json(sourceSheet, { header: 1, defval: '' });
    console.log(`  Source data: ${sourceData.length} rows`);
    if (sourceData.length > 0) {
      console.log(`  First row: [${sourceData[0].slice(0, 5).join(', ')}...]`);
      console.log(`  Sample data in column range:`, sourceData.slice(0, 3).map(row =>
        row.slice(columnRange.startNum, columnRange.endNum + 1)
      ));
    }

    console.log(`Replacing data in columns ${columnRange.start}-${columnRange.end}, preserving other columns`);

    // Copy data to target sheet within detected column range only
    let copiedCells = 0;
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
            copiedCells++;

            // Debug first few cells
            if (copiedCells <= 5) {
              console.log(`  Cell copy: [${rowIndex},${colIndex}] → [${rowIndex},${targetColIndex}] (${cellAddress}) = "${cellValue}"`);
            }
          }
        }
      }
    }
    console.log(`  ✅ Copied ${copiedCells} cells from ${sourceData.length} rows`);

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
function processPowerBIFilesWithSheetMapping(outputWorkbook, downloadedWorkbooks, sheetMapping, originalDownloadResults = null) {
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

        console.log(`  Detected column range for ${fileKey}:`, columnRange);

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
          message: `Successfully replaced columns ${columnRange.start}-${columnRange.end} in sheet ${targetSheetName}`,
          isPlaceholder: originalDownloadResults && originalDownloadResults[fileKey] ? originalDownloadResults[fileKey].isPlaceholder || false : false
        });

        processedSheets.add(targetSheetName);
        console.log(`  ✅ Successfully processed ${fileKey} → ${targetSheetName}`);

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
    let downloadedWorkbooks = {};

    if (downloadResults.workbooks) {
      // Results from downloadAndProcessPowerBIFiles (old format)
      downloadedWorkbooks = downloadResults.workbooks;
    } else if (downloadResults.downloadResults) {
      // New format - need to convert raw data to workbooks
      console.log('Converting downloaded files to workbooks...');
      const XLSX = getXLSX();

      for (const [fileKey, result] of Object.entries(downloadResults.downloadResults)) {
        if (result.success && result.data) {
          try {
            console.log(`Converting ${fileKey} to workbook...`);

            // Convert object back to Uint8Array if needed
            let dataArray = result.data;
            if (result.data.constructor.name === 'Object' && typeof result.data['0'] === 'number') {
              // Data was serialized as object, convert back to Uint8Array
              const dataLength = Object.keys(result.data).length;
              dataArray = new Uint8Array(dataLength);
              for (let i = 0; i < dataLength; i++) {
                dataArray[i] = result.data[i];
              }
              console.log(`  Converted object to Uint8Array, length: ${dataArray.length}`);
              console.log(`  First few bytes:`, Array.from(dataArray.slice(0, 20)).map(b => b.toString(16).padStart(2, '0')).join(' '));
            }

            const workbook = XLSX.read(dataArray, {
              type: 'array',
              cellDates: true,
              cellNF: false,
              cellStyles: false
            });

            downloadedWorkbooks[fileKey] = workbook;
            console.log(`✅ Converted ${fileKey} to workbook with sheets:`, workbook.SheetNames);

            // Debug: Check if workbook has data
            if (workbook.SheetNames.length > 0) {
              const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
              const range = firstSheet['!ref'];
              console.log(`  Sheet range: ${range}`);
              if (range) {
                const dataArray = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
                console.log(`  Data rows: ${dataArray.length}, first few rows:`, dataArray.slice(0, 3));
              }
            }

            // Save individual PowerBI file for inspection
            try {
              const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
              const filename = `powerbi_${fileKey}_${timestamp}.xlsx`;
              downloadExcelFile(workbook, filename);
              console.log(`📁 Saved PowerBI file: ${filename}`);
            } catch (saveError) {
              console.warn(`Could not save PowerBI file ${fileKey}:`, saveError);
            }
          } catch (error) {
            console.error(`❌ Failed to convert ${fileKey} to workbook:`, error);
            // Continue with other files even if one fails
          }
        }
      }
    }

    // Validate sheet mapping
    const validation = validateSheetMappingInWorkbook(outputWorkbook, sheetMapping);
    console.log('Sheet mapping validation:', validation);

    // Process files with sheet mapping
    const processingResults = processPowerBIFilesWithSheetMapping(
      outputWorkbook,
      downloadedWorkbooks,
      sheetMapping,
      downloadResults.downloadResults
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

// Create a new Excel workbook from downloaded PowerBI files
async function createNewExcelFromPowerBIFiles(downloadResults) {
  try {
    const XLSX = getXLSX();

    // Create a new workbook
    const newWorkbook = XLSX.utils.book_new();
    const results = [];
    const createdSheets = [];

    // Convert download results to workbooks if needed
    let downloadedWorkbooks = {};

    if (downloadResults.workbooks) {
      // Results from downloadAndProcessPowerBIFiles (old format)
      downloadedWorkbooks = downloadResults.workbooks;
    } else if (downloadResults.downloadResults) {
      // New format - need to convert raw data to workbooks
      console.log('Converting downloaded files to workbooks...');

      for (const [fileKey, result] of Object.entries(downloadResults.downloadResults)) {
        if (result.success && result.data) {
          try {
            console.log(`Converting ${fileKey} to workbook...`);

            // Convert object back to Uint8Array if needed
            let dataArray = result.data;
            if (result.data.constructor.name === 'Object' && typeof result.data['0'] === 'number') {
              // Data was serialized as object, convert back to Uint8Array
              const dataLength = Object.keys(result.data).length;
              dataArray = new Uint8Array(dataLength);
              for (let i = 0; i < dataLength; i++) {
                dataArray[i] = result.data[i];
              }
              console.log(`  Converted object to Uint8Array, length: ${dataArray.length}`);
            }

            const workbook = XLSX.read(dataArray, {
              type: 'array',
              cellDates: true,
              cellNF: false,
              cellStyles: false
            });

            downloadedWorkbooks[fileKey] = workbook;
            console.log(`✅ Converted ${fileKey} to workbook with sheets:`, workbook.SheetNames);

          } catch (error) {
            console.error(`❌ Failed to convert ${fileKey} to workbook:`, error);

            // Create placeholder data for failed conversions
            const placeholderData = [
              ['Error', 'Failed to process downloaded data'],
              ['File Key', fileKey],
              ['Error Message', error.message],
              ['Download Time', new Date().toISOString()],
              ['', ''],
              ['Note', 'This is placeholder data due to processing error']
            ];

            const placeholderWorkbook = XLSX.utils.book_new();
            const placeholderSheet = XLSX.utils.aoa_to_sheet(placeholderData);
            XLSX.utils.book_append_sheet(placeholderWorkbook, placeholderSheet, 'Error_Data');

            downloadedWorkbooks[fileKey] = placeholderWorkbook;

            results.push({
              fileKey,
              sheetName: fileKey,
              success: false,
              error: error.message,
              isPlaceholder: true
            });
          }
        } else {
          console.warn(`No data for ${fileKey}, creating placeholder`);

          // Create placeholder for missing data
          const placeholderData = [
            ['Status', 'Download Failed'],
            ['File Key', fileKey],
            ['Error', result.error || 'No data available'],
            ['Download Time', new Date().toISOString()],
            ['', ''],
            ['Note', 'This is placeholder data due to download failure']
          ];

          const placeholderWorkbook = XLSX.utils.book_new();
          const placeholderSheet = XLSX.utils.aoa_to_sheet(placeholderData);
          XLSX.utils.book_append_sheet(placeholderWorkbook, placeholderSheet, 'Missing_Data');

          downloadedWorkbooks[fileKey] = placeholderWorkbook;

          results.push({
            fileKey,
            sheetName: fileKey,
            success: false,
            error: result.error || 'Download failed',
            isPlaceholder: true
          });
        }
      }
    }

    // Add each downloaded PowerBI file as a separate sheet
    for (const [fileKey, workbook] of Object.entries(downloadedWorkbooks)) {
      try {
        if (workbook && workbook.SheetNames && workbook.SheetNames.length > 0) {
          const sourceSheetName = workbook.SheetNames[0]; // Use first sheet from downloaded file
          const sourceSheet = workbook.Sheets[sourceSheetName];

          // Create a clean sheet name (remove invalid characters)
          let targetSheetName = fileKey.replace(/[\/\\\?\*\[\]]/g, '_');

          // Ensure sheet name is not too long (Excel limit is 31 characters)
          if (targetSheetName.length > 31) {
            targetSheetName = targetSheetName.substring(0, 31);
          }

          // Ensure unique sheet name
          let finalSheetName = targetSheetName;
          let counter = 1;
          while (newWorkbook.Sheets[finalSheetName]) {
            const suffix = `_${counter}`;
            const maxLength = 31 - suffix.length;
            finalSheetName = targetSheetName.substring(0, maxLength) + suffix;
            counter++;
          }

          console.log(`Adding sheet: ${fileKey} → ${finalSheetName}`);

          // Clone the sheet to avoid reference issues
          let clonedSheet = JSON.parse(JSON.stringify(sourceSheet));

          // Add the sheet to the new workbook
          XLSX.utils.book_append_sheet(newWorkbook, clonedSheet, finalSheetName);

          createdSheets.push(finalSheetName);

          // Check if this was a placeholder (error case)
          const isPlaceholder = results.some(r => r.fileKey === fileKey && r.isPlaceholder);

          if (!isPlaceholder) {
            results.push({
              fileKey,
              sheetName: finalSheetName,
              success: true,
              message: `Successfully added sheet ${finalSheetName} from ${fileKey}`,
              isPlaceholder: false
            });
          } else {
            // Update the existing placeholder result with sheet name
            const placeholderResult = results.find(r => r.fileKey === fileKey && r.isPlaceholder);
            if (placeholderResult) {
              placeholderResult.sheetName = finalSheetName;
            }
          }

          console.log(`✅ Successfully added sheet ${finalSheetName} from ${fileKey}`);

        } else {
          throw new Error(`No sheets found in downloaded file ${fileKey}`);
        }

      } catch (error) {
        console.error(`Failed to add sheet for ${fileKey}:`, error);

        // Create error sheet for this file
        const errorData = [
          ['Error Processing File', fileKey],
          ['Error Message', error.message],
          ['Timestamp', new Date().toISOString()],
          ['', ''],
          ['Note', 'This sheet represents a failed file processing']
        ];

        const errorSheet = XLSX.utils.aoa_to_sheet(errorData);
        let errorSheetName = `Error_${fileKey}`.substring(0, 31);

        // Ensure unique error sheet name
        let counter = 1;
        while (newWorkbook.Sheets[errorSheetName]) {
          const suffix = `_${counter}`;
          const maxLength = 31 - suffix.length;
          errorSheetName = `Error_${fileKey}`.substring(0, maxLength) + suffix;
          counter++;
        }

        XLSX.utils.book_append_sheet(newWorkbook, errorSheet, errorSheetName);
        createdSheets.push(errorSheetName);

        results.push({
          fileKey,
          sheetName: errorSheetName,
          success: false,
          error: error.message,
          isPlaceholder: true
        });
      }
    }

    // Add a summary sheet
    const summaryData = [
      ['PowerBI Data Export Summary'],
      ['Generated on', new Date().toISOString()],
      ['Total Files', Object.keys(downloadedWorkbooks).length],
      ['Successful Downloads', results.filter(r => r.success && !r.isPlaceholder).length],
      ['Failed Downloads', results.filter(r => !r.success || r.isPlaceholder).length],
      [''],
      ['Sheet Details:'],
      ['Sheet Name', 'Source File', 'Status'],
      ...results.map(r => [
        r.sheetName,
        r.fileKey,
        r.success && !r.isPlaceholder ? 'Success' : (r.isPlaceholder ? 'Placeholder' : 'Error')
      ])
    ];

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(newWorkbook, summarySheet, 'Summary');
    createdSheets.unshift('Summary'); // Add to beginning

    console.log(`Created new Excel workbook with ${createdSheets.length} sheets:`, createdSheets);

    return {
      success: true,
      workbook: newWorkbook,
      results: results,
      createdSheets: createdSheets
    };

  } catch (error) {
    console.error('Error creating new Excel file from PowerBI data:', error);
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
  clearColumnRange,
  copyDataToSheetSpecific,
  downloadExcelFile,
  validateRequiredSheets,
  readFileAsArrayBuffer,
  detectColumnsFromBIFiles,
  replaceMatchingColumns,
  processPowerBIFilesWithSheetMapping,
  validateSheetMappingInWorkbook,
  processDownloadedPowerBIFiles,
  createNewExcelFromPowerBIFiles
};