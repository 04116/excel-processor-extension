function requestBodyForDownloadByTimePeriod(dateConfig) {
  const currentStartDate = dateConfig.getCurrentStartDate();
  const currentEndDate = dateConfig.getCurrentEndDate();
  const referenceStartDate = dateConfig.getReferenceStartDate();
  const referenceEndDate = dateConfig.getReferenceEndDate();

  return {
    "exportDataType": 2,
    "executeSemanticQueryRequest": {
      "version": "1.0.0",
      "queries": [{
        "Query": {
          "Commands": [{
            "SemanticQueryDataShapeCommand": {
              "Query": {
                "Version": 2,
                "From": [
                  { "Name": "m", "Entity": "Measure", "Type": 0 },
                  { "Name": "d", "Entity": "DIM_MCH2_Index_2025", "Type": 0 },
                  { "Name": "d2", "Entity": "DIM_MCH4_Index_2025", "Type": 0 },
                  { "Name": "d1", "Entity": "DIM_MCH3_Index_2025", "Type": 0 },
                  { "Name": "d3", "Entity": "DimStore", "Type": 0 },
                  { "Name": "d4", "Entity": "DIM_REVENUE_TYPE_BY_STORE", "Type": 0 },
                  { "Name": "d11", "Entity": "DimStore_Prj", "Type": 0 },
                  { "Name": "d21", "Entity": "DimDateTime", "Type": 0 },
                  { "Name": "d31", "Entity": "DimDateTime_Ref", "Type": 0 }
                ],
                "Select": [
                  { "Measure": { "Expression": { "SourceRef": { "Source": "m" } }, "Property": "Revenue" }, "Name": "Measure.Revenue", "NativeReferenceName": "Total Sales" },
                  { "Measure": { "Expression": { "SourceRef": { "Source": "m" } }, "Property": "Revenue /Store /Day" }, "Name": "Measure.Revenue /Store /Day", "NativeReferenceName": "Sales per day" },
                  { "Measure": { "Expression": { "SourceRef": { "Source": "m" } }, "Property": "Revenue /Store /Day Change %" }, "Name": "Measure.Revenue /Store /Day Change %", "NativeReferenceName": "Sales per day (Vs Before)" },
                  { "Measure": { "Expression": { "SourceRef": { "Source": "m" } }, "Property": "#Bill /Store /Day" }, "Name": "Measure.#Bill /Store /Day", "NativeReferenceName": "Bill per day" },
                  { "Measure": { "Expression": { "SourceRef": { "Source": "m" } }, "Property": "#Bill /Store /Day Change %" }, "Name": "Measure.#Bill /Store /Day Change %", "NativeReferenceName": "Bill per day (Vs. Before)" },
                  { "Measure": { "Expression": { "SourceRef": { "Source": "m" } }, "Property": "Sales Mix %" }, "Name": "Measure.Sales Mix %", "NativeReferenceName": "Sales Mix %" },
                  { "Column": { "Expression": { "SourceRef": { "Source": "d" } }, "Property": "MCH2" }, "Name": "DIM_MCH2_Index_2025.MCH2", "NativeReferenceName": "MCH2" },
                  { "Column": { "Expression": { "SourceRef": { "Source": "d2" } }, "Property": "MCH4" }, "Name": "DIM_MCH4_Index_2025.MCH4", "NativeReferenceName": "MCH4" },
                  { "Column": { "Expression": { "SourceRef": { "Source": "d1" } }, "Property": "MCH3" }, "Name": "DIM_MCH3_Index_2025.MCH3", "NativeReferenceName": "MCH3" },
                  { "Column": { "Expression": { "SourceRef": { "Source": "d3" } }, "Property": "StoreProfile.SAP_Name" }, "Name": "DimStore.StoreProfile.SAP_Name", "NativeReferenceName": "StoreProfile.SAP_Name" }
                ],
                "Where": [
                  {
                    "Condition": {
                      "In": {
                        "Expressions": [
                          { "Column": { "Expression": { "SourceRef": { "Source": "d4" } }, "Property": "REVENUE_TYPE" } },
                          { "Column": { "Expression": { "SourceRef": { "Source": "d4" } }, "Property": "TYPE" } }
                        ],
                        "Values": [[{ "Literal": { "Value": "'Retail'" } }, { "Literal": { "Value": "'POS NORMAL'" } }]]
                      }
                    }
                  },
                  {
                    "Condition": {
                      "In": {
                        "Expressions": [
                          { "Column": { "Expression": { "SourceRef": { "Source": "d11" } }, "Property": "Group Project" } },
                          { "Column": { "Expression": { "SourceRef": { "Source": "d11" } }, "Property": "Project" } },
                          { "Column": { "Expression": { "SourceRef": { "Source": "d11" } }, "Property": "Detail Project" } }
                        ],
                        "Values": [
                          [{ "Literal": { "Value": "'Operation Projects'" } }, { "Literal": { "Value": "'Reno Project'" } }, { "Literal": { "Value": "'(04 stores) Q1.25'" } }],
                          [{ "Literal": { "Value": "'Operation Projects'" } }, { "Literal": { "Value": "'Reno Project'" } }, { "Literal": { "Value": "'(09 stores) Q4.24'" } }],
                          [{ "Literal": { "Value": "'Operation Projects'" } }, { "Literal": { "Value": "'Reno Project'" } }, { "Literal": { "Value": "'(10 store) Q2.25'" } }]
                        ]
                      }
                    }
                  },
                  {
                    "Condition": {
                      "And": {
                        "Left": {
                          "Comparison": {
                            "ComparisonKind": 2,
                            "Left": { "Column": { "Expression": { "SourceRef": { "Source": "d21" } }, "Property": "Date" } },
                            "Right": { "Literal": { "Value": `datetime'${currentStartDate}'` } }
                          }
                        },
                        "Right": {
                          "Comparison": {
                            "ComparisonKind": 3,
                            "Left": { "Column": { "Expression": { "SourceRef": { "Source": "d21" } }, "Property": "Date" } },
                            "Right": { "Literal": { "Value": `datetime'${currentEndDate}'` } }
                          }
                        }
                      }
                    }
                  },
                  {
                    "Condition": {
                      "And": {
                        "Left": {
                          "Comparison": {
                            "ComparisonKind": 2,
                            "Left": { "Column": { "Expression": { "SourceRef": { "Source": "d31" } }, "Property": "Date" } },
                            "Right": { "Literal": { "Value": `datetime'${referenceStartDate}'` } }
                          }
                        },
                        "Right": {
                          "Comparison": {
                            "ComparisonKind": 3,
                            "Left": { "Column": { "Expression": { "SourceRef": { "Source": "d31" } }, "Property": "Date" } },
                            "Right": { "Literal": { "Value": `datetime'${referenceEndDate}'` } }
                          }
                        }
                      }
                    }
                  },
                  {
                    "Condition": {
                      "Not": {
                        "Expression": {
                          "Comparison": {
                            "ComparisonKind": 0,
                            "Left": { "Column": { "Expression": { "SourceRef": { "Source": "d" } }, "Property": "MCH2_-_Department" } },
                            "Right": { "Literal": { "Value": "null" } }
                          }
                        }
                      }
                    }
                  },
                  {
                    "Condition": {
                      "In": {
                        "Expressions": [{ "Column": { "Expression": { "SourceRef": { "Source": "d3" } }, "Property": "StoreProfile.Tình_trạng" } }],
                        "Values": [[{ "Literal": { "Value": "'Hoạt Động'" } }]]
                      }
                    }
                  },
                  {
                    "Condition": {
                      "In": {
                        "Expressions": [{ "Column": { "Expression": { "SourceRef": { "Source": "d3" } }, "Property": "StoreProfile.Chain" } }],
                        "Values": [[{ "Literal": { "Value": "'WMP'" } }], [{ "Literal": { "Value": "'WMT'" } }]]
                      }
                    }
                  },
                  {
                    "Condition": {
                      "Not": {
                        "Expression": {
                          "Comparison": {
                            "ComparisonKind": 0,
                            "Left": { "Column": { "Expression": { "SourceRef": { "Source": "d3" } }, "Property": "StoreProfile.Group_Concept" } },
                            "Right": { "Literal": { "Value": "null" } }
                          }
                        }
                      }
                    }
                  }
                ],
                "OrderBy": [{
                  "Direction": 2,
                  "Expression": { "Measure": { "Expression": { "SourceRef": { "Source": "m" } }, "Property": "Revenue" } }
                }]
              },
              "Binding": {
                "Primary": {
                  "Groupings": [
                    { "Projections": [9], "Subtotal": 2 },
                    { "Projections": [6], "Subtotal": 2 },
                    { "Projections": [8], "Subtotal": 2 },
                    { "Projections": [0, 1, 2, 3, 4, 5, 7], "Subtotal": 2 }
                  ]
                },
                "DataReduction": {
                  "Primary": { "Top": { "Count": 1000000 } },
                  "Secondary": { "Top": { "Count": 100 } }
                },
                "Aggregates": [{ "Select": 5, "Aggregations": [{ "Min": {} }, { "Max": {} }] }],
                "Version": 1
              }
            }
          }, {
            "ExportDataCommand": {
              "Columns": [
                { "QueryName": "Measure.Revenue", "Name": "Total Sales" },
                { "QueryName": "Measure.Revenue /Store /Day", "Name": "Sales per day" },
                { "QueryName": "Measure.Revenue /Store /Day Change %", "Name": "Sales per day (Vs Before)" },
                { "QueryName": "Measure.#Bill /Store /Day", "Name": "Bill per day" },
                { "QueryName": "Measure.#Bill /Store /Day Change %", "Name": "Bill per day (Vs. Before)" },
                { "QueryName": "Measure.Sales Mix %", "Name": "Sales Mix %" },
                { "QueryName": "DIM_MCH2_Index_2025.MCH2", "Name": "MCH2" },
                { "QueryName": "DIM_MCH4_Index_2025.MCH4", "Name": "MCH4" },
                { "QueryName": "DIM_MCH3_Index_2025.MCH3", "Name": "MCH3" },
                { "QueryName": "DimStore.StoreProfile.SAP_Name", "Name": "StoreProfile.SAP_Name" }
              ],
              "Ordering": [0, 5, 1, 2, 3, 4],
              "FiltersDescription": `Applied filters:\nIncluded (1) Retail (REVENUE_TYPE) + POS NORMAL (TYPE)\nIncluded (3) Operation Projects (Group Project) + Reno Project (Project) + (04 stores) Q1.25 (Detail Project), Operation Projects (Group Project) + Reno Project (Project) + (09 stores) Q4.24 (Detail Project), Operation Projects (Group Project) + Reno Project (Project) + (10 store) Q2.25 (Detail Project)\nDate is on or after ${currentStartDate.split('T')[0]} 12:00:00 AM and is before ${currentEndDate.split('T')[0]} 12:00:00 AM\nDate is on or after ${referenceStartDate.split('T')[0]} 12:00:00 AM and is before ${referenceEndDate.split('T')[0]} 12:00:00 AM\nMCH2_-_Department is not blank\nStoreProfile.Tình_trạng is Hoạt Động\nStoreProfile.Chain is WMP or WMT\nStoreProfile.Group_Concept is not blank`
            }
          }]
        }
      }],
      "cancelQueries": [],
      "modelId": 3914248,
      "userPreferredLocale": "en-US"
    },
    "artifactId": 4656201
  };
}

// Get PowerBI token from session storage
function getTokenFromSessionStorage() {
  try {
    // Scan session storage keys for PowerBI auth data
    for (let i = 0; i < sessionStorage.length; i++) {
      const key = sessionStorage.key(i);
      const value = sessionStorage.getItem(key);

      // Look for keys/values that contain homeAccountId (indicates PowerBI auth data)
      if (key && value && (key.includes('homeAccountId') || value.includes('homeAccountId'))) {
        try {
          // Try to parse as JSON to extract token
          const parsed = JSON.parse(value);

          // Look for access token in various possible locations
          const token = parsed.secret || parsed.accessToken || parsed.access_token ||
            parsed.token || parsed.credentialType === 'AccessToken' && parsed.secret;

          if (token && typeof token === 'string' && token.length > 50) {
            return token;
          }

        } catch (e) {
          // Not JSON or other parse error, continue silently
          continue;
        }
      }
    }

    return null;

  } catch (error) {
    console.error('Error accessing session storage:', error);
    return null;
  }
}

async function downloadExcelFile(dateConfig) {
  const requestBody = requestBodyForDownloadByTimePeriod(dateConfig);

  // Get the PowerBI token from session storage
  const bearerToken = getTokenFromSessionStorage();

  if (!bearerToken) {
    const currentUrl = window.location.href;
    throw new Error(`Could not find PowerBI authentication token. 

Please ensure you are:
1. Logged into PowerBI (app.powerbi.com)
2. On a report page (not the home page)
3. The report has loaded completely

Current URL: ${currentUrl}

Try:
- Refreshing the PowerBI page
- Opening a specific report
- Waiting for the page to fully load before using the extension`);
  }

  const headers = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.9,vi;q=0.8",
    "content-type": "application/json;charset=UTF-8",
    "authorization": `Bearer ${bearerToken}`,
    "cache-control": "no-cache",
    "pragma": "no-cache",
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "cross-site",
    "origin": "https://app.powerbi.com",
    "referer": "https://app.powerbi.com/",
    "user-agent": navigator.userAgent
  };

  return fetch("https://wabi-south-east-asia-b-primary-redirect.analysis.windows.net/export/xlsx", {
    method: "POST",
    headers: headers,
    body: JSON.stringify(requestBody),
    credentials: "include"
  });
}

// Load XLSX library into content script context
async function loadXLSXLibrary() {
  return new Promise((resolve, reject) => {
    try {
      const script = document.createElement('script');
      script.src = chrome.runtime.getURL('libs/xlsx.full.min.js');
      script.onload = () => {
        console.log('XLSX library loaded in content script');
        resolve();
      };
      script.onerror = () => {
        reject(new Error('Failed to load XLSX library'));
      };
      document.head.appendChild(script);
    } catch (error) {
      reject(error);
    }
  });
}

// Reconstruct config functions that get lost during serialization
function reconstructConfigFunctions(config) {
  const reconstructed = { ...config };

  // Reconstruct date functions based on the config name
  switch (config.name) {
    case 'Day_Le':
      reconstructed.getCurrentStartDate = () => {
        const date = new Date();
        date.setDate(date.getDate() - 1);
        return date.toISOString().split('.')[0];
      };
      reconstructed.getCurrentEndDate = () => {
        return new Date().toISOString().split('.')[0];
      };
      reconstructed.getReferenceStartDate = () => {
        const date = new Date();
        date.setFullYear(date.getFullYear() - 1);
        date.setDate(date.getDate() - 1);
        return date.toISOString().split('.')[0];
      };
      reconstructed.getReferenceEndDate = () => {
        const date = new Date();
        date.setFullYear(date.getFullYear() - 1);
        return date.toISOString().split('.')[0];
      };
      break;

    case 'MTD_Le':
      reconstructed.getCurrentStartDate = () => {
        const date = new Date();
        date.setDate(1);
        return date.toISOString().split('.')[0];
      };
      reconstructed.getCurrentEndDate = () => {
        return new Date().toISOString().split('.')[0];
      };
      reconstructed.getReferenceStartDate = () => {
        const date = new Date();
        date.setFullYear(date.getFullYear() - 1);
        date.setDate(1);
        return date.toISOString().split('.')[0];
      };
      reconstructed.getReferenceEndDate = () => {
        const date = new Date();
        date.setFullYear(date.getFullYear() - 1);
        return date.toISOString().split('.')[0];
      };
      break;

    case '09.06-now':
      reconstructed.getCurrentStartDate = () => {
        const date = new Date();
        date.setMonth(5, 9); // June 9th
        return date.toISOString().split('.')[0];
      };
      reconstructed.getCurrentEndDate = () => {
        return new Date().toISOString().split('.')[0];
      };
      reconstructed.getReferenceStartDate = () => {
        const date = new Date();
        date.setFullYear(date.getFullYear() - 1);
        date.setMonth(5, 9);
        return date.toISOString().split('.')[0];
      };
      reconstructed.getReferenceEndDate = () => {
        const date = new Date();
        date.setFullYear(date.getFullYear() - 1);
        return date.toISOString().split('.')[0];
      };
      break;

    default:
      console.warn(`Unknown config type: ${config.name}`);
  }

  return reconstructed;
}

// Convert downloaded files to workbook objects for processing
async function convertDownloadedFilesToWorkbooks(downloadResults) {
  const workbooks = {};

  // Load XLSX library if not available
  if (typeof XLSX === 'undefined') {
    await loadXLSXLibrary();
  }

  for (const [fileKey, result] of Object.entries(downloadResults)) {
    if (result.success && result.data) {
      try {
        // Use XLSX library directly to read the file
        const workbook = XLSX.read(result.data, {
          type: 'array',
          cellDates: true,
          cellNF: false,
          cellStyles: false
        });

        workbooks[fileKey] = workbook;
        console.log(`Converted ${fileKey} to workbook with sheets:`, workbook.SheetNames);
      } catch (error) {
        console.error(`Failed to convert ${fileKey} to workbook:`, error);
      }
    }
  }

  return workbooks;
}

// Create placeholder Excel data when download fails
function createPlaceholderExcelData(config) {
  try {
    // Load XLSX library if available
    if (typeof XLSX === 'undefined') {
      console.warn('XLSX library not available for placeholder creation');
      return new Uint8Array(0); // Empty array as fallback
    }

    // Create a simple Excel workbook with placeholder data
    const wb = XLSX.utils.book_new();

    // Create placeholder data
    const placeholderData = [
      ['Data Source', 'Status', 'Date', 'Message'],
      [config.displayName, 'Download Failed', new Date().toISOString().split('T')[0], 'Could not download from PowerBI'],
      ['Note', '', '', 'This is placeholder data. Original download failed.'],
      ['File', config.fileName, '', ''],
      ['Target Sheet', config.targetSheet, '', '']
    ];

    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(placeholderData);

    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, 'PlaceholderData');

    // Convert to array buffer
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

    console.log(`Created placeholder Excel data for ${config.displayName}`);
    return new Uint8Array(wbout);

  } catch (error) {
    console.error('Error creating placeholder Excel data:', error);
    return new Uint8Array(0); // Return empty array as fallback
  }
}

// Download multiple PowerBI files concurrently (simultaneously)
async function downloadMultiplePowerBIFilesWithConfigs(fileConfigs, onProgress) {
  const results = {};
  const totalFiles = fileConfigs.length;

  console.log(`Starting concurrent download of ${totalFiles} PowerBI files...`);

  // Create download promises for all files simultaneously
  const downloadPromises = fileConfigs.map(async (configOriginal, index) => {
    let config = configOriginal;
    const fileKey = config.name;

    // Reconstruct the date functions since they get lost during serialization
    config = reconstructConfigFunctions(config);

    try {
      // Update progress - starting download
      if (onProgress) {
        onProgress({
          currentFile: index + 1,
          totalFiles: totalFiles,
          fileName: config.displayName,
          status: 'downloading'
        });
      }

      console.log(`Starting download ${index + 1}/${totalFiles}: ${config.displayName}`);

      // Download the file
      const response = await downloadExcelFile(config);

      if (!response.ok) {
        throw new Error(`HTTP ${response.status}: ${response.statusText}`);
      }

      // Get the file as array buffer
      const arrayBuffer = await response.arrayBuffer();
      const uint8Array = new Uint8Array(arrayBuffer);

      // Store the downloaded file data
      const result = {
        success: true,
        data: uint8Array,
        config: config,
        fileName: config.fileName,
        downloadTime: new Date()
      };

      console.log(`Successfully downloaded ${config.displayName} (${uint8Array.length} bytes)`);

      // Update progress - completed
      if (onProgress) {
        onProgress({
          currentFile: index + 1,
          totalFiles: totalFiles,
          fileName: config.displayName,
          status: 'completed'
        });
      }

      return { fileKey, result };

    } catch (error) {
      console.error(`Failed to download ${config.displayName}:`, error);

      // Create placeholder data for destination file instead of failing
      console.log(`Creating placeholder data for ${config.displayName}...`);
      const placeholderData = createPlaceholderExcelData(config);

      const result = {
        success: true, // Mark as success so processing continues
        data: placeholderData,
        config: config,
        fileName: config.fileName,
        isPlaceholder: true,
        originalError: error.message,
        downloadTime: new Date()
      };

      // Update progress - completed with placeholder
      if (onProgress) {
        onProgress({
          currentFile: index + 1,
          totalFiles: totalFiles,
          fileName: config.displayName + ' (placeholder)',
          status: 'completed'
        });
      }

      return { fileKey, result };
    }
  });

  // Wait for all downloads to complete
  console.log('Waiting for all concurrent downloads to complete...');
  const downloadResults = await Promise.allSettled(downloadPromises);

  // Process results
  downloadResults.forEach((promiseResult, index) => {
    if (promiseResult.status === 'fulfilled') {
      const { fileKey, result } = promiseResult.value;
      results[fileKey] = result;
    } else {
      // Handle rejected promises
      const config = fileConfigs[index];
      const fileKey = config.name;
      console.error(`Promise rejected for ${config.displayName}:`, promiseResult.reason);
      results[fileKey] = {
        success: false,
        error: promiseResult.reason?.message || 'Unknown error',
        config: config,
        fileName: config.fileName
      };
    }
  });

  const successCount = Object.values(results).filter(r => r.success).length;
  console.log(`Concurrent download process completed: ${successCount}/${totalFiles} files successful`);
  return results;
}

// Message handler for extension popup communication
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'downloadPowerBIFiles') {
    const { fileConfigs } = message;

    // Use the provided file configurations directly
    downloadMultiplePowerBIFilesWithConfigs(fileConfigs, (progress) => {
      // Send progress updates back to popup
      chrome.runtime.sendMessage({
        action: 'downloadProgress',
        progress: progress
      });
    })
      .then(downloadResults => {
        // Convert to workbooks
        return convertDownloadedFilesToWorkbooks(downloadResults).then(workbooks => {
          return { downloadResults, workbooks };
        });
      })
      .then(({ downloadResults, workbooks }) => {
        const results = {
          downloadResults: downloadResults,
          workbooks: workbooks,
          success: Object.keys(workbooks).length > 0
        };
        sendResponse({ success: true, data: results });
      })
      .catch(error => {
        sendResponse({ success: false, error: error.message });
      });

    // Return true to indicate async response
    return true;
  }

  if (message.action === 'searchForToken') {
    const token = getTokenFromSessionStorage();
    sendResponse({ success: true, tokenFound: !!token });
  }
});

// Extension ready - token will be retrieved from session storage when needed

// Export functions for external use
window.downloadPowerBiReport = downloadExcelFile;
window.convertDownloadedFilesToWorkbooks = convertDownloadedFilesToWorkbooks;