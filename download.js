//#################################################################################
//# POWERBI CONFIGURATION SECTION
//#################################################################################

// PowerBI Configuration for 3 File Downloads and Sheet Mapping
const POWERBI_CONFIG = {
  // Configuration for 3 PowerBI files to download
  files: {
    'Day_Le': {
      name: 'Day_Le',
      displayName: 'Day Le (Yesterday)',
      description: 'Yesterday vs Same Day Last Year',
      fileName: 'Day_Le_Report.xlsx',
      targetSheet: 'Day_Le',
      getCurrentStartDate: () => {
        const today = new Date();
        return new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
      },
      getCurrentEndDate: () => {
        const today = new Date();
        return new Date(today.getFullYear(), today.getMonth(), today.getDate());
      },
      getReferenceStartDate: () => {
        const today = new Date();
        return new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());
      },
      getReferenceEndDate: () => {
        const today = new Date();
        return new Date(today.getFullYear() - 1, today.getMonth(), today.getDate() + 1);
      }
    },
    'MTD_Le': {
      name: 'MTD_Le',
      displayName: 'MTD Le (Month to Date)',
      description: 'Month to Date vs Same Period Last Year',
      fileName: 'MTD_Le_Report.xlsx',
      targetSheet: 'MTD_Le',
      getCurrentStartDate: () => {
        const today = new Date();
        return new Date(today.getFullYear(), today.getMonth(), 1);
      },
      getCurrentEndDate: () => {
        const today = new Date();
        return new Date(today.getFullYear(), today.getMonth(), today.getDate());
      },
      getReferenceStartDate: () => {
        const today = new Date();
        return new Date(today.getFullYear() - 1, today.getMonth(), 1);
      },
      getReferenceEndDate: () => {
        const today = new Date();
        return new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());
      }
    },
    '09.06-now': {
      name: '09.06-now',
      displayName: '09.06-now (June 9 to Now)',
      description: 'June 9 to Today vs Same Period Last Year',
      fileName: '09.06-now_Report.xlsx',
      targetSheet: '09.06-now',
      getCurrentStartDate: () => {
        const today = new Date();
        return new Date(today.getFullYear(), 5, 9);
      },
      getCurrentEndDate: () => {
        const today = new Date();
        return new Date(today.getFullYear(), today.getMonth(), today.getDate());
      },
      getReferenceStartDate: () => {
        const today = new Date();
        return new Date(today.getFullYear() - 1, 5, 10);
      },
      getReferenceEndDate: () => {
        const today = new Date();
        return new Date(today.getFullYear() - 1, today.getMonth(), today.getDate() + 1);
      }
    }
  },

  // PowerBI API endpoint
  apiEndpoint: 'https://wabi-south-east-asia-b-primary-redirect.analysis.windows.net/export/xlsx'
};

//#################################################################################
//# POWERBI API REQUEST BODY GENERATION
//#################################################################################

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

//#################################################################################
//# POWERBI AUTHENTICATION
//#################################################################################

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

//#################################################################################
//# EXCEL LIBRARY MANAGEMENT
//#################################################################################

// Load XLSX library into content script context
async function loadXLSXLibrary() {
  return new Promise((resolve, reject) => {
    try {
      // Check if already loaded
      if (typeof XLSX !== 'undefined' || window.XLSX) {
        console.log('XLSX library already available');
        resolve();
        return;
      }

      const script = document.createElement('script');
      script.src = chrome.runtime.getURL('libs/xlsx.full.min.js');
      script.onload = () => {
        console.log('XLSX library script loaded in content script');

        // Wait a moment for the library to initialize and try multiple ways to access it
        setTimeout(() => {
          // Try multiple ways to access XLSX
          if (typeof XLSX !== 'undefined') {
            console.log('XLSX library confirmed available via global XLSX');
            window.XLSX = XLSX; // Ensure it's on window object
            resolve();
          } else if (window.XLSX) {
            console.log('XLSX library confirmed available via window.XLSX');
            resolve();
          } else if (typeof SheetJS !== 'undefined') {
            console.log('XLSX library available via SheetJS');
            window.XLSX = SheetJS;
            resolve();
          } else {
            // Try to manually execute the script content
            console.log('Attempting alternative XLSX library access...');
            try {
              // Check if there are any global objects added
              const possibleXLSX = window.XLSX || globalThis.XLSX ||
                (typeof exports !== 'undefined' ? exports.XLSX : null);
              if (possibleXLSX) {
                window.XLSX = possibleXLSX;
                console.log('XLSX library found via alternative method');
                resolve();
              } else {
                console.error('XLSX library loaded but not accessible through any method');
                reject(new Error('XLSX library not accessible after loading'));
              }
            } catch (altError) {
              console.error('Alternative XLSX access failed:', altError);
              reject(new Error('XLSX library not available after loading'));
            }
          }
        }, 100);
      };
      script.onerror = (error) => {
        console.error('Failed to load XLSX library script:', error);
        reject(new Error('Failed to load XLSX library'));
      };
      document.head.appendChild(script);
    } catch (error) {
      reject(error);
    }
  });
}



// Convert downloaded files to workbook objects for processing
async function convertDownloadedFilesToWorkbooks(downloadResults) {
  const workbooks = {};

  // Load XLSX library if not available
  if (typeof XLSX === 'undefined' && !window.XLSX) {
    console.log('Loading XLSX library...');
    await loadXLSXLibrary();

    // Wait a bit more to ensure the library is fully loaded
    await new Promise(resolve => setTimeout(resolve, 100));

    // Check again if XLSX is available
    if (typeof XLSX === 'undefined' && !window.XLSX) {
      console.error('XLSX library failed to load properly');
      return workbooks;
    }
  }

  // Use the available XLSX reference
  const XLSXLib = window.XLSX || XLSX;

  for (const [fileKey, result] of Object.entries(downloadResults)) {
    if (result.success && result.data) {
      try {
        // Double-check XLSX is available before using it
        if (!XLSXLib) {
          console.error(`XLSX library not available when processing ${fileKey}`);
          continue;
        }

        // Use XLSX library directly to read the file
        const workbook = XLSXLib.read(result.data, {
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
async function createPlaceholderExcelData(config) {
  try {
    // Ensure XLSX library is loaded
    if (typeof XLSX === 'undefined' && !window.XLSX) {
      console.log('Loading XLSX library for placeholder creation...');
      await loadXLSXLibrary();

      if (typeof XLSX === 'undefined' && !window.XLSX) {
        console.warn('XLSX library not available for placeholder creation');
        return new Uint8Array(0); // Empty array as fallback
      }
    }

    // Use the available XLSX reference
    const XLSXLib = window.XLSX || XLSX;

    // Create a simple Excel workbook with placeholder data
    const wb = XLSXLib.utils.book_new();

    // Create placeholder data
    const placeholderData = [
      ['Data Source', 'Status', 'Date', 'Message'],
      [config.displayName, 'Download Failed', new Date().toISOString().split('T')[0], 'Could not download from PowerBI'],
      ['Note', '', '', 'This is placeholder data. Original download failed.'],
      ['File', config.fileName, '', ''],
      ['Target Sheet', config.targetSheet, '', '']
    ];

    // Create worksheet
    const ws = XLSXLib.utils.aoa_to_sheet(placeholderData);

    // Add worksheet to workbook
    XLSXLib.utils.book_append_sheet(wb, ws, 'PlaceholderData');

    // Convert to array buffer
    const wbout = XLSXLib.write(wb, { bookType: 'xlsx', type: 'array' });

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
  const downloadPromises = fileConfigs.map(async (config, index) => {
    const fileKey = config.name;

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
      const placeholderData = await createPlaceholderExcelData(config);

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

//#################################################################################
//# CONFIGURATION LOADING
//#################################################################################

// Load config directly from embedded configuration
function loadConfigFromFile() {
  try {
    console.log('Using embedded PowerBI configuration');
    return POWERBI_CONFIG;
  } catch (error) {
    console.error('Error accessing config:', error);
    throw error;
  }
}

//#################################################################################
//# DATE CONVERSION UTILITIES
//#################################################################################

// Convert Date to GMT+7 timezone with time reset to 00:00:00
function convertToGMT7ISOString(date) {
  if (!(date instanceof Date)) return date;

  // Get GMT+7 timezone offset (7 hours = 7 * 60 minutes)
  const gmt7OffsetMinutes = 7 * 60;

  // Create new date in GMT+7 timezone
  const gmt7Date = new Date(date.getTime() + (gmt7OffsetMinutes * 60 * 1000));

  // Get year, month, day in GMT+7 timezone
  const year = gmt7Date.getUTCFullYear();
  const month = String(gmt7Date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(gmt7Date.getUTCDate()).padStart(2, '0');

  // Return date with time reset to 00:00:00
  return `${year}-${month}-${day}T00:00:00`;
}

//#################################################################################
//# TIME PERIODS DATA FORMATTING
//#################################################################################

// Get formatted time periods data for popup display
function getFormattedTimePeriodsData() {
  const powerBIConfig = loadConfigFromFile();
  const timePeriodsData = {};

  Object.entries(powerBIConfig.files).forEach(([fileKey, config]) => {
    try {
      const currentStart = config.getCurrentStartDate();
      const currentEnd = config.getCurrentEndDate();
      const referenceStart = config.getReferenceStartDate();
      const referenceEnd = config.getReferenceEndDate();

      const formatDate = (date) => {
        const options = {
          timeZone: 'Asia/Ho_Chi_Minh',
          year: 'numeric',
          month: '2-digit',
          day: '2-digit'
        };

        return date.toLocaleDateString('en-GB', options);
      };

      const formatDateRange = (startDate, endDate) => {
        const start = formatDate(startDate);
        const end = formatDate(endDate);

        return `${start} - ${end}`;
      };

      timePeriodsData[fileKey] = {
        targetSheet: config.targetSheet,
        currentPeriod: formatDateRange(currentStart, currentEnd),
        referencePeriod: formatDateRange(referenceStart, referenceEnd)
      };

    } catch (error) {
      console.error(`Error getting date range info for ${fileKey}:`, error);
      timePeriodsData[fileKey] = {
        targetSheet: config.targetSheet,
        currentPeriod: 'Date calculation error',
        referencePeriod: 'Date calculation error'
      };
    }
  });

  return timePeriodsData;
}

//#################################################################################
//# CONFIG FORMAT CONVERSION
//#################################################################################

// Convert PowerBI config to file configs format expected by download function
function convertPowerBIConfigToFileConfigs(powerBIConfig) {
  const fileConfigs = [];

  Object.entries(powerBIConfig.files).forEach(([key, config]) => {
    fileConfigs.push({
      name: config.name,
      displayName: config.displayName,
      description: config.description,
      fileName: config.fileName,
      targetSheet: config.targetSheet,
      // Convert Date objects to GMT+7 ISO strings for API
      getCurrentStartDate: () => {
        const date = config.getCurrentStartDate();
        return convertToGMT7ISOString(date);
      },
      getCurrentEndDate: () => {
        const date = config.getCurrentEndDate();
        return convertToGMT7ISOString(date);
      },
      getReferenceStartDate: () => {
        const date = config.getReferenceStartDate();
        return convertToGMT7ISOString(date);
      },
      getReferenceEndDate: () => {
        const date = config.getReferenceEndDate();
        return convertToGMT7ISOString(date);
      }
    });
  });

  return fileConfigs;
}

//#################################################################################
//# CHROME EXTENSION MESSAGE HANDLING
//#################################################################################

// Message handler for extension popup communication
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'downloadPowerBIFiles') {
    try {
      // Load config from embedded configuration
      const powerBIConfig = loadConfigFromFile();

      // Convert config to expected format
      const fileConfigs = convertPowerBIConfigToFileConfigs(powerBIConfig);

      console.log('Loaded file configs:', fileConfigs.map(c => c.displayName));

      // Start download process with loaded configs
      downloadMultiplePowerBIFilesWithConfigs(fileConfigs, (progress) => {
        // Send progress updates back to popup
        chrome.runtime.sendMessage({
          action: 'downloadProgress',
          progress: progress
        });
      })
        .then(downloadResults => {
          // Return raw download results - workbook conversion will be handled in popup
          const results = {
            downloadResults: downloadResults,
            success: Object.values(downloadResults).some(r => r.success)
          };
          sendResponse({ success: true, data: results });
        })
        .catch(error => {
          console.error('Download process failed:', error);
          sendResponse({ success: false, error: error.message });
        });

    } catch (error) {
      console.error('Config loading failed:', error);
      sendResponse({ success: false, error: error.message });
    }

    // Return true to indicate async response
    return true;
  }

  if (message.action === 'searchForToken') {
    const token = getTokenFromSessionStorage();
    sendResponse({ success: true, tokenFound: !!token });
  }

  if (message.action === 'getTimePeriodsData') {
    try {
      const timePeriodsData = getFormattedTimePeriodsData();
      sendResponse({ success: true, timePeriodsData });
    } catch (error) {
      console.error('Error getting time periods data:', error);
      sendResponse({ success: false, error: error.message });
    }
  }
});

//#################################################################################
//# INITIALIZATION
//#################################################################################

// Initialize XLSX library early to prevent timing issues
(async function initializeXLSXLibrary() {
  try {
    if (typeof XLSX === 'undefined') {
      console.log('Pre-loading XLSX library...');
      await loadXLSXLibrary();
      console.log('XLSX library pre-loaded successfully');
    }
  } catch (error) {
    console.warn('Failed to pre-load XLSX library:', error);
  }
})();

//#################################################################################
//# EXTERNAL API EXPORTS
//#################################################################################

// Extension ready - token will be retrieved from session storage when needed

// Export functions for external use
window.downloadPowerBiReport = downloadExcelFile;
window.convertDownloadedFilesToWorkbooks = convertDownloadedFilesToWorkbooks;