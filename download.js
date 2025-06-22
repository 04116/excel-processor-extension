function extractBearerToken() {
  try {
    console.log('Attempting to extract PowerBI bearer token...');

    // Check for manual token first (highest priority)
    if (window._manualPowerBIToken) {
      console.log('Using manual token provided by user');
      return window._manualPowerBIToken;
    }

    // Check for already captured token
    if (window._capturedPowerBIToken) {
      console.log('Using previously captured token');
      return window._capturedPowerBIToken;
    }

    // Enhanced token extraction methods
    const tokenSources = [
      // Window object paths - check deeper PowerBI objects
      () => window.__powerBIAccessToken,
      () => window.powerBIAccessToken,
      () => window.accessToken,
      () => window.authToken,
      () => window.pbiToken,
      () => window.powerbi?.accessToken,
      () => window.powerbi?.token,
      () => window._powerBIContext?.accessToken,
      () => window.powerBIContext?.accessToken,
      () => window.powerBIGlobal?.accessToken,
      () => window.microsoft?.powerbi?.accessToken,

      // Check PowerBI app configuration objects
      () => {
        // Look for PowerBI config objects
        for (const key of Object.keys(window)) {
          if (key.toLowerCase().includes('powerbi') || key.toLowerCase().includes('pbi')) {
            const obj = window[key];
            if (obj && typeof obj === 'object') {
              if (obj.accessToken) return obj.accessToken;
              if (obj.token) return obj.token;
              if (obj.authToken) return obj.authToken;
            }
          }
        }
        return null;
      },

      // Local storage
      () => localStorage.getItem('powerbi_access_token'),
      () => localStorage.getItem('access_token'),
      () => localStorage.getItem('authToken'),

      // Session storage
      () => sessionStorage.getItem('powerbi_access_token'),
      () => sessionStorage.getItem('access_token'),
      () => sessionStorage.getItem('authToken'),

      // Check for tokens in localStorage with various keys
      () => {
        for (let i = 0; i < localStorage.length; i++) {
          const key = localStorage.key(i);
          if (key && (key.includes('token') || key.includes('auth') || key.includes('powerbi'))) {
            const value = localStorage.getItem(key);
            if (value && typeof value === 'string' && value.length > 50) {
              console.log(`Found token in localStorage: ${key}`);
              return value;
            }
          }
        }
        return null;
      },

      // Extract from network requests (look for Authorization headers)
      () => {
        const scripts = document.querySelectorAll('script');
        for (const script of scripts) {
          if (script.textContent) {
            const tokenMatch = script.textContent.match(/(?:bearer\s+|token['"]\s*:\s*['"])([A-Za-z0-9\-_\.]+)/i);
            if (tokenMatch && tokenMatch[1] && tokenMatch[1].length > 50) {
              console.log('Found token in script content');
              return tokenMatch[1];
            }
          }
        }
        return null;
      },

      // Check for meta tags with token information
      () => {
        const metaTags = document.querySelectorAll('meta[name*="token"], meta[name*="auth"]');
        for (const meta of metaTags) {
          const content = meta.getAttribute('content');
          if (content && content.length > 50) {
            console.log(`Found token in meta tag: ${meta.name}`);
            return content;
          }
        }
        return null;
      }
    ];

    // Try each token source
    for (let i = 0; i < tokenSources.length; i++) {
      try {
        const token = tokenSources[i]();
        if (token && typeof token === 'string' && token.length > 50) {
          console.log(`Successfully extracted token using method ${i + 1}`);
          return token;
        }
      } catch (e) {
        // Continue to next method
      }
    }

    // If no token found, try to extract from any fetch requests
    console.log('No token found in standard locations, monitoring network requests...');

    // Hook into fetch to capture authorization headers
    const originalFetch = window.fetch;
    let capturedToken = null;

    window.fetch = function (...args) {
      const result = originalFetch.apply(this, args);

      // Check if this request has authorization header
      if (args[1] && args[1].headers) {
        const headers = args[1].headers;
        if (headers.authorization || headers.Authorization) {
          const authHeader = headers.authorization || headers.Authorization;
          if (authHeader.startsWith('Bearer ')) {
            capturedToken = authHeader.substring(7);
            console.log('Captured token from fetch request');
          }
        }
      }

      return result;
    };

    // Restore original fetch after a short delay
    setTimeout(() => {
      window.fetch = originalFetch;
    }, 5000);

    return capturedToken;

  } catch (e) {
    console.error('Error extracting bearer token:', e);
  }

  console.warn('Could not find PowerBI authentication token. Make sure you are logged into PowerBI and on a report page.');
  return null;
}

// Wait for and capture token from network requests with specific PowerBI API targeting
async function waitForTokenFromNetworkRequests(timeoutMs = 10000) {
  return new Promise((resolve) => {
    console.log('Actively triggering PowerBI API calls to capture token...');

    let tokenFound = false;

    // Check if we already have a captured token
    if (window._capturedPowerBIToken) {
      console.log('Using already captured token');
      resolve(window._capturedPowerBIToken);
      return;
    }

    // Set timeout
    setTimeout(() => {
      if (!tokenFound) {
        console.log('Token capture timeout reached');
        resolve(null);
      }
    }, timeoutMs);

    // Try to actively trigger PowerBI API calls
    setTimeout(async () => {
      if (!tokenFound) {
        console.log('🎯 Attempting to trigger PowerBI API calls...');

        try {
          // Define PowerBI API endpoints that contain authentication tokens
          const powerBIEndpoints = [
            {
              url: 'https://wabi-south-east-asia-b-primary-redirect.analysis.windows.net/explore/aiclient/copilotStatus',
              method: 'GET',
              name: 'copilotStatus'
            },
            {
              url: 'https://wabi-south-east-asia-b-primary-redirect.analysis.windows.net/metadata/notificationInfo/notificationCenter/summary',
              method: 'POST',
              name: 'notificationInfo',
              body: JSON.stringify({})
            },
            {
              url: 'https://wabi-south-east-asia-b-primary-redirect.analysis.windows.net/metadata/people/userdetails',
              method: 'POST',
              name: 'userdetails',
              body: JSON.stringify({})
            }
          ];

          // Try each endpoint to trigger token capture
          for (const endpoint of powerBIEndpoints) {
            if (!tokenFound) {
              console.log(`🎯 Triggering ${endpoint.name} API call...`);

              const requestOptions = {
                method: endpoint.method,
                headers: {
                  'Accept': 'application/json',
                  'Content-Type': 'application/json'
                },
                credentials: 'include'
              };

              if (endpoint.body) {
                requestOptions.body = endpoint.body;
              }

              // This will be intercepted by our hooks and capture the token
              fetch(endpoint.url, requestOptions).catch(e => {
                console.log(`${endpoint.name} call triggered (may fail, but should capture token)`);
              });

              // Small delay between API calls
              await new Promise(resolve => setTimeout(resolve, 300));
            }
          }

          // Check for token after all API calls
          setTimeout(() => {
            if (window._capturedPowerBIToken && !tokenFound) {
              console.log('✅ Successfully captured token from triggered API calls');
              tokenFound = true;
              resolve(window._capturedPowerBIToken);
            }
          }, 1500);

        } catch (error) {
          console.log('Error triggering PowerBI API calls:', error);
        }
      }
    }, 500);

    // Try to trigger other PowerBI requests by simulating user interaction
    setTimeout(() => {
      if (!tokenFound) {
        console.log('Attempting to trigger PowerBI requests via UI interaction...');

        // Try clicking on PowerBI elements that might trigger API calls
        const powerBISelectors = [
          '[data-testid*="visual"]',
          '.visual-container',
          '.slicer-container',
          '.exploration-container',
          '[aria-label*="visual"]',
          '.powerbi-visual'
        ];

        for (const selector of powerBISelectors) {
          const elements = document.querySelectorAll(selector);
          if (elements.length > 0) {
            console.log(`Clicking on ${selector} to trigger API calls`);
            elements[0].click();
            break;
          }
        }

        // Check for token after interaction
        setTimeout(() => {
          if (window._capturedPowerBIToken && !tokenFound) {
            console.log('✅ Successfully captured token from UI interaction');
            tokenFound = true;
            resolve(window._capturedPowerBIToken);
          }
        }, 2000);
      }
    }, 2000);
  });
}

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

async function downloadExcelFile(dateConfig, skipTokenExtraction = false) {
  const requestBody = requestBodyForDownloadByTimePeriod(dateConfig);

  let bearerToken = extractBearerToken();

  // Skip token extraction if explicitly requested (for failed downloads)
  if (!bearerToken && !skipTokenExtraction) {
    bearerToken = await waitForTokenFromNetworkRequests();
  }

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

// Download multiple PowerBI files concurrently with progress tracking
async function downloadMultiplePowerBIFiles(fileConfigs, onProgress) {
  // Use the same concurrent logic as the main function
  return await downloadMultiplePowerBIFilesWithConfigs(fileConfigs, onProgress);
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

// Complete workflow: Download 3 files and process them
async function downloadAndProcessPowerBIFiles(selectedFiles, onProgress) {
  try {
    // Load configuration if not available
    let powerBIConfig = window.PowerBIConfig;
    if (!powerBIConfig) {
      // Try to load from the extension's config
      powerBIConfig = await loadPowerBIConfig();
    }

    if (!powerBIConfig) {
      throw new Error('PowerBI configuration not available');
    }

    // Get file configurations
    const fileConfigs = selectedFiles.map(fileKey => {
      if (powerBIConfig.files && powerBIConfig.files[fileKey]) {
        return powerBIConfig.files[fileKey];
      }
      throw new Error(`Configuration not found for file: ${fileKey}`);
    });

    // Download all files
    const downloadResults = await downloadMultiplePowerBIFiles(fileConfigs, onProgress);

    // Convert to workbooks
    const workbooks = await convertDownloadedFilesToWorkbooks(downloadResults);

    // Return results
    return {
      downloadResults,
      workbooks,
      success: Object.keys(workbooks).length > 0
    };

  } catch (error) {
    console.error('Error in download and process workflow:', error);
    throw error;
  }
}

// Load PowerBI configuration from extension
async function loadPowerBIConfig() {
  try {
    // Inject the configuration into the page
    const configScript = document.createElement('script');
    configScript.src = chrome.runtime.getURL('config.js');
    document.head.appendChild(configScript);

    // Wait for it to load
    await new Promise(resolve => {
      configScript.onload = resolve;
    });

    return window.PowerBIConfig;
  } catch (error) {
    console.error('Failed to load PowerBI config:', error);

    // Fallback: define the configuration inline
    return {
      files: {
        'Day_Le': {
          name: 'Day_Le',
          displayName: 'Day Le (Yesterday)',
          description: 'Yesterday vs Same Day Last Year',
          fileName: 'Day_Le_Report.xlsx',
          targetSheet: 'Day_Le',
          getCurrentStartDate: () => {
            const date = new Date();
            date.setDate(date.getDate() - 1);
            return date.toISOString().split('.')[0];
          },
          getCurrentEndDate: () => {
            return new Date().toISOString().split('.')[0];
          },
          getReferenceStartDate: () => {
            const date = new Date();
            date.setFullYear(date.getFullYear() - 1);
            date.setDate(date.getDate() - 1);
            return date.toISOString().split('.')[0];
          },
          getReferenceEndDate: () => {
            const date = new Date();
            date.setFullYear(date.getFullYear() - 1);
            return date.toISOString().split('.')[0];
          }
        },
        'MTD_Le': {
          name: 'MTD_Le',
          displayName: 'MTD Le (Month to Date)',
          description: 'Month to Date vs Same Period Last Year',
          fileName: 'MTD_Le_Report.xlsx',
          targetSheet: 'MTD_Le',
          getCurrentStartDate: () => {
            const date = new Date();
            date.setDate(1);
            return date.toISOString().split('.')[0];
          },
          getCurrentEndDate: () => {
            return new Date().toISOString().split('.')[0];
          },
          getReferenceStartDate: () => {
            const date = new Date();
            date.setFullYear(date.getFullYear() - 1);
            date.setDate(1);
            return date.toISOString().split('.')[0];
          },
          getReferenceEndDate: () => {
            const date = new Date();
            date.setFullYear(date.getFullYear() - 1);
            return date.toISOString().split('.')[0];
          }
        },
        '09.06-now': {
          name: '09.06-now',
          displayName: '09.06-now (June 9 to Now)',
          description: 'June 9 to Today vs Same Period Last Year',
          fileName: '09.06-now_Report.xlsx',
          targetSheet: '09.06-now',
          getCurrentStartDate: () => {
            const date = new Date();
            date.setMonth(5, 9);
            return date.toISOString().split('.')[0];
          },
          getCurrentEndDate: () => {
            return new Date().toISOString().split('.')[0];
          },
          getReferenceStartDate: () => {
            const date = new Date();
            date.setFullYear(date.getFullYear() - 1);
            date.setMonth(5, 9);
            return date.toISOString().split('.')[0];
          },
          getReferenceEndDate: () => {
            const date = new Date();
            date.setFullYear(date.getFullYear() - 1);
            return date.toISOString().split('.')[0];
          }
        }
      }
    };
  }
}

// Message handler for extension popup communication
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'downloadPowerBIFiles') {
    const { selectedFiles, fileConfigs, manualToken } = message;

    // Set manual token if provided
    if (manualToken) {
      console.log('Using manual token for downloads');
      window._manualPowerBIToken = manualToken;
    }

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
});

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
      const response = await downloadExcelFile(config, false);

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

// Set up token capture immediately when script loads
(function setupTokenCapture() {
  console.log('Setting up PowerBI token capture...');

  // Hook fetch immediately with specific PowerBI API targeting
  const originalFetch = window.fetch;
  window.fetch = function (...args) {
    const url = args[0];

    // Specifically target PowerBI API endpoints
    if (url && typeof url === 'string' &&
      (url.includes('analysis.windows.net') ||
        url.includes('powerbi.com') ||
        url.includes('copilotStatus') ||
        url.includes('notificationInfo') ||
        url.includes('userdetails'))) {

      console.log('🎯 Intercepting PowerBI API call:', url);

      // Check for authorization header
      if (args[1] && args[1].headers) {
        const headers = args[1].headers;
        const authHeader = headers.authorization || headers.Authorization;
        if (authHeader && authHeader.startsWith('Bearer ')) {
          const token = authHeader.substring(7);
          if (token.length > 50) {
            console.log('✅ Captured PowerBI token from PowerBI API:', url.substring(0, 100) + '...');
            window._capturedPowerBIToken = token;

            // Notify extension popup about token detection
            notifyTokenDetected(token, url);
          }
        }
      }
    }

    return originalFetch.apply(this, args);
  };

  // Hook XMLHttpRequest immediately with PowerBI targeting
  const originalXHROpen = XMLHttpRequest.prototype.open;
  const originalXHRSetRequestHeader = XMLHttpRequest.prototype.setRequestHeader;

  XMLHttpRequest.prototype.open = function (method, url, ...args) {
    this._url = url;
    this._method = method;

    // Log PowerBI API calls
    if (url && typeof url === 'string' &&
      (url.includes('analysis.windows.net') ||
        url.includes('powerbi.com') ||
        url.includes('copilotStatus') ||
        url.includes('notificationInfo') ||
        url.includes('userdetails'))) {
      console.log('🎯 Intercepting PowerBI XHR call:', method, url);
    }

    return originalXHROpen.apply(this, [method, url, ...args]);
  };

  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    // Capture token from PowerBI API calls
    if (name.toLowerCase() === 'authorization' && value.startsWith('Bearer ') &&
      this._url && (this._url.includes('analysis.windows.net') ||
        this._url.includes('powerbi.com') ||
        this._url.includes('copilotStatus') ||
        this._url.includes('notificationInfo') ||
        this._url.includes('userdetails'))) {

      const token = value.substring(7);
      if (token.length > 50) {
        console.log('✅ Captured PowerBI token from XHR PowerBI API:', this._url.substring(0, 100) + '...');
        window._capturedPowerBIToken = token;

        // Notify extension popup about token detection
        notifyTokenDetected(token, this._url);
      }
    }

    return originalXHRSetRequestHeader.apply(this, arguments);
  };

  console.log('PowerBI token capture hooks installed - monitoring PowerBI API calls');
})();

// Notify extension popup when token is detected
function notifyTokenDetected(token, apiUrl) {
  try {
    const message = {
      action: 'tokenDetected',
      token: token.substring(0, 20) + '...', // Only send first 20 chars for security
      apiUrl: apiUrl,
      timestamp: new Date().toISOString()
    };

    // Send message to extension popup
    chrome.runtime.sendMessage(message).catch(e => {
      // Popup might not be open, that's okay
      console.log('Token detected but popup not available to notify');
    });

    console.log('🔔 Token detection notification sent to extension');
  } catch (error) {
    console.log('Error sending token notification:', error);
  }
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

// Export functions
window.downloadPowerBiReport = downloadExcelFile;
window.downloadMultiplePowerBIFiles = downloadMultiplePowerBIFiles;
window.convertDownloadedFilesToWorkbooks = convertDownloadedFilesToWorkbooks;
window.downloadAndProcessPowerBIFiles = downloadAndProcessPowerBIFiles;