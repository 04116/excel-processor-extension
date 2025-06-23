// PowerBI Configuration for 3 File Downloads and Sheet Mapping
const today = new Date();
const POWERBI_CONFIG = {
  // Configuration for 3 PowerBI files to download
  files: {
    'Day_Le': {
      name: 'Day_Le',
      displayName: 'Day Le (Yesterday)',
      description: 'Yesterday vs Same Day Last Year',
      fileName: 'Day_Le_Report.xlsx',
      targetSheet: 'Day_Le', // Default target sheet name
      getCurrentStartDate: () => {
        // Yesterday (21/06/2025 if today is 22/06/2025)
        return new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
      },
      getCurrentEndDate: () => {
        // Same as start date - single day (21/06/2025)
        return new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
      },
      getReferenceStartDate: () => {
        // Today last year (22/06/2024 if today is 22/06/2025)
        return new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());
      },
      getReferenceEndDate: () => {
        // Same as start date - single day (22/06/2024)
        return new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());
      }
    },
    'MTD_Le': {
      name: 'MTD_Le',
      displayName: 'MTD Le (Month to Date)',
      description: 'Month to Date vs Same Period Last Year',
      fileName: 'MTD_Le_Report.xlsx',
      targetSheet: 'MTD_Le',
      getCurrentStartDate: () => {
        // First day of current month (01/06/2025)
        return new Date(today.getFullYear(), today.getMonth(), 1);
      },
      getCurrentEndDate: () => {
        // Yesterday (21/06/2025 if today is 22/06/2025)
        return new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
      },
      getReferenceStartDate: () => {
        // First day of same month last year (01/06/2024)
        return new Date(today.getFullYear() - 1, today.getMonth(), 1);
      },
      getReferenceEndDate: () => {
        // Yesterday last year (21/06/2024)
        return new Date(today.getFullYear() - 1, today.getMonth(), today.getDate() - 1);
      }
    },
    '09.06-now': {
      name: '09.06-now',
      displayName: '09.06-now (June 9 to Now)',
      description: 'June 9 to Today vs Same Period Last Year',
      fileName: '09.06-now_Report.xlsx',
      targetSheet: '09.06-now',
      getCurrentStartDate: () => {
        // 09/06/2025 07:00:00 GMT+7
        // month is 0-indexed in JavaScript
        return new Date(today.getFullYear(), 5, 9);
      },
      getCurrentEndDate: () => {
        // Yesterday (21/06/2025 if today is 22/06/2025)
        return new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
      },
      getReferenceStartDate: () => {
        // 10/06/2024 07:00:00 GMT+7
        // month is 0-indexed in JavaScript
        return new Date(today.getFullYear() - 1, 5, 10);
      },
      getReferenceEndDate: () => {
        // Today last year (22/06/2024)
        return new Date(today.getFullYear() - 1, today.getMonth(), today.getDate());
      }
    }
  },

  // Sheet mapping configuration - allows user to customize which downloaded file goes to which sheet
  sheetMapping: {
    // Default mapping - can be overridden by user selection
    'Day_Le': 'Day_Le',     // Day_Le file → Day_Le sheet
    'MTD_Le': 'MTD_Le',     // MTD_Le file → MTD_Le sheet
    '09.06-now': '09.06-now' // 09.06-now file → 09.06-now sheet
  },

  // Download sequence configuration
  downloadSequence: ['Day_Le', 'MTD_Le', '09.06-now'],

  // PowerBI API endpoint
  apiEndpoint: 'https://wabi-south-east-asia-b-primary-redirect.analysis.windows.net/export/xlsx'
};

// Legacy support - keep existing TIME_PERIODS_CONFIG for backward compatibility
const TIME_PERIODS_CONFIG = {};
Object.keys(POWERBI_CONFIG.files).forEach(key => {
  TIME_PERIODS_CONFIG[key] = POWERBI_CONFIG.files[key];
});

// Helper functions for PowerBI configuration
function getPowerBIFiles() {
  return Object.keys(POWERBI_CONFIG.files);
}

function getPowerBIFileConfig(fileName) {
  return POWERBI_CONFIG.files[fileName];
}

function getSheetMapping() {
  return POWERBI_CONFIG.sheetMapping;
}

function updateSheetMapping(fileKey, targetSheet) {
  POWERBI_CONFIG.sheetMapping[fileKey] = targetSheet;
}

function getDownloadSequence() {
  return POWERBI_CONFIG.downloadSequence;
}

// Validate that all target sheets exist in the destination workbook
function validateSheetMapping(outputWorkbook) {
  const availableSheets = outputWorkbook.SheetNames;
  const mappingValidation = {};

  Object.entries(POWERBI_CONFIG.sheetMapping).forEach(([fileKey, targetSheet]) => {
    mappingValidation[fileKey] = {
      targetSheet: targetSheet,
      exists: availableSheets.includes(targetSheet),
      available: availableSheets
    };
  });

  return mappingValidation;
}

// Export configuration
window.PowerBIConfig = {
  files: POWERBI_CONFIG.files,
  sheetMapping: POWERBI_CONFIG.sheetMapping,
  downloadSequence: POWERBI_CONFIG.downloadSequence,
  apiEndpoint: POWERBI_CONFIG.apiEndpoint,
  getPowerBIFiles,
  getPowerBIFileConfig,
  getSheetMapping,
  updateSheetMapping,
  getDownloadSequence,
  validateSheetMapping
};
