// Excel File Processor - PowerBI Integration
// Main popup script for handling 3-file PowerBI download and sheet replacement workflow

let outputWorkbook = null;
let outputFileName = '';
let isProcessing = false;
let hasToken = false;

// DOM elements
const outFileInput = document.getElementById('outFile');
const processBtn = document.getElementById('processBtn');
const statusDiv = document.getElementById('status');
const inlineProgress = document.getElementById('inlineProgress');
const progressText = document.getElementById('progressText');
const progressFill = document.getElementById('progressFill');


// PowerBI workflow elements
const powerbiControls = document.getElementById('powerbiControls');
const periodCheckboxes = document.getElementById('periodCheckboxes');
const destSheetSection = document.getElementById('destSheetSection');
const destSheets = document.getElementById('destSheets');

// Initialize extension
document.addEventListener('DOMContentLoaded', function () {
  initializeExtension();
  setupEventListeners();
});

function initializeExtension() {
  // Show PowerBI controls by default
  powerbiControls.style.display = 'block';

  // Generate period checkboxes for the 3 PowerBI files
  generatePowerBIFileCheckboxes();

  // Check token availability on startup to warn user if needed
  checkInitialTokenStatus();

  console.log('Extension initialized for PowerBI 3-file workflow');
}

function setupEventListeners() {
  // Output file selection
  outFileInput.addEventListener('change', handleOutputFileSelection);

  // Process button
  processBtn.addEventListener('click', handleProcessFiles);
}

function generatePowerBIFileCheckboxes() {
  if (!window.PowerBIConfig) {
    console.error('PowerBIConfig not loaded');
    return;
  }

  const files = window.PowerBIConfig.files;
  periodCheckboxes.innerHTML = '';

  Object.entries(files).forEach(([fileKey, config]) => {
    const checkboxDiv = document.createElement('div');
    checkboxDiv.className = 'period-checkbox';

    // Get actual date ranges for this configuration
    const dateInfo = getDateRangeInfo(config);

    checkboxDiv.innerHTML = `
      <div style="font-size: 11px; line-height: 1.2;">
        <strong style="color: #333;">${config.targetSheet}:</strong>
        <span style="color: #f57c00;">${dateInfo.referencePeriod}</span> vs.
        <span style="color: #2e7d32;">${dateInfo.currentPeriod}</span>
      </div>
    `;

    periodCheckboxes.appendChild(checkboxDiv);
  });
}

// Get formatted date range information for display
function getDateRangeInfo(config) {
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

    return {
      currentPeriod: formatDateRange(currentStart, currentEnd),
      referencePeriod: formatDateRange(referenceStart, referenceEnd)
    };

  } catch (error) {
    console.error('Error getting date range info:', error);
    return {
      currentPeriod: 'Date calculation error',
      referencePeriod: 'Date calculation error'
    };
  }
}

async function handleOutputFileSelection(event) {
  const file = event.target.files[0];
  if (!file) return;

  try {
    showProgress('Loading output file...');

    // Read the output file
    const buffer = await window.ExcelProcessor.readFileAsArrayBuffer(file);
    outputWorkbook = window.ExcelProcessor.readExcelFile(buffer, file.name);
    outputFileName = file.name;

    console.log('Output file loaded:', outputWorkbook.SheetNames);

    // Update sheet selectors with available sheets
    updateSheetSelectors(outputWorkbook.SheetNames);

    // Show sheet mapping validation
    displaySheetMappingValidation(outputWorkbook.SheetNames);

    // Show destination sheet section
    destSheetSection.style.display = 'block';

    // Enable process button if we have valid configuration
    updateProcessButtonState();

    hideProgress();
    showStatus('Output file loaded successfully', 'success');

  } catch (error) {
    console.error('Error loading output file:', error);
    hideProgress();
    showStatus(`Error loading output file: ${error.message}`, 'error');
  }
}

function updateSheetSelectors(availableSheets) {
  // Update each sheet selector dropdown
  document.querySelectorAll('.sheet-select').forEach(select => {
    const currentValue = select.value;
    select.innerHTML = '';

    availableSheets.forEach(sheetName => {
      const option = document.createElement('option');
      option.value = sheetName;
      option.textContent = sheetName;
      option.selected = sheetName === currentValue;
      select.appendChild(option);
    });
  });
}

function displaySheetMappingValidation(availableSheets) {
  if (!window.PowerBIConfig) return;

  const destSheets = document.getElementById('destSheets');
  destSheets.innerHTML = '';

  const files = window.PowerBIConfig.files;
  Object.entries(files).forEach(([fileKey, config]) => {
    const targetSheet = config.targetSheet;
    const exists = availableSheets.includes(targetSheet);

    const mappingDiv = document.createElement('div');
    mappingDiv.className = 'period-checkbox'; // Use same class as periods for consistent styling

    const statusIcon = exists ? '✅' : '❌';
    const statusColor = exists ? '#2e7d32' : '#d32f2f';

    mappingDiv.innerHTML = `
      <div style="font-size: 11px; line-height: 1.2;">
        <span style="color: ${statusColor};">"${targetSheet}" ${statusIcon}</span>
      </div>
    `;

    destSheets.appendChild(mappingDiv);
  });
}

function updateProcessButtonState() {
  const hasOutputFile = outputWorkbook !== null;
  const hasSelectedFiles = getSelectedPowerBIFiles().length > 0;

  processBtn.disabled = !hasOutputFile || !hasSelectedFiles || isProcessing;

  if (isProcessing) {
    processBtn.textContent = 'Processing...';
  } else if (hasOutputFile && hasSelectedFiles) {
    processBtn.textContent = 'Download & Fillup data';
  } else if (!hasOutputFile) {
    processBtn.textContent = 'Select Output File First';
  } else {
    processBtn.textContent = 'Select Files to Download';
  }
}

function getSelectedPowerBIFiles() {
  // Since checkboxes are removed, return all available PowerBI files
  if (!window.PowerBIConfig || !window.PowerBIConfig.files) {
    return [];
  }
  return Object.keys(window.PowerBIConfig.files);
}

function getCurrentSheetMapping() {
  const mapping = {};

  // Since checkboxes are removed, map all available PowerBI files
  if (window.PowerBIConfig && window.PowerBIConfig.files) {
    Object.keys(window.PowerBIConfig.files).forEach(fileKey => {
      const sheetSelect = document.getElementById(`sheet_${fileKey}`);
      if (sheetSelect) {
        mapping[fileKey] = sheetSelect.value;
      } else {
        // Use default target sheet from config
        mapping[fileKey] = window.PowerBIConfig.files[fileKey].targetSheet;
      }
    });
  }

  return mapping;
}

function updateSheetMappingFromSelectors() {
  if (!window.PowerBIConfig) return;

  // Update the PowerBI configuration based on user selections
  document.querySelectorAll('.sheet-select').forEach(select => {
    const fileKey = select.id.replace('sheet_', '');
    const selectedSheet = select.value;

    if (window.PowerBIConfig.files[fileKey]) {
      window.PowerBIConfig.files[fileKey].targetSheet = selectedSheet;
      // Also update the global sheet mapping
      if (window.PowerBIConfig.sheetMapping) {
        window.PowerBIConfig.sheetMapping[fileKey] = selectedSheet;
      }
    }
  });
}

async function handleProcessFiles() {
  if (isProcessing) return;

  try {
    isProcessing = true;

    updateProcessButtonState();

    const selectedFiles = getSelectedPowerBIFiles();
    const sheetMapping = getCurrentSheetMapping();

    console.log('Starting PowerBI download process:', { selectedFiles, sheetMapping });

    showStatus('Connecting to PowerBI...', 'info');

    // Send message to content script to start download
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab.url.includes('app.powerbi.com')) {
      throw new Error('❌ Please navigate to app.powerbi.com first');
    }

    // Check if user is on a report page
    if (!tab.url.includes('/reports/') && !tab.url.includes('/groups/')) {
      showStatus('⚠️ For best results, please open a specific PowerBI report before downloading', 'info');
      await new Promise(resolve => setTimeout(resolve, 2000));
    }

    // Inject debug helper for better troubleshooting
    try {
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        files: ['powerbi-debug.js']
      });
      console.log('Debug helper injected');
    } catch (error) {
      console.warn('Could not inject debug helper:', error);
    }

    // Get file configurations to pass to content script
    const fileConfigs = selectedFiles.map(fileKey => {
      if (window.PowerBIConfig && window.PowerBIConfig.files[fileKey]) {
        return window.PowerBIConfig.files[fileKey];
      }
      throw new Error(`Configuration not found for file: ${fileKey}`);
    });

    // Inject content script and start download
    const response = await chrome.tabs.sendMessage(tab.id, {
      action: 'downloadPowerBIFiles',
      selectedFiles: selectedFiles,
      fileConfigs: fileConfigs
    });

    if (!response.success) {
      throw new Error(response.error || 'Failed to download PowerBI files');
    }

    showStatus('Processing downloaded files...', 'info');

    // Process the downloaded files
    const processingResult = await window.ExcelProcessor.processDownloadedPowerBIFiles(
      outputWorkbook,
      response.data,
      sheetMapping
    );

    if (!processingResult.success) {
      throw new Error(processingResult.error || 'Failed to process files');
    }

    // Generate output filename
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
    const processedFileName = outputFileName.replace(/\.([^.]+)$/, `_processed_${timestamp}.$1`);

    // Download the processed file
    window.ExcelProcessor.downloadExcelFile(outputWorkbook, processedFileName);

    // Show success message with placeholder info
    const placeholderCount = processingResult.results.filter(r => r.isPlaceholder).length;
    const successCount = processingResult.results.filter(r => r.success && !r.isPlaceholder).length;

    let successMessage = `✅ Processing completed: ${successCount} files downloaded`;
    if (placeholderCount > 0) {
      successMessage += `, ${placeholderCount} placeholder(s) created`;
    }
    successMessage += `\n\nUpdated sheets: ${processingResult.processedSheets.join(', ')}`;

    if (placeholderCount > 0) {
      successMessage += `\n\n⚠️ Some files failed to download and were replaced with placeholders. Check the processed file for details.`;
    }

    showStatus(successMessage, 'success');

    console.log('Processing completed successfully:', processingResult);

  } catch (error) {
    console.error('Error processing files:', error);
    showStatus(`❌ Error: ${error.message}`, 'error');
  } finally {
    isProcessing = false;
    updateProcessButtonState();
  }
}

// Listen for messages from content script
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.action === 'downloadProgress') {
    const progress = message.progress;

    switch (progress.status) {
      case 'downloading':
        showStatus(`📥 Starting download: ${progress.fileName}...`, 'info');
        break;
      case 'completed':
        showStatus(`✅ Completed: ${progress.fileName}`, 'success');
        break;
      case 'error':
        showStatus(`❌ Failed: ${progress.fileName} - ${progress.error}`, 'error');
        break;
    }
  } else if (message.action === 'tokenDetected') {
    handleTokenDetected(message);
  }
});

// Handle token detection notification
function handleTokenDetected(tokenInfo) {
  console.log('🔔 PowerBI token detected:', tokenInfo);

  // Update token status
  onTokenFound();
}

// Get friendly API name from URL
function getApiName(apiUrl) {
  if (apiUrl.includes('copilotStatus')) return 'Copilot API';
  if (apiUrl.includes('notificationInfo')) return 'Notification API';
  if (apiUrl.includes('userdetails')) return 'User Details API';
  if (apiUrl.includes('analysis.windows.net')) return 'PowerBI API';
  return 'PowerBI Service';
}

// Token status management
// Token status is no longer needed - tokens are checked during processing

async function checkTokenAvailability() {
  try {
    // Get current tab
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab.url.includes('app.powerbi.com')) {
      return false;
    }

    // Content script is already injected by manifest.json, just send message
    const response = await chrome.tabs.sendMessage(tab.id, {
      action: 'searchForToken'
    });

    return response && response.tokenFound;

  } catch (error) {
    console.error('Error checking token availability:', error);
    return false;
  }
}

async function checkInitialTokenStatus() {
  try {
    // Get current tab
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab.url.includes('app.powerbi.com')) {
      showStatus('⚠️ Please navigate to app.powerbi.com to use this extension', 'error');
      return;
    }

    // Check if token is available
    const tokenAvailable = await checkTokenAvailability();

    if (tokenAvailable) {
      hasToken = true;
      console.log('✅ PowerBI token found - extension ready');
    } else {
      showStatus('⚠️ PowerBI token not found. Please make sure you are logged into PowerBI and refresh the page if needed.', 'error');
    }

  } catch (error) {
    console.error('Error checking initial token status:', error);
    showStatus('⚠️ Could not verify PowerBI access. Please make sure you are logged into PowerBI.', 'error');
  }
}

function onTokenFound() {
  hasToken = true; // Set the flag first
}

function showProgress(message) {
  progressText.textContent = message;
  inlineProgress.classList.add('show');
}

function hideProgress() {
  inlineProgress.classList.remove('show');
}

function showStatus(message, type) {
  statusDiv.textContent = message;
  statusDiv.className = `status ${type}`;
  statusDiv.style.display = 'block';

  // Auto-hide success messages after 5 seconds
  if (type === 'success') {
    setTimeout(() => {
      statusDiv.style.display = 'none';
    }, 5000);
  }
}

// Add event listeners for dynamic elements
document.addEventListener('change', function (event) {
  if (event.target.type === 'checkbox' && event.target.id.startsWith('period_')) {
    updateProcessButtonState();
  }

  // Update validation when sheet mapping changes
  if (event.target.classList.contains('sheet-select')) {
    updateSheetMappingFromSelectors();
    if (outputWorkbook) {
      displaySheetMappingValidation(outputWorkbook.SheetNames);
    }
  }
});

console.log('PowerBI Excel File Processor popup loaded');
