// Excel File Processor - PowerBI Integration
// Main popup script for handling 3-file PowerBI download and sheet replacement workflow

let outputWorkbook = null;
let outputFileName = '';
let isProcessing = false;
let tokenStatus = 'searching'; // 'searching', 'found', 'none'
let hasToken = false;

// DOM elements
const outFileInput = document.getElementById('outFile');
const processBtn = document.getElementById('processBtn');
const statusDiv = document.getElementById('status');
const inlineProgress = document.getElementById('inlineProgress');
const progressText = document.getElementById('progressText');
const progressFill = document.getElementById('progressFill');
const tokenStatusDiv = document.getElementById('tokenStatus');

// PowerBI workflow elements
const powerbiControls = document.getElementById('powerbiControls');
const periodCheckboxes = document.getElementById('periodCheckboxes');
const destSheetSection = document.getElementById('destSheetSection');
const destSheets = document.getElementById('destSheets');

// Initialize extension
document.addEventListener('DOMContentLoaded', function () {
  initializeExtension();
  setupEventListeners();
  loadRememberedFiles();
});

function initializeExtension() {
  // Show PowerBI controls by default
  powerbiControls.style.display = 'block';

  // Generate period checkboxes for the 3 PowerBI files
  generatePowerBIFileCheckboxes();

  // Start token search
  startTokenSearch();

  console.log('Extension initialized for PowerBI 3-file workflow');
}

function setupEventListeners() {
  // Output file selection
  outFileInput.addEventListener('change', handleOutputFileSelection);

  // Process button
  processBtn.addEventListener('click', handleProcessFiles);



  // File memory system
  document.querySelectorAll('.reuse-btn').forEach(btn => {
    btn.addEventListener('click', handleReuseFile);
  });

  document.querySelectorAll('.clear-btn').forEach(btn => {
    btn.addEventListener('click', handleClearFile);
  });
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

    // Remember this file
    rememberFile('out', file.name, buffer);

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
    // Get current target sheet from dropdown (if exists) or default from config
    const sheetSelect = document.getElementById(`sheet_${fileKey}`);
    const targetSheet = sheetSelect ? sheetSelect.value : config.targetSheet;
    const exists = availableSheets.includes(targetSheet);

    const validationDiv = document.createElement('div');
    validationDiv.className = 'sheet-checkbox';

    const statusIcon = exists ? '✅' : '❌';
    const statusClass = exists ? 'success' : 'error';
    const statusText = exists ? 'Found' : 'Missing';

    validationDiv.innerHTML = `
      <div style="display: flex; align-items: center; width: 100%;">
        <span style="margin-right: 8px; font-size: 14px;">${statusIcon}</span>
        <div style="flex: 1;">
          <strong>${config.displayName}</strong>
          <div class="sheet-info">Target sheet: "${targetSheet}" - <span style="color: ${exists ? 'green' : 'red'}">${statusText}</span></div>
        </div>
      </div>
    `;

    if (!exists) {
      validationDiv.style.backgroundColor = '#ffebee';
      validationDiv.style.borderLeft = '3px solid #f44336';
    } else {
      validationDiv.style.backgroundColor = '#e8f5e8';
      validationDiv.style.borderLeft = '3px solid #4CAF50';
    }

    destSheets.appendChild(validationDiv);
  });

  // Show available sheets if there are missing mappings
  const missingSheets = Object.entries(files).filter(([fileKey, config]) => {
    const sheetSelect = document.getElementById(`sheet_${fileKey}`);
    const targetSheet = sheetSelect ? sheetSelect.value : config.targetSheet;
    return !availableSheets.includes(targetSheet);
  });

  if (missingSheets.length > 0) {
    const availableSheetsDiv = document.createElement('div');
    availableSheetsDiv.className = 'sheet-checkbox';
    availableSheetsDiv.style.backgroundColor = '#f0f8ff';
    availableSheetsDiv.style.borderLeft = '3px solid #2196F3';
    availableSheetsDiv.style.marginTop = '10px';

    availableSheetsDiv.innerHTML = `
      <div>
        <strong>📋 Available sheets in your file:</strong>
        <div class="sheet-info">${availableSheets.join(', ')}</div>
        <div class="sheet-info" style="margin-top: 5px; font-style: italic;">
          💡 Tip: Use the dropdowns above to map PowerBI files to existing sheets
        </div>
      </div>
    `;

    destSheets.appendChild(availableSheetsDiv);
  }
}

function updateProcessButtonState() {
  const hasOutputFile = outputWorkbook !== null;
  const hasSelectedFiles = getSelectedPowerBIFiles().length > 0;
  const waitingForToken = tokenStatus === 'searching';
  const noTokenFound = tokenStatus === 'none' && !hasToken;

  processBtn.disabled = !hasOutputFile || !hasSelectedFiles || isProcessing || waitingForToken || noTokenFound;

  if (waitingForToken) {
    processBtn.textContent = 'Waiting for PowerBI Token...';
  } else if (noTokenFound) {
    processBtn.textContent = 'No PowerBI Token Found - Cannot Process';
  } else if (hasOutputFile && hasSelectedFiles && hasToken) {
    processBtn.textContent = 'Download & Process Files';
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

    // Start token search if not already found
    if (!hasToken) {
      startTokenSearch();
    }

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
function updateTokenStatus(status) {
  tokenStatus = status;

  if (!tokenStatusDiv) return;

  switch (status) {
    case 'searching':
      tokenStatusDiv.textContent = 'Finding token...';
      tokenStatusDiv.className = 'token-status searching';
      tokenStatusDiv.style.display = 'block';
      tokenStatusDiv.style.cursor = 'default';
      tokenStatusDiv.onclick = null;
      tokenStatusDiv.title = '';
      // Don't change hasToken when searching - might already have one
      break;
    case 'found':
      // Token found - hide the status completely  
      tokenStatusDiv.style.display = 'none';
      tokenStatusDiv.onclick = null;
      tokenStatusDiv.title = '';
      // Don't change hasToken here - it's set in onTokenFound
      break;
    case 'none':
      if (!hasToken) {
        tokenStatusDiv.textContent = 'No token found';
        tokenStatusDiv.className = 'token-status';
        tokenStatusDiv.style.display = 'block';
        tokenStatusDiv.style.color = '#d32f2f';
        tokenStatusDiv.style.backgroundColor = '#ffebee';
        tokenStatusDiv.style.cursor = 'pointer';
        tokenStatusDiv.title = 'Click to retry token search';
        tokenStatusDiv.onclick = retryTokenSearch;
      } else {
        tokenStatusDiv.style.display = 'none';
      }
      // Only set hasToken = false if we really don't have a token
      break;
  }

  // Update process button state
  updateProcessButtonState();
}

async function startTokenSearch() {
  updateTokenStatus('searching');

  try {
    // Get current tab
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    if (!tab.url.includes('app.powerbi.com')) {
      console.log('Not on PowerBI page, skipping token search');
      hasToken = false; // Explicitly set to false since we can't search
      updateTokenStatus('none');
      showStatus('❌ Please navigate to app.powerbi.com to find PowerBI token', 'error');
      return;
    }

    // Inject content script if needed
    try {
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        files: ['download.js']
      });
    } catch (error) {
      console.log('Content script already injected or error injecting:', error.message);
    }

    // Request token search from content script
    const response = await chrome.tabs.sendMessage(tab.id, {
      action: 'searchForToken'
    });

    if (response && response.tokenFound) {
      console.log('Token found immediately in session storage');
      // Token detection notification will come via separate message
    } else {
      console.log('Token not found in session storage');
      // Set timeout to stop searching if no token found
      setTimeout(() => {
        if (tokenStatus === 'searching') {
          console.log('No token found after search');
          hasToken = false; // Explicitly set to false since we didn't find one
          updateTokenStatus('none');
          showStatus('❌ No PowerBI token found. Please make sure you are logged into PowerBI and try refreshing the page.', 'error');
        }
      }, 3000);
    }

  } catch (error) {
    console.error('Error starting token search:', error);
    updateTokenStatus('none');
  }
}

function onTokenFound() {
  hasToken = true; // Set the flag first
  updateTokenStatus('none'); // Hide token status immediately when found
}

function retryTokenSearch() {
  console.log('Retrying token search...');
  startTokenSearch();
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

// File memory system
function rememberFile(type, filename, buffer) {
  const key = `remembered_${type}_file`;
  chrome.storage.local.set({
    [key]: {
      filename: filename,
      buffer: Array.from(buffer),
      timestamp: Date.now()
    }
  });

  updateRememberedFileDisplay(type, filename);
}

function updateRememberedFileDisplay(type, filename) {
  const rememberedDiv = document.getElementById(`remembered${type.charAt(0).toUpperCase() + type.slice(1)}File`);
  if (rememberedDiv) {
    rememberedDiv.style.display = 'flex';
    rememberedDiv.querySelector('.file-name').textContent = filename;
  }
}

function loadRememberedFiles() {
  chrome.storage.local.get(['remembered_out_file'], (result) => {
    if (result.remembered_out_file) {
      updateRememberedFileDisplay('out', result.remembered_out_file.filename);
    }
  });
}

function handleReuseFile(event) {
  const type = event.target.getAttribute('data-type');
  const key = `remembered_${type}_file`;

  chrome.storage.local.get([key], (result) => {
    if (result[key]) {
      const buffer = new Uint8Array(result[key].buffer);

      if (type === 'out') {
        try {
          outputWorkbook = window.ExcelProcessor.readExcelFile(buffer, result[key].filename);
          outputFileName = result[key].filename;

          updateSheetSelectors(outputWorkbook.SheetNames);
          destSheetSection.style.display = 'block';
          updateProcessButtonState();

          showStatus('Reused remembered output file', 'success');
        } catch (error) {
          showStatus(`Error loading remembered file: ${error.message}`, 'error');
        }
      }
    }
  });
}

function handleClearFile(event) {
  const type = event.target.getAttribute('data-type');
  const key = `remembered_${type}_file`;

  chrome.storage.local.remove([key]);

  const rememberedDiv = document.getElementById(`remembered${type.charAt(0).toUpperCase() + type.slice(1)}File`);
  if (rememberedDiv) {
    rememberedDiv.style.display = 'none';
  }

  if (type === 'out') {
    outputWorkbook = null;
    outputFileName = '';
    destSheetSection.style.display = 'none';
    updateProcessButtonState();
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
