// Excel File Processor - PowerBI Integration
// Main popup script for downloading PowerBI files and creating new Excel file

let isProcessing = false;
let hasToken = false;
let timePeriodsData = null; // Store time periods data received from download.js

// DOM elements
const processBtn = document.getElementById('processBtn');
const statusDiv = document.getElementById('status');
const inlineProgress = document.getElementById('inlineProgress');
const progressText = document.getElementById('progressText');
const progressFill = document.getElementById('progressFill');

// PowerBI workflow elements
const powerbiControls = document.getElementById('powerbiControls');
const periodCheckboxes = document.getElementById('periodCheckboxes');

// Initialize extension
document.addEventListener('DOMContentLoaded', function () {
  initializeExtension();
  setupEventListeners();
});

function initializeExtension() {
  // Show PowerBI controls by default
  powerbiControls.style.display = 'block';

  // Request time periods data from download.js
  requestTimePeriodsData();

  // Check token availability on startup to warn user if needed
  checkInitialTokenStatus();

  console.log('Extension initialized for PowerBI download workflow');
}

function setupEventListeners() {
  // Process button
  processBtn.addEventListener('click', handleProcessFiles);
}

// Request time periods data from download.js via message
async function requestTimePeriodsData() {
  try {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const tab = tabs[0];

    if (!tab) {
      console.error('No active tab found');
      return;
    }

    const response = await chrome.tabs.sendMessage(tab.id, {
      action: 'getTimePeriodsData'
    });

    if (response && response.success && response.timePeriodsData) {
      timePeriodsData = response.timePeriodsData;
      generatePowerBIFileCheckboxes();
      updateProcessButtonState();
    } else {
      console.error('Failed to get time periods data:', response?.error || 'No response');
      showStatus('Failed to load time periods data', 'error');
    }
  } catch (error) {
    console.error('Error requesting time periods data:', error);
    showStatus('Failed to load time periods data', 'error');
  }
}

function generatePowerBIFileCheckboxes() {
  if (!timePeriodsData) {
    console.error('Time periods data not loaded');
    return;
  }

  periodCheckboxes.innerHTML = '';

  Object.entries(timePeriodsData).forEach(([fileKey, dateInfo]) => {
    const checkboxDiv = document.createElement('div');
    checkboxDiv.className = 'period-checkbox';

    checkboxDiv.innerHTML = `
      <div style="font-size: 11px; line-height: 1.2;">
        <strong style="color: #333;">${dateInfo.targetSheet}:</strong>
        <span style="color: #f57c00;">${dateInfo.referencePeriod}</span> vs.
        <span style="color: #2e7d32;">${dateInfo.currentPeriod}</span>
      </div>
    `;

    periodCheckboxes.appendChild(checkboxDiv);
  });
}

function updateProcessButtonState() {
  const hasTimePeriodsData = timePeriodsData !== null;
  const hasSelectedFiles = getSelectedPowerBIFiles().length > 0;

  processBtn.disabled = !hasTimePeriodsData || !hasSelectedFiles || isProcessing;

  if (isProcessing) {
    processBtn.textContent = 'Processing...';
  } else if (hasTimePeriodsData && hasSelectedFiles) {
    processBtn.textContent = 'Download PowerBI Data & Create Excel File';
  } else if (!hasTimePeriodsData) {
    processBtn.textContent = 'Loading Time Periods...';
  } else {
    processBtn.textContent = 'No Files to Download';
  }
}

function getSelectedPowerBIFiles() {
  // Return all available PowerBI files since we're downloading all of them
  if (!timePeriodsData) {
    return [];
  }
  return Object.keys(timePeriodsData);
}

async function handleProcessFiles() {
  if (isProcessing) return;

  try {
    isProcessing = true;
    updateProcessButtonState();

    const selectedFiles = getSelectedPowerBIFiles();

    console.log('Starting PowerBI download process:', { selectedFiles });

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

    // Download PowerBI files
    const response = await chrome.tabs.sendMessage(tab.id, {
      action: 'downloadPowerBIFiles'
    });

    if (!response.success) {
      throw new Error(response.error || 'Failed to download PowerBI files');
    }

    showStatus('Processing downloaded files...', 'info');

    // Create new Excel workbook from downloaded files
    const processingResult = await window.ExcelProcessor.createNewExcelFromPowerBIFiles(
      response.data
    );

    if (!processingResult.success) {
      throw new Error(processingResult.error || 'Failed to process downloaded files');
    }

    hideProgress();

    // Generate output filename with timestamp
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[:-]/g, '');
    const filename = `PowerBI_Data_${timestamp}.xlsx`;

    // Download the created file
    window.ExcelProcessor.downloadExcelFile(processingResult.workbook, filename);

    // Show success message
    const successCount = processingResult.results.filter(r => r.success && !r.isPlaceholder).length;
    const placeholderCount = processingResult.results.filter(r => r.isPlaceholder).length;

    let successMessage = `✅ Processing completed: ${successCount} files downloaded`;
    if (placeholderCount > 0) {
      successMessage += `, ${placeholderCount} placeholder(s) created`;
    }
    successMessage += `\n\nCreated sheets: ${processingResult.createdSheets.join(', ')}`;
    successMessage += `\n\nFile saved as: ${filename}`;

    if (placeholderCount > 0) {
      successMessage += `\n\n⚠️ Some files failed to download and were replaced with placeholders. Check the Excel file for details.`;
    }

    showStatus(successMessage, 'success');

    console.log('Processing completed successfully:', processingResult);

  } catch (error) {
    console.error('Error processing files:', error);
    hideProgress();
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

    // Update progress display
    const progressPercent = Math.round((progress.currentFile / progress.totalFiles) * 100);
    progressFill.style.width = `${progressPercent}%`;
    progressText.textContent = `Downloading ${progress.fileName} (${progress.currentFile}/${progress.totalFiles})`;

    // Show progress if not already visible
    if (!inlineProgress.classList.contains('show')) {
      showProgress(`Downloading files...`);
    }

    return true;
  } else if (message.action === 'tokenDetected') {
    handleTokenDetected(message);
    return true;
  }
});

function handleTokenDetected(tokenInfo) {
  console.log('Token detected from:', getApiName(tokenInfo.url));
  hasToken = true;
  onTokenFound();
}

function getApiName(apiUrl) {
  if (apiUrl.includes('wabi-south-east-asia')) return 'PowerBI South East Asia';
  if (apiUrl.includes('wabi-us')) return 'PowerBI US';
  if (apiUrl.includes('wabi-europe')) return 'PowerBI Europe';
  return 'PowerBI API';
}

async function checkTokenAvailability() {
  try {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    const tab = tabs[0];

    if (!tab) {
      console.error('No active tab found');
      return;
    }

    // Content script is already injected by manifest.json, just send message
    const response = await chrome.tabs.sendMessage(tab.id, {
      action: 'searchForToken'
    });

    if (response && response.tokenFound) {
      hasToken = true;
      onTokenFound();
    } else {
      console.log('No PowerBI token found in session storage');
    }
  } catch (error) {
    console.error('Error checking token:', error);
  }
}

async function checkInitialTokenStatus() {
  try {
    await checkTokenAvailability();

    if (!hasToken) {
      console.log('No token found on startup, will check again when needed');
    }
  } catch (error) {
    console.error('Error during initial token check:', error);
  }
}

function onTokenFound() {
  console.log('PowerBI authentication token is available');
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

console.log('PowerBI Excel File Processor popup loaded');
