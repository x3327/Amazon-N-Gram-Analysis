/**
 * N-gram Automation Tool - Frontend JavaScript
 */

// DOM Elements
const uploadBox = document.getElementById('upload-box');
const fileInput = document.getElementById('file-input');
const browseBtn = document.getElementById('browse-btn');
const fileInfo = document.getElementById('file-info');
const settingsSection = document.getElementById('settings-section');
const processBtn = document.getElementById('process-btn');
const progressSection = document.getElementById('progress-section');
const progressText = document.getElementById('progress-text');
const resultsSection = document.getElementById('results-section');
const errorSection = document.getElementById('error-section');
const errorMessage = document.getElementById('error-message');
const downloadBtn = document.getElementById('download-btn');
const resetBtn = document.getElementById('reset-btn');
const retryBtn = document.getElementById('retry-btn');

// State
let selectedFile = null;
let outputFilename = null;

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
});

function initializeEventListeners() {
    // Browse button click
    browseBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        fileInput.click();
    });

    // Upload box click
    uploadBox.addEventListener('click', () => {
        fileInput.click();
    });

    // File input change
    fileInput.addEventListener('change', handleFileSelect);

    // Drag and drop
    uploadBox.addEventListener('dragover', handleDragOver);
    uploadBox.addEventListener('dragleave', handleDragLeave);
    uploadBox.addEventListener('drop', handleDrop);

    // Process button
    processBtn.addEventListener('click', processFile);

    // Download button
    downloadBtn.addEventListener('click', downloadFile);

    // Reset button
    resetBtn.addEventListener('click', resetForm);

    // Retry button
    retryBtn.addEventListener('click', resetForm);
}

function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadBox.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadBox.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    uploadBox.classList.remove('dragover');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
}

function handleFile(file) {
    // Validate file type
    if (!file.name.toLowerCase().endsWith('.csv')) {
        showError('Please upload a CSV file.');
        return;
    }

    // Validate file size (50MB max)
    if (file.size > 50 * 1024 * 1024) {
        showError('File size exceeds 50MB limit.');
        return;
    }

    selectedFile = file;
    
    // Update UI
    fileInfo.textContent = `Selected: ${file.name} (${formatFileSize(file.size)})`;
    fileInfo.classList.add('success');
    
    // Show settings section
    settingsSection.style.display = 'block';
    
    // Scroll to settings
    settingsSection.scrollIntoView({ behavior: 'smooth' });
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function processFile() {
    if (!selectedFile) {
        showError('Please select a file first.');
        return;
    }

    // Show progress
    showSection('progress');
    updateProgress('Uploading file...');

    // Prepare form data
    const formData = new FormData();
    formData.append('file', selectedFile);
    formData.append('min_clicks', document.getElementById('min-clicks').value);
    formData.append('min_spend', document.getElementById('min-spend').value);

    try {
        updateProgress('Processing CSV data...');
        
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (result.success) {
            outputFilename = result.output_file;
            displayResults(result);
        } else {
            showError(result.error || 'An error occurred while processing the file.');
        }
    } catch (error) {
        console.error('Error:', error);
        showError('Network error. Please check your connection and try again.');
    }
}

function displayResults(result) {
    // Update summary cards
    document.getElementById('total-rows').textContent = result.summary.original_rows.toLocaleString();
    document.getElementById('asins-removed').textContent = result.summary.asins_removed.toLocaleString();
    document.getElementById('campaigns-count').textContent = result.summary.campaigns_processed.toLocaleString();
    document.getElementById('flagged-count').textContent = result.summary.total_flagged.toLocaleString();

    // Update metrics
    document.getElementById('total-spend').textContent = formatCurrency(result.summary.total_spend);
    document.getElementById('total-sales').textContent = formatCurrency(result.summary.total_sales);

    // Update campaigns list
    const campaignsList = document.getElementById('campaign-names');
    campaignsList.innerHTML = '';
    result.campaigns.forEach(campaign => {
        const li = document.createElement('li');
        li.textContent = campaign;
        campaignsList.appendChild(li);
    });

    // Show results section
    showSection('results');
}

function formatCurrency(value) {
    return '$' + parseFloat(value).toLocaleString('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

function downloadFile() {
    if (!outputFilename) {
        showError('No file available for download.');
        return;
    }

    // Create download link
    const link = document.createElement('a');
    link.href = `/download/${outputFilename}`;
    link.download = outputFilename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function showSection(section) {
    // Hide all sections
    document.querySelector('.upload-section').style.display = 'none';
    settingsSection.style.display = 'none';
    progressSection.style.display = 'none';
    resultsSection.style.display = 'none';
    errorSection.style.display = 'none';

    // Show requested section
    switch (section) {
        case 'upload':
            document.querySelector('.upload-section').style.display = 'block';
            break;
        case 'settings':
            document.querySelector('.upload-section').style.display = 'block';
            settingsSection.style.display = 'block';
            break;
        case 'progress':
            progressSection.style.display = 'block';
            break;
        case 'results':
            resultsSection.style.display = 'block';
            break;
        case 'error':
            errorSection.style.display = 'block';
            break;
    }
}

function updateProgress(message) {
    progressText.textContent = message;
}

function showError(message) {
    errorMessage.textContent = message;
    showSection('error');
}

function resetForm() {
    // Reset state
    selectedFile = null;
    outputFilename = null;

    // Reset file input
    fileInput.value = '';

    // Reset file info
    fileInfo.textContent = '';
    fileInfo.classList.remove('success');

    // Reset settings
    document.getElementById('min-clicks').value = '3';
    document.getElementById('min-spend').value = '0.01';

    // Show upload section
    showSection('upload');
}
