import './styles.css';

// Configuration - Update this for production
const API_BASE_URL = process.env.NODE_ENV === 'production' 
  ? 'https://backend-rewordLy.onrender.com' 
  : 'http://localhost:5000';

// DOM Elements
const selectedTextArea = document.getElementById('selectedText');
const instructionsInput = document.getElementById('instructions');
const rewordBtn = document.getElementById('rewordBtn');
const composeBtn = document.getElementById('composeBtn');
const outputSection = document.getElementById('outputSection');
const outputTextArea = document.getElementById('outputText');
const copyBtn = document.getElementById('copyBtn');
const replaceBtn = document.getElementById('replaceBtn');
const statusDiv = document.getElementById('status');

// State
let currentGeneratedText = '';

// Initialize Office.js
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        console.log('Rewordly Add-in loaded in Outlook');
        initializeAddIn();
    }
});

function initializeAddIn() {
    // Set up event listeners
    rewordBtn.addEventListener('click', handleReword);
    composeBtn.addEventListener('click', handleCompose);
    copyBtn.addEventListener('click', handleCopy);
    replaceBtn.addEventListener('click', handleReplace);
    
    // Load selected text on initialization
    loadSelectedText();
    
    // Set up periodic checking for selected text changes
    setInterval(loadSelectedText, 1000);
}

async function loadSelectedText() {
    try {
        const item = Office.context.mailbox.item;
        
        if (item && item.body) {
            item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const bodyText = result.value;
                    
                    if (bodyText && bodyText.trim()) {
                        selectedTextArea.value = bodyText.substring(0, 500) + (bodyText.length > 500 ? '...' : '');
                    }
                }
            });
        }
    } catch (error) {
        console.error('Error loading selected text:', error);
    }
}

async function handleReword() {
    const selectedText = selectedTextArea.value.trim();
    const instructions = instructionsInput.value.trim();
    
    if (!selectedText) {
        showStatus('Please select or enter text to reword', 'error');
        return;
    }
    
    if (!instructions) {
        showStatus('Please enter tone instructions', 'error');
        return;
    }
    
    setLoading(true);
    
    try {
        const response = await fetch(`${API_BASE_URL}/api/reword`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                selectedText: selectedText,
                toneInstructions: instructions
            })
        });
        
        const data = await response.json();
        
        if (response.ok && data.success) {
            currentGeneratedText = data.rewording_text;
            outputTextArea.value = currentGeneratedText;
            outputSection.style.display = 'block';
            showStatus('Text rewording completed successfully!', 'success');
        } else {
            throw new Error(data.error || 'Failed to reword text');
        }
    } catch (error) {
        console.error('Error rewording text:', error);
        showStatus(`Error: ${error.message}`, 'error');
    } finally {
        setLoading(false);
    }
}

async function handleCompose() {
    const instructions = instructionsInput.value.trim();
    
    if (!instructions) {
        showStatus('Please enter composition instructions', 'error');
        return;
    }
    
    setLoading(true);
    
    try {
        const response = await fetch(`${API_BASE_URL}/api/compose`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                compositionInstructions: instructions
            })
        });
        
        const data = await response.json();
        
        if (response.ok && data.success) {
            currentGeneratedText = data.composed_email;
            outputTextArea.value = currentGeneratedText;
            outputSection.style.display = 'block';
            showStatus('Email composition completed successfully!', 'success');
        } else {
            throw new Error(data.error || 'Failed to compose email');
        }
    } catch (error) {
        console.error('Error composing email:', error);
        showStatus(`Error: ${error.message}`, 'error');
    } finally {
        setLoading(false);
    }
}

async function handleCopy() {
    try {
        await navigator.clipboard.writeText(currentGeneratedText);
        showStatus('Text copied to clipboard!', 'success');
    } catch (error) {
        console.error('Error copying text:', error);
        showStatus('Failed to copy text', 'error');
    }
}

async function handleReplace() {
    try {
        const item = Office.context.mailbox.item;
        
        if (item && item.body) {
            item.body.setAsync(currentGeneratedText, { coercionType: Office.CoercionType.Text }, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showStatus('Text replaced successfully!', 'success');
                    outputSection.style.display = 'none';
                    outputTextArea.value = '';
                    currentGeneratedText = '';
                } else {
                    showStatus('Failed to replace text', 'error');
                }
            });
        } else {
            showStatus('No active email item found', 'error');
        }
    } catch (error) {
        console.error('Error replacing text:', error);
        showStatus('Failed to replace text', 'error');
    }
}

function setLoading(isLoading) {
    rewordBtn.disabled = isLoading;
    composeBtn.disabled = isLoading;
    
    if (isLoading) {
        rewordBtn.textContent = '⏳ Processing...';
        composeBtn.textContent = '⏳ Processing...';
    } else {
        rewordBtn.textContent = '🔄 Reword';
        composeBtn.textContent = '✍️ Compose';
    }
}

function showStatus(message, type) {
    statusDiv.textContent = message;
    statusDiv.className = `status ${type}`;
    
    setTimeout(() => {
        statusDiv.textContent = '';
        statusDiv.className = 'status';
    }, 5000);
}

// Make selected text area editable for manual input
selectedTextArea.addEventListener('click', function() {
    if (this.readOnly) {
        this.readOnly = false;
        this.placeholder = 'Enter or paste text here...';
        this.style.backgroundColor = '#fff';
    }
});

// Handle Enter key in instructions input
instructionsInput.addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
        handleReword();
    }
}); 