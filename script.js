/* --- TAB SWITCHING --- */
function switchTab(tabName) {
    document.getElementById('ocrSection').style.display = 'none';
    document.getElementById('splitSection').style.display = 'none';
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));

    if (tabName === 'ocr') {
        document.getElementById('ocrSection').style.display = 'block';
        document.querySelectorAll('.tab-btn')[0].classList.add('active');
    } else {
        document.getElementById('splitSection').style.display = 'block';
        document.querySelectorAll('.tab-btn')[1].classList.add('active');
    }
}

/* --- OCR LOGIC (Existing) --- */
const imageInput = document.getElementById('imageInput');
const imagePreview = document.getElementById('imagePreview');
const convertBtn = document.getElementById('convertBtn');
const ocrResultContainer = document.getElementById('ocrResultContainer');
const ocrResultText = document.getElementById('ocrResultText');
const statusText = document.getElementById('statusText');
const progressBar = document.getElementById('progressBar');
const loadingBar = document.getElementById('loadingBar');

if(imageInput) {
    imageInput.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            imagePreview.src = URL.createObjectURL(file);
            imagePreview.style.display = 'block';
            convertBtn.disabled = false;
            ocrResultContainer.style.display = 'none';
            statusText.innerText = "";
            progressBar.style.width = '0%';
        }
    });

    convertBtn.addEventListener('click', function() {
        const file = imageInput.files[0];
        if (!file) return;

        convertBtn.disabled = true;
        convertBtn.innerText = "Processing...";
        loadingBar.style.display = 'block';
        
        Tesseract.recognize(
            file, 'eng',
            { logger: m => {
                if (m.status === 'recognizing text') {
                    progressBar.style.width = `${Math.round(m.progress * 100)}%`;
                    statusText.innerText = `Scanning: ${Math.round(m.progress * 100)}%`;
                }
            }}
        ).then(({ data: { text } }) => {
            statusText.innerText = "✅ Done!";
            convertBtn.innerText = "Extract Text";
            convertBtn.disabled = false;
            ocrResultContainer.style.display = 'block';
            ocrResultText.value = text;
        }).catch(err => {
            statusText.innerText = "Error: " + err.message;
            convertBtn.disabled = false;
        });
    });
    
    document.getElementById('copyOcrBtn').addEventListener('click', function() {
        copyToClipboard('ocrResultText', this);
    });
}

/* --- SPLIT COLUMNS LOGIC (Updated) --- */
const splitBtn = document.getElementById('splitBtn');
const splitInput = document.getElementById('splitInput');
const col1Text = document.getElementById('col1Text');
const col2Text = document.getElementById('col2Text');
const splitResultContainer = document.getElementById('splitResultContainer');

if(splitBtn) {
    splitBtn.addEventListener('click', () => {
        const rawText = splitInput.value;
        if (!rawText.trim()) return;

        let column1 = [];
        let column2 = [];

        // 1. Split into lines
        const lines = rawText.trim().split('\n');

        lines.forEach(line => {
            // 2. Remove empty space and split by whitespace
            const parts = line.trim().split(/\s+/); // Splits by space or tab
            
            if (parts.length >= 1) column1.push(parts[0]); // First item
            if (parts.length >= 2) column2.push(parts[parts.length - 1]); // Last item (Handles cases with middle spaces better)
            else column2.push(""); // Empty if no second column
        });

        // 3. Fill the boxes
        col1Text.value = column1.join('\n');
        col2Text.value = column2.join('\n');

        splitResultContainer.style.display = 'block';
    });
}

/* --- COPY HELPER --- */
function copyToClipboard(elementId, btnElement) {
    const textArea = document.getElementById(elementId);
    textArea.select();
    navigator.clipboard.writeText(textArea.value).then(() => {
        const originalText = btnElement.innerText;
        btnElement.innerText = "Copied! ✅";
        setTimeout(() => { btnElement.innerText = originalText; }, 1500);
    });
}