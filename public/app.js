// Global state
let sessionId = null;
let processedFiles = []; // Track all processed files
let pendingDuplicates = null; // Track current duplicate resolution state

// DOM Elements
const archiveInput = document.getElementById('archiveInput');
const archiveDropZone = document.getElementById('archiveDropZone');
const archiveProgress = document.getElementById('archiveProgress');
const archiveProgressFill = document.getElementById('archiveProgressFill');
const archiveProgressText = document.getElementById('archiveProgressText');
const archiveSuccess = document.getElementById('archiveSuccess');
const archiveSection = document.getElementById('archiveSection');

const excelInput = document.getElementById('excelInput');
const excelDropZone = document.getElementById('excelDropZone');
const excelProgress = document.getElementById('excelProgress');
const excelProgressFill = document.getElementById('excelProgressFill');
const excelProgressText = document.getElementById('excelProgressText');
const excelSection = document.getElementById('excelSection');

const resultsSection = document.getElementById('resultsSection');
const downloadList = document.getElementById('downloadList');
const continueBtn = document.getElementById('continueBtn');

const errorMessage = document.getElementById('errorMessage');
const errorText = document.getElementById('errorText');

// Utility Functions
function showError(message) {
    errorText.textContent = message;
    errorMessage.style.display = 'block';
    setTimeout(() => {
        errorMessage.style.display = 'none';
    }, 5000);
}

function hideError() {
    errorMessage.style.display = 'none';
}

function updateStep(stepNumber) {
    const steps = document.querySelectorAll('.step');
    steps.forEach((step, index) => {
        const num = index + 1;
        if (num < stepNumber) {
            step.classList.add('completed');
            step.classList.remove('active');
        } else if (num === stepNumber) {
            step.classList.add('active');
            step.classList.remove('completed');
        } else {
            step.classList.remove('active', 'completed');
        }
    });
}

function updateDownloadList() {
    downloadList.innerHTML = '';

    processedFiles.forEach((file, index) => {
        const item = document.createElement('div');
        item.className = 'download-item';

        // Build unmatched passengers warning if any
        let unmatchedWarning = '';
        if (file.unmatchedCount && file.unmatchedCount > 0) {
            const passengerList = file.unmatchedPassengers.map(p =>
                `<li>房号 ${p.room} - ${p.name}</li>`
            ).join('');

            unmatchedWarning = `
                <div class="unmatched-warning">
                    <div class="warning-header" onclick="toggleUnmatched(${index})">
                        <span class="warning-icon">⚠️</span>
                        <span class="warning-text">${file.unmatchedCount} 位游客未找到船票</span>
                        <span class="toggle-icon" id="toggle-${index}">▼</span>
                    </div>
                    <ul class="unmatched-list" id="unmatched-${index}" style="display: none;">
                        ${passengerList}
                    </ul>
                </div>
            `;
        }

        item.innerHTML = `
            <div class="download-info">
                <div class="download-name">${file.excelName}</div>
                <div class="download-stats">匹配: ${file.matched}/${file.total} 个PDF</div>
                ${unmatchedWarning}
            </div>
            <button class="btn btn-success btn-small" onclick="downloadFile('${file.downloadUrl}')">
                <span class="btn-icon">⬇</span>
                下载 ${file.downloadFilename}
            </button>
        `;
        downloadList.appendChild(item);
    });
}

function toggleUnmatched(index) {
    const list = document.getElementById(`unmatched-${index}`);
    const icon = document.getElementById(`toggle-${index}`);
    if (list.style.display === 'none') {
        list.style.display = 'block';
        icon.textContent = '▲';
    } else {
        list.style.display = 'none';
        icon.textContent = '▼';
    }
}

function downloadFile(url) {
    window.location.href = url;
}

// Archive Upload Handlers
archiveInput.addEventListener('change', handleArchiveUpload);

archiveDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    archiveDropZone.classList.add('drag-over');
});

archiveDropZone.addEventListener('dragleave', () => {
    archiveDropZone.classList.remove('drag-over');
});

archiveDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    archiveDropZone.classList.remove('drag-over');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        archiveInput.files = files;
        handleArchiveUpload();
    }
});

async function handleArchiveUpload() {
    const file = archiveInput.files[0];

    if (!file) return;

    // Validate file type
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['zip', 'rar'].includes(ext)) {
        showError('请上传ZIP或RAR格式的压缩文件');
        return;
    }

    // Validate file size (200MB)
    if (file.size > 200 * 1024 * 1024) {
        showError('文件大小不能超过200MB');
        return;
    }

    // Show progress
    archiveDropZone.style.display = 'none';
    archiveProgress.style.display = 'block';

    try {
        const formData = new FormData();
        formData.append('archive', file);

        const xhr = new XMLHttpRequest();

        xhr.upload.addEventListener('progress', (e) => {
            if (e.lengthComputable) {
                const percent = (e.loaded / e.total) * 100;
                archiveProgressFill.style.width = percent + '%';
                archiveProgressText.textContent = `上传中... ${Math.round(percent)}%`;
            }
        });

        xhr.addEventListener('load', () => {
            if (xhr.status === 200) {
                const response = JSON.parse(xhr.responseText);
                sessionId = response.sessionId;

                archiveProgress.style.display = 'none';
                archiveSuccess.style.display = 'block';

                // Move to next step
                setTimeout(() => {
                    updateStep(2);
                    excelSection.style.display = 'block';
                }, 1000);
            } else {
                const error = JSON.parse(xhr.responseText);
                showError(error.error || '上传失败，请重试');
                resetArchiveUpload();
            }
        });

        xhr.addEventListener('error', () => {
            showError('网络错误，请检查连接后重试');
            resetArchiveUpload();
        });

        xhr.open('POST', '/api/upload-archive');
        xhr.send(formData);

    } catch (error) {
        console.error('Upload error:', error);
        showError('上传失败: ' + error.message);
        resetArchiveUpload();
    }
}

function resetArchiveUpload() {
    archiveProgress.style.display = 'none';
    archiveDropZone.style.display = 'block';
    archiveProgressFill.style.width = '0%';
}

// Excel Upload Handlers
excelInput.addEventListener('change', handleExcelUpload);

excelDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    excelDropZone.classList.add('drag-over');
});

excelDropZone.addEventListener('dragleave', () => {
    excelDropZone.classList.remove('drag-over');
});

excelDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    excelDropZone.classList.remove('drag-over');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        excelInput.files = files;
        handleExcelUpload();
    }
});

async function handleExcelUpload() {
    const files = Array.from(excelInput.files);

    if (files.length === 0) return;

    // Validate all files first
    for (const file of files) {
        const ext = file.name.split('.').pop().toLowerCase();
        if (!['xlsx', 'xls'].includes(ext)) {
            showError(`文件 ${file.name} 不是Excel格式`);
            return;
        }
    }

    if (!sessionId) {
        showError('请先上传压缩包');
        return;
    }

    // Show progress
    excelDropZone.style.display = 'none';
    excelProgress.style.display = 'block';

    try {
        // Process files sequentially
        for (let i = 0; i < files.length; i++) {
            const file = files[i];
            const progress = ((i + 1) / files.length) * 100;

            excelProgressFill.style.width = progress + '%';
            excelProgressText.textContent = `处理中... ${i + 1}/${files.length} - ${file.name}`;

            const formData = new FormData();
            formData.append('excel', file);
            formData.append('sessionId', sessionId);

            const response = await fetch('/api/upload-excel', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (!response.ok) {
                throw new Error(`处理 ${file.name} 失败: ${result.error || '未知错误'}`);
            }

            // Check if there are duplicates that need resolution
            if (result.hasDuplicates) {
                console.log('检测到重名乘客，显示选择界面');

                // Hide progress and reset
                excelProgress.style.display = 'none';
                excelDropZone.style.display = 'block';

                // Show duplicate resolution modal
                showDuplicateModal(result.duplicates, result.pendingId, file.name);

                // Don't continue processing other files - wait for user to resolve
                return;
            }

            // Update processed files list
            if (result.allProcessed) {
                processedFiles = result.allProcessed;
            }
        }

        excelProgressFill.style.width = '100%';
        excelProgressText.textContent = `完成！已处理 ${files.length} 个文件`;

        // Show results
        setTimeout(() => {
            updateStep(3);
            excelSection.style.display = 'none';
            resultsSection.style.display = 'block';

            updateDownloadList();

            // Reset Excel input for next upload
            excelInput.value = '';
            resetExcelUpload();

        }, 800);

    } catch (error) {
        console.error('Excel processing error:', error);
        showError('处理Excel文件失败: ' + error.message);

        // Reset Excel input to allow retry
        excelInput.value = '';
        resetExcelUpload();
    }
}

function resetExcelUpload() {
    excelProgress.style.display = 'none';
    excelDropZone.style.display = 'block';
    excelProgressFill.style.width = '0%';
}

// Continue uploading more Excel files
continueBtn.addEventListener('click', () => {
    resultsSection.style.display = 'none';
    excelSection.style.display = 'block';
    updateStep(2);
});

// Duplicate Resolution Functions
function showDuplicateModal(duplicates, pendingId, excelName) {
    pendingDuplicates = { duplicates, pendingId, excelName };

    const duplicateList = document.getElementById('duplicateList');
    duplicateList.innerHTML = '';

    duplicates.forEach((duplicate, index) => {
        const item = document.createElement('div');
        item.className = 'duplicate-item';

        // Build options radio buttons
        const optionsHtml = duplicate.options.map((option, optIndex) => `
            <div class="pdf-option">
                <input 
                    type="radio" 
                    id="dup-${index}-opt-${optIndex}" 
                    name="duplicate-${index}" 
                    value="${option.filename}"
                    ${optIndex === 0 ? 'checked' : ''}
                >
                <label for="dup-${index}-opt-${optIndex}">
                    <span class="option-room">房号: ${option.room}</span>
                    <span class="option-filename">${option.filename}</span>
                </label>
            </div>
        `).join('');

        item.innerHTML = `
            <div class="duplicate-header">
                <div class="duplicate-name">
                    <span class="name-icon">👤</span>
                    <strong>${duplicate.name}</strong>
                </div>
                <div class="duplicate-idcard">
                    身份证: ${duplicate.idCard}
                </div>
            </div>
            <div class="duplicate-options">
                <p class="options-label">请选择正确的船票 (找到${duplicate.options.length}个匹配):</p>
                ${optionsHtml}
            </div>
        `;

        duplicateList.appendChild(item);
    });

    document.getElementById('duplicateModal').style.display = 'flex';
}

function closeDuplicateModal() {
    document.getElementById('duplicateModal').style.display = 'none';
    pendingDuplicates = null;
}

async function confirmDuplicates() {
    if (!pendingDuplicates) return;

    const { duplicates, pendingId } = pendingDuplicates;
    const selections = [];

    // Collect user selections
    duplicates.forEach((duplicate, index) => {
        const selectedRadio = document.querySelector(`input[name="duplicate-${index}"]:checked`);
        if (selectedRadio) {
            selections.push({
                name: duplicate.name,
                idCard: duplicate.idCard,
                selectedFilename: selectedRadio.value
            });
        }
    });

    // Validate all selections made
    if (selections.length !== duplicates.length) {
        showError('请为所有重名乘客选择船票');
        return;
    }

    // Close modal and show processing
    closeDuplicateModal();
    excelDropZone.style.display = 'none';
    excelProgress.style.display = 'block';
    excelProgressFill.style.width = '50%';
    excelProgressText.textContent = '正在处理您的选择...';

    try {
        const response = await fetch('/api/resolve-duplicates', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                sessionId: sessionId,
                pendingId: pendingId,
                selections: selections
            })
        });

        const result = await response.json();

        if (!response.ok) {
            throw new Error(result.error || '处理失败');
        }

        // Update processed files list
        if (result.allProcessed) {
            processedFiles = result.allProcessed;
        }

        excelProgressFill.style.width = '100%';
        excelProgressText.textContent = '完成！';

        // Show results
        setTimeout(() => {
            updateStep(3);
            excelSection.style.display = 'none';
            resultsSection.style.display = 'block';
            updateDownloadList();
            resetExcelUpload();
        }, 800);

    } catch (error) {
        console.error('Duplicate resolution error:', error);
        showError('处理重名选择失败: ' + error.message);
        resetExcelUpload();
    }
}

// Attach confirm button handler
document.getElementById('confirmDuplicatesBtn').addEventListener('click', confirmDuplicates);

// Initialize
updateStep(1);
