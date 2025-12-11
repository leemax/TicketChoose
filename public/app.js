// Global state
let sessionId = null;
let processedFiles = []; // Track all processed files
let pendingDuplicates = null; // Track current duplicate resolution state
let duplicateQueue = []; // Queue for files requiring manual resolution

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
                `<li>${p.room ? `房号 ${p.room}` : ''} ${p.name} ${p.idCard ? `(${p.idCard})` : ''}</li>`
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

        // Determine status and action button
        let statusHtml = '';
        let actionHtml = '';

        if (file.matched === 0) {
            // No matches found
            statusHtml = `<div class="download-stats error-text">未找到匹配的PDF文件</div>`;
            actionHtml = `
                <button class="btn btn-secondary btn-small" disabled>
                    <span class="btn-icon">⚠️</span>
                    无匹配文件
                </button>
            `;
        } else {
            // Matches found
            statusHtml = `<div class="download-stats">匹配: ${file.matched}/${file.total} 个PDF</div>`;
            actionHtml = `
                <button class="btn btn-success btn-small" onclick="downloadFile('${file.downloadUrl}')">
                    <span class="btn-icon">⬇</span>
                    下载 ${file.downloadFilename}
                </button>
            `;
        }

        item.innerHTML = `
            <div class="download-info">
                <div class="download-name">${file.excelName}</div>
                ${statusHtml}
                ${unmatchedWarning}
            </div>
            ${actionHtml}
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

    // Reset queue for this batch
    duplicateQueue = [];

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
                // If one file fails (e.g. format error), we log it but continue? 
                // Currently throwing error stops everything. Let's log and continue to be robust.
                console.error(`处理 ${file.name} 失败: ${result.error}`);
                // throw new Error(`处理 ${file.name} 失败: ${result.error || '未知错误'}`);
                // Better UX: continue but show error later? For now, let's stick to simple "log and continue" or throw?
                // Request says "don't affect other tables". So we should log and continue.
                continue;
            }

            // Check if there are duplicates that need resolution
            if (result.hasDuplicates) {
                console.log(`检测到重名乘客 (${file.name})，加入待处理队列`);

                // Add to queue instead of blocking
                duplicateQueue.push({
                    duplicates: result.duplicates,
                    pendingId: result.pendingId,
                    excelName: file.name
                });

                // IMPORTANT: Do NOT return here. Continue to next file.
            }

            // Update processed files list (if this file succeeded, or even if it's pending, others might be done)
            if (result.allProcessed) {
                processedFiles = result.allProcessed;
            }
        }

        excelProgressFill.style.width = '100%';
        excelProgressText.textContent = `批量处理完成，正在检查待人工确认项...`;

        // Check if we have queued duplicates
        if (duplicateQueue.length > 0) {
            console.log(`开始处理重名队列，共 ${duplicateQueue.length} 个文件需要确认`);
            processNextDuplicate();
        } else {
            // No duplicates, finish immediately
            completeExcelProcessing();
        }

    } catch (error) {
        console.error('Excel processing error:', error);
        showError('部分文件处理失败: ' + error.message);

        // Even if error, check if we have results to show
        if (processedFiles.length > 0) {
            completeExcelProcessing();
        } else {
            resetExcelUpload();
        }
    }
}

function processNextDuplicate() {
    if (duplicateQueue.length === 0) {
        // All done
        completeExcelProcessing();
        return;
    }

    // Get next item
    const item = duplicateQueue.shift(); // Remove from front

    // Hide progress, show modal
    excelProgress.style.display = 'none';

    // Show duplicate resolution modal for this item
    showDuplicateModal(item.duplicates, item.pendingId, item.excelName);
}

function completeExcelProcessing() {
    setTimeout(() => {
        updateStep(3);
        excelSection.style.display = 'none';
        resultsSection.style.display = 'block';

        updateDownloadList();

        // Reset Excel input for next upload
        excelInput.value = '';
        resetExcelUpload();

    }, 800);
}

function resetExcelUpload() {
    excelProgress.style.display = 'none';
    excelDropZone.style.display = 'block';
    excelProgressFill.style.width = '0%';
}

// ... (ContinueBtn handler remains same) ...

// Duplicate Resolution Functions
function showDuplicateModal(duplicates, pendingId, excelName) {
    pendingDuplicates = { duplicates, pendingId, excelName };

    // Update Modal Title/Context to show which file is being processed
    const title = document.querySelector('#duplicateModal h3') || document.querySelector('#duplicateModal .modal-header');
    if (title) {
        // You might need to add an ID to the modal title in HTML, or just prepend text
        // For now, let's assume standard layout. We can adding file info to the list container.
    }

    const duplicateList = document.getElementById('duplicateList');
    duplicateList.innerHTML = '';

    // Add file info header
    const fileHeader = document.createElement('div');
    fileHeader.className = 'duplicate-file-info';
    fileHeader.style.padding = '10px';
    fileHeader.style.backgroundColor = '#f8f9fa';
    fileHeader.style.marginBottom = '15px';
    fileHeader.style.borderRadius = '5px';
    fileHeader.innerHTML = `<strong>正在处理文件:</strong> ${excelName} <span style="color: #666; font-size: 0.9em;">(待处理剩余: ${duplicateQueue.length} 个)</span>`;
    duplicateList.appendChild(fileHeader);

    duplicates.forEach((duplicate, index) => {
        // ... (existing duplicate item generation logic) ...
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

    // If pendingDuplicates is still set, it means the user Cancelled (didn't Confirm)
    if (pendingDuplicates) {
        console.log("用户取消了当前文件的重名处理:", pendingDuplicates.excelName);
        pendingDuplicates = null;

        // Treat as "Skip current file resolution" and move to next
        // Use setTimeout to ensure UI updates finish
        setTimeout(() => {
            processNextDuplicate();
        }, 100);
    }
}

async function confirmDuplicates() {
    if (!pendingDuplicates) return;

    // Capture data needed for API call before clearing global state
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

    // Clear global state to signal "Confirmed" intention to closeDuplicateModal
    pendingDuplicates = null;

    // Close modal and show processing
    // Since pendingDuplicates is null, closeDuplicateModal WON'T trigger processNextDuplicate
    closeDuplicateModal();

    // Don't hide dropzone yet, we might show modal again.
    // Show progress overlay? Or just keep "ExcelProgress" visible but maybe update text.
    excelDropZone.style.display = 'none';
    excelProgress.style.display = 'block';
    excelProgressFill.style.width = '100%';
    excelProgressText.textContent = '保存选择并继续...';

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

        // Process next item in queue
        processNextDuplicate();

    } catch (error) {
        console.error('Duplicate resolution error:', error);
        showError('处理重名选择失败: ' + error.message);

        // If fail, reset upload UI
        resetExcelUpload();
    }
}

// Attach confirm button handler
document.getElementById('confirmDuplicatesBtn').addEventListener('click', confirmDuplicates);

// Initialize
updateStep(1);
