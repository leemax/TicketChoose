const express = require('express');
const multer = require('multer');
const AdmZip = require('adm-zip');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const child_process = require('child_process');
const { createExtractorFromFile } = require('node-unrar-js'); // Removed to prevent crash

const app = express();
const PORT = 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Create necessary directories
const uploadsDir = path.join(__dirname, 'uploads');
const tempDir = path.join(__dirname, 'temp');
const outputDir = path.join(__dirname, 'output');

[uploadsDir, tempDir, outputDir].forEach(dir => {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
});

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, uploadsDir);
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        // Keep original filename encoding
        cb(null, uniqueSuffix + '-' + file.originalname);
    }
});

const upload = multer({
    storage: storage,
    limits: { fileSize: 200 * 1024 * 1024 } // 200MB limit
});

// Store processing sessions
// Each session can have multiple processed Excel files
const sessions = new Map();

// Helper function to check if command exists
function commandExists(command) {
    try {
        child_process.execSync(`which ${command}`, { stdio: 'ignore' });
        return true;
    } catch (error) {
        return false;
    }
}

// Helper function to check archive integrity
function checkArchiveIntegrity(archivePath) {
    const ext = path.extname(archivePath).toLowerCase();

    if (ext === '.zip') {
        if (commandExists('unzip')) {
            try {
                console.log('正在校验 ZIP 文件完整性...');
                // -t: test integrity
                // -q: quiet (less output)
                child_process.execSync(`unzip -t -q "${archivePath}"`, { stdio: 'ignore' });
                console.log('完整性校验通过');
                return true;
            } catch (error) {
                console.error('完整性校验失败: 文件可能已损坏或未上传完整');
                return false;
            }
        }
        // If no unzip command, skip check or try AdmZip (AdmZip.test() is not reliable for truncation detection)
        console.log('未检测到 unzip 命令，跳过系统级完整性校验');
        return true;
    }
    // Skip check for other formats or implement specific checks
    return true;
}

// Helper function to extract archive
async function extractArchive(archivePath, extractPath) {
    const ext = path.extname(archivePath).toLowerCase();

    // Ensure extract directory exists
    if (!fs.existsSync(extractPath)) {
        fs.mkdirSync(extractPath, { recursive: true });
    }

    try {
        if (ext === '.zip') {
            // Try using system unzip with GBK encoding support first (common issue with Windows zips on Linux)
            // The -O CP936 flag specifies GBK encoding
            if (commandExists('unzip')) {
                try {
                    console.log('尝试使用系统 unzip 命令解压 (尝试 CP936 编码)...');
                    // -o: overwrite
                    // -O CP936: Specify character encoding as GBK (Windows Simplified Chinese)
                    // -d: destination directory
                    child_process.execSync(`unzip -o -O CP936 "${archivePath}" -d "${extractPath}"`, { stdio: 'ignore' });
                    console.log('系统 unzip 解压成功');
                    return;
                } catch (cmdError) {
                    console.log('系统 unzip 带编码参数失败，尝试标准 unzip...');
                    try {
                        // Try standard unzip without encoding flag (for MacOS or non-patched unzip)
                        child_process.execSync(`unzip -o "${archivePath}" -d "${extractPath}"`, { stdio: 'ignore' });
                        console.log('标准 unzip 解压成功');
                        return;
                    } catch (stdError) {
                        console.log('系统 unzip 命令执行失败，回退到 node-adm-zip');
                    }
                }
            }

            // Fallback to AdmZip (Pure JS, but might have encoding issues with non-UTF8 names)
            console.log('使用 AdmZip 解压...');
            const zip = new AdmZip(archivePath);
            zip.extractAllTo(extractPath, true);
        } else if (ext === '.rar') {
            const extractor = await createExtractorFromFile({
                filepath: archivePath,
                targetPath: extractPath
            });

            [...extractor.extract().files];
        } else {
            throw new Error('Unsupported archive format');
        }
    } catch (error) {
        console.error('Extraction error:', error);
        throw error;
    }
}

// Helper function to parse Excel
// Helper function to format name (insert spaces for Chinese names)
// ... (rest of code)

// Helper function to parse Excel
// Helper function to format name (insert spaces for Chinese names)
function formatName(name) {
    if (!name) return '';
    const str = name.toString().trim();
    // Check if contains Chinese characters
    if (/[\u4e00-\u9fa5]/.test(str)) {
        return str.split('').join(' ');
    }
    return str;
}

// Helper function to parse a single Excel sheet
// Returns { mode, records } or null if sheet is invalid/empty
function parseExcelSheet(worksheet, sheetName) {
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    console.log(`\n=== 开始解析工作表: ${sheetName} ===`);
    console.log(`总行数: ${data.length}`);

    if (data.length === 0) {
        console.log(`工作表 ${sheetName} 为空，跳过`);
        return null;
    }

    // Auto-detect header row - skip title rows and find the actual column headers
    let headerRowIndex = 0;
    let headers = data[0];

    if (!headers || headers.length === 0) {
        console.log(`工作表 ${sheetName} 没有有效的表头，跳过`);
        return null;
    }

    // Check if first row looks like a title (single merged cell or very few columns)
    // Real headers should have multiple distinct columns like "姓名", "房号", "身份证" etc.
    const firstRowNonEmpty = data[0].filter(cell => cell !== undefined && cell !== null && cell !== '').length;

    console.log(`第一行非空单元格数: ${firstRowNonEmpty}`);

    if (firstRowNonEmpty <= 2) {
        // First row is likely a title, try second row
        console.log('第一行可能是标题行，尝试使用第二行作为列标题');
        headerRowIndex = 1;
        headers = data[1];
        if (!headers || headers.length === 0) {
            console.log(`工作表 ${sheetName} 第二行也没有有效表头，跳过`);
            return null;
        }
    }

    console.log(`使用第 ${headerRowIndex + 1} 行作为列标题`);

    // Room number column - fuzzy match
    const roomIndex = headers.findIndex(h => {
        if (!h) return false;
        const str = h.toString().toLowerCase();
        return str.includes('房号') || str.includes('房间') || str.includes('room');
    });

    // Name columns detection
    // We look for both "Full Name" (Chinese) and "Split Name" (Surname + Given Name)

    // 1. Find Chinese Name / Full Name column
    let chineseNameIndex = -1;
    let maxPriority = 0;

    headers.forEach((h, index) => {
        if (!h) return;
        const str = h.toString();
        const lowerStr = str.toLowerCase();
        let priority = 0;

        // Highest priority: exact matches
        if (str === '中文姓名' || str === '姓名' || str === '名字') {
            priority = 100;
        }
        // High priority: contains specific combinations
        else if (str.includes('中文姓名') || str.includes('中文名姓')) {
            priority = 90;
        }
        // Medium priority: contains "姓名" or "中文名"
        else if (str.includes('姓名') || str.includes('中文名')) {
            priority = 80;
        }

        if (priority > maxPriority) {
            chineseNameIndex = index;
            maxPriority = priority;
        }
    });

    // 2. Find Surname and Given Name columns (Pinyin/English)
    const surnameIndex = headers.findIndex(h => {
        if (!h) return false;
        const str = h.toString();
        const lowerStr = str.toLowerCase();
        return (str.includes('拼音姓') || lowerStr.includes('surname') || lowerStr.includes('last name')) && !str.includes('名');
    });

    const givenNameIndex = headers.findIndex(h => {
        if (!h) return false;
        const str = h.toString();
        const lowerStr = str.toLowerCase();
        return (str.includes('拼音名') || lowerStr.includes('given name') || lowerStr.includes('first name')) && !str.includes('姓');
    });

    // ID card column - enhanced fuzzy match with more variations
    const idCardIndex = headers.findIndex(h => {
        if (!h) return false;
        const str = h.toString();
        const lowerStr = str.toLowerCase();

        // Chinese variations - check original string (case-sensitive for Chinese)
        if (str.includes('身份证') ||
            str.includes('证件号') ||
            str.includes('证件') ||
            str.includes('证号') ||
            str.includes('护照')) {
            return true;
        }

        // English variations - case insensitive
        if ((lowerStr.includes('id') && (lowerStr.includes('card') || lowerStr.includes('number') || lowerStr.includes('no'))) ||
            lowerStr.includes('passport') ||
            lowerStr.includes('identity')) {
            return true;
        }

        return false;
    });

    console.log(`房号列索引: ${roomIndex} ${roomIndex !== -1 ? '("' + headers[roomIndex] + '")' : ''}`);
    console.log(`中文姓名列索引: ${chineseNameIndex} ${chineseNameIndex !== -1 ? '("' + headers[chineseNameIndex] + '")' : ''}`);
    console.log(`身份证列索引: ${idCardIndex} ${idCardIndex !== -1 ? '("' + headers[idCardIndex] + '")' : ''}`);

    if (chineseNameIndex === -1 && (surnameIndex === -1 || givenNameIndex === -1)) {
        console.log(`工作表 ${sheetName} 无法找到必需的列：姓名，跳过`);
        return null;
    }

    // Determine matching mode
    let matchingMode;
    if (roomIndex !== -1) {
        matchingMode = 'room-name';
        console.log('匹配模式: 房号+姓名');
    } else if (idCardIndex !== -1) {
        matchingMode = 'name-only';
        console.log('匹配模式: 仅姓名（将使用身份证号区分重名）');
    } else {
        console.log(`工作表 ${sheetName} 缺少房号或身份证列，跳过`);
        return null;
    }

    // Extract records based on matching mode
    const records = [];
    let lastRoom = null; // Track the last valid room number for merged cells

    // Start from the row after the header
    for (let i = headerRowIndex + 1; i < data.length; i++) {
        const row = data[i];

        // 1. Get Room
        let currentRoom = null;
        if (roomIndex !== -1) {
            currentRoom = row[roomIndex];
            if (currentRoom !== undefined && currentRoom !== null && currentRoom !== '') {
                currentRoom = currentRoom.toString().trim();
                lastRoom = currentRoom;
            } else if (lastRoom) {
                currentRoom = lastRoom;
            }
        }

        // 2. Get Name (with fallback and formatting)
        let currentName = '';

        // Try Chinese name first
        if (chineseNameIndex !== -1 && row[chineseNameIndex]) {
            currentName = formatName(row[chineseNameIndex]);
        }

        // If empty, try Pinyin/English name
        if (!currentName && surnameIndex !== -1 && givenNameIndex !== -1) {
            const surname = row[surnameIndex] ? row[surnameIndex].toString().trim() : '';
            const givenName = row[givenNameIndex] ? row[givenNameIndex].toString().trim() : '';
            if (surname || givenName) {
                currentName = `${surname} ${givenName}`.trim();
            }
        }

        // 3. Get ID Card
        const currentIdCard = idCardIndex !== -1 ? row[idCardIndex] : null;

        if (matchingMode === 'room-name') {
            if (currentRoom && currentName) {
                records.push({
                    room: currentRoom,
                    name: currentName,
                    idCard: currentIdCard ? currentIdCard.toString().trim() : ''
                });
            }
        } else {
            if (currentName && currentIdCard) {
                records.push({
                    name: currentName,
                    idCard: currentIdCard ? currentIdCard.toString().trim() : ''
                });
            }
        }
    }

    console.log(`解析出的记录数: ${records.length}`);
    console.log(`=== 工作表 ${sheetName} 解析完成 ===\n`);

    if (records.length === 0) {
        return null;
    }

    return {
        mode: matchingMode,
        records: records,
        sheetName: sheetName
    };
}

// Helper function to parse Excel - returns array of sheet results
function parseExcel(excelPath) {
    const workbook = XLSX.readFile(excelPath);
    const sheetNames = workbook.SheetNames;

    console.log(`\n=== 开始解析Excel文件 ===`);
    console.log(`发现 ${sheetNames.length} 个工作表: ${sheetNames.join(', ')}`);

    const results = [];

    for (const sheetName of sheetNames) {
        const worksheet = workbook.Sheets[sheetName];
        const sheetResult = parseExcelSheet(worksheet, sheetName);

        if (sheetResult) {
            results.push(sheetResult);
        }
    }

    console.log(`\n=== Excel文件解析完成，有效工作表数: ${results.length} ===\n`);

    // For backward compatibility: if only one valid sheet, return old format
    // Otherwise return array
    if (results.length === 1) {
        return results[0];
    }

    return results;
}

// Helper function to extract room and name from PDF filename
function extractInfoFromFilename(filename) {
    // Pattern: .*-(\d+)-([^-]+)--.*\.pdf
    // Reverted to strict matching for room/name separation to avoid capturing the prefix '2-' as room
    const match = filename.match(/.*?-(\d+)-([^-]+)--.*\.pdf$/i);
    if (match) {
        return {
            room: match[1].trim(),
            name: match[2].trim()
        };
    }
    return null;
}

// Helper function to normalize name for fuzzy matching (O vs 0, and remove spaces)
function normalizeName(name) {
    if (!name) return '';
    return name.toString()
        .toUpperCase()
        .replace(/\s+/g, '') // Remove all spaces
        .replace(/0/g, 'O') // Replace all zeros with letter O
        .trim();
}

// Helper function to find matching PDFs
function findMatchingPDFs(pdfDir, parseResult, session) {
    const { mode, records: excelRecords } = parseResult;
    const matchedFiles = [];
    const matchedRecords = new Set();
    const duplicates = []; // Track passengers with duplicate name matches

    console.log(`\n=== 开始匹配PDF文件 ===`);
    console.log(`匹配模式: ${mode}`);
    console.log(`Excel记录数: ${excelRecords.length}`);

    // Use cached PDF files if available, otherwise scan and cache
    let pdfFiles;
    if (session && session.pdfFiles) {
        console.log('使用缓存的PDF文件列表');
        pdfFiles = session.pdfFiles;
    } else {
        console.log(`正在扫描PDF文件目录: ${pdfDir}`);
        if (!fs.existsSync(pdfDir)) {
            console.error(`目录不存在: ${pdfDir}`);
            return { matchedFiles: [], unmatchedPassengers: [], duplicates: [] };
        }

        // Manual recursive file scanning function
        function getAllFiles(dirPath, arrayOfFiles) {
            let files = fs.readdirSync(dirPath);

            arrayOfFiles = arrayOfFiles || [];

            files.forEach(function (file) {
                const fullPath = path.join(dirPath, file);
                if (fs.statSync(fullPath).isDirectory()) {
                    arrayOfFiles = getAllFiles(fullPath, arrayOfFiles);
                } else {
                    arrayOfFiles.push(fullPath);
                }
            });

            return arrayOfFiles;
        }

        const allFiles = getAllFiles(pdfDir);
        console.log(`找到的文件总数: ${allFiles.length}`);
        if (allFiles.length > 0) {
            console.log(`前5个文件: ${allFiles.slice(0, 5).map(f => path.basename(f)).join(', ')}`);
        }

        // Build PDF index
        pdfFiles = [];
        allFiles.forEach(filePath => {
            if (filePath.toLowerCase().endsWith('.pdf')) {
                const filename = path.basename(filePath);
                const info = extractInfoFromFilename(filename);
                if (info) {
                    pdfFiles.push({
                        path: filePath,
                        filename: filename,
                        room: info.room,
                        name: info.name,
                        normalizedName: normalizeName(info.name) // Pre-calculate normalized name
                    });
                } else {
                    console.log(`⚠️ 无法解析文件名 (格式不匹配): ${filename}`);
                }
            }
        });

        // Cache the result if session is provided
        if (session) {
            session.pdfFiles = pdfFiles;
            console.log('PDF文件列表已缓存');
        }
    }

    console.log(`有效PDF文件总数: ${pdfFiles.length}`);

    if (mode === 'room-name') {
        // Original matching logic: room + name
        excelRecords.forEach((record, index) => {
            const normalizedRecordName = normalizeName(record.name);
            const matchedPdf = pdfFiles.find(pdf =>
                pdf.room === record.room && pdf.normalizedName === normalizedRecordName
            );

            if (matchedPdf) {
                console.log(`✓ 匹配成功: 房号${record.room} - ${record.name}`);
                matchedRecords.add(index);
                matchedFiles.push(matchedPdf);
            } else {
                console.log(`✗ 严格匹配失败: 房号${record.room} - ${record.name}`);

                // Fallback: Try to find by name only
                const nameMatches = pdfFiles.filter(pdf => pdf.normalizedName === normalizedRecordName);
                if (nameMatches.length > 0) {
                    console.log(`! 发现同名文件 (房号不匹配): ${nameMatches.length}个`);
                    duplicates.push({
                        name: record.name,
                        idCard: record.idCard,
                        recordIndex: index,
                        options: nameMatches.map(pdf => ({
                            filename: pdf.filename,
                            path: pdf.path,
                            room: pdf.room
                        }))
                    });
                }
            }
        });
    } else {
        // New matching logic: name only
        excelRecords.forEach((record, index) => {
            const normalizedRecordName = normalizeName(record.name);

            // Find all PDFs matching this name (fuzzy match)
            const matchingPdfs = pdfFiles.filter(pdf => pdf.normalizedName === normalizedRecordName);

            if (matchingPdfs.length === 0) {
                console.log(`✗ 未匹配: ${record.name} (身份证: ${record.idCard})`);
            } else if (matchingPdfs.length === 1) {
                // Unique match
                console.log(`✓ 匹配成功: ${record.name} (身份证: ${record.idCard})`);
                matchedRecords.add(index);
                matchedFiles.push(matchingPdfs[0]);
            } else {
                // Multiple matches - need user selection
                console.log(`⚠ 重名检测: ${record.name} (身份证: ${record.idCard}) - 找到${matchingPdfs.length}个PDF`);
                duplicates.push({
                    name: record.name,
                    idCard: record.idCard,
                    recordIndex: index,
                    options: matchingPdfs.map(pdf => ({
                        filename: pdf.filename,
                        path: pdf.path,
                        room: pdf.room
                    }))
                });
            }
        });
    }

    // Find unmatched passengers
    const unmatchedPassengers = [];
    excelRecords.forEach((record, index) => {
        if (!matchedRecords.has(index)) {
            // Check if this is a duplicate (already in duplicates list)
            const isDuplicate = duplicates.some(d => d.recordIndex === index);
            if (!isDuplicate) {
                unmatchedPassengers.push(
                    mode === 'room-name'
                        ? { room: record.room, name: record.name }
                        : { name: record.name, idCard: record.idCard }
                );
            }
        }
    });

    console.log(`匹配成功的PDF数: ${matchedFiles.length}`);
    console.log(`重名待选择数: ${duplicates.length}`);
    console.log(`未匹配的游客数: ${unmatchedPassengers.length}`);
    if (unmatchedPassengers.length > 0) {
        console.log(`未匹配的游客:`, unmatchedPassengers);
    }
    console.log(`=== 匹配完成 ===\n`);

    return {
        matchedFiles,
        unmatchedPassengers,
        duplicates,
        mode
    };
}

// Helper function to create output ZIP
function createOutputZip(matchedFiles, outputPath) {
    const zip = new AdmZip();

    matchedFiles.forEach(file => {
        zip.addLocalFile(file.path);
    });

    zip.writeZip(outputPath);
}

// Helper function to clean up old files
function cleanupOldFiles() {
    const maxAge = 24 * 60 * 60 * 1000; // 24 hours
    const now = Date.now();

    [uploadsDir, tempDir, outputDir].forEach(dir => {
        if (fs.existsSync(dir)) {
            fs.readdirSync(dir).forEach(file => {
                const filePath = path.join(dir, file);
                const stats = fs.statSync(filePath);

                if (now - stats.mtimeMs > maxAge) {
                    if (stats.isDirectory()) {
                        fs.rmSync(filePath, { recursive: true, force: true });
                    } else {
                        fs.unlinkSync(filePath);
                    }
                }
            });
        }
    });
}

// Cleanup interval
setInterval(cleanupOldFiles, 60 * 60 * 1000); // Every 1 hour

// Routes

// Upload archive
app.post('/api/upload-archive', upload.single('archive'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: '请上传压缩文件' });
        }

        const sessionId = Date.now().toString();
        const extractPath = path.join(tempDir, sessionId);

        console.log(`\n接收到压缩文件: ${req.file.originalname}`);
        console.log(`文件大小: ${(req.file.size / 1024 / 1024).toFixed(2)} MB`);
        console.log(`保存路径: ${req.file.path}`);

        // Integrity Check
        if (!checkArchiveIntegrity(req.file.path)) {
            // Cleanup
            fs.unlinkSync(req.file.path);
            return res.status(400).json({ error: '文件上传不完整或已损坏，请重新上传' });
        }

        // Extract archive
        await extractArchive(req.file.path, extractPath);

        // Store session info
        sessions.set(sessionId, {
            archivePath: req.file.path,
            extractPath: extractPath,
            createdAt: Date.now(),
            pdfFiles: null // Initialize cache
        });

        res.json({
            success: true,
            sessionId: sessionId,
            message: '压缩文件上传并解压成功'
        });

    } catch (error) {
        console.error('Archive upload error:', error);
        res.status(500).json({ error: '处理压缩文件时出错: ' + error.message });
    }
});

// Upload Excel and process
app.post('/api/upload-excel', upload.single('excel'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: '请上传Excel文件' });
        }

        const { sessionId } = req.body;
        if (!sessionId || !sessions.has(sessionId)) {
            return res.status(400).json({ error: '无效的会话ID' });
        }

        const session = sessions.get(sessionId);

        // Initialize processedFiles array if not exists
        if (!session.processedFiles) {
            session.processedFiles = [];
        }

        // Get original Excel filename without extension
        // Fix encoding: multer receives filename in latin1, need to convert to utf8
        const rawFilename = req.file.originalname;
        const originalFilename = Buffer.from(rawFilename, 'latin1').toString('utf8');
        const excelBaseName = path.basename(originalFilename, path.extname(originalFilename));

        console.log(`\n处理Excel文件: ${originalFilename}`);

        // Parse Excel - may return single object or array of sheet results
        const parseResult = parseExcel(req.file.path);

        // Normalize to array for unified processing
        const sheetResults = Array.isArray(parseResult) ? parseResult : [parseResult];

        if (sheetResults.length === 0 || (sheetResults.length === 1 && (!sheetResults[0] || sheetResults[0].records.length === 0))) {
            fs.unlinkSync(req.file.path); // Cleanup
            return res.status(400).json({ error: 'Excel文件中没有找到有效数据' });
        }

        console.log(`\n发现 ${sheetResults.length} 个有效工作表`);

        // Track all duplicates across sheets for user resolution
        const allDuplicates = [];
        const pendingSheets = [];

        // Process each sheet
        for (const sheetData of sheetResults) {
            if (!sheetData || sheetData.records.length === 0) continue;

            const sheetName = sheetData.sheetName || 'Sheet1';

            // For multi-sheet files, include sheet name in output filename
            const outputBaseName = sheetResults.length > 1
                ? `${excelBaseName}-${sheetName}`
                : excelBaseName;

            console.log(`\n处理工作表: ${sheetName}, 记录数: ${sheetData.records.length}`);

            // Find matching PDFs
            const matchResult = findMatchingPDFs(session.extractPath, sheetData, session);
            const matchedFiles = matchResult.matchedFiles;
            const unmatchedPassengers = matchResult.unmatchedPassengers;
            const duplicates = matchResult.duplicates;

            // If there are duplicates, queue them for user resolution
            if (duplicates && duplicates.length > 0) {
                const pendingId = `${sessionId}-${sheetName}-${Date.now()}`;
                session.pendingDuplicates = session.pendingDuplicates || {};
                session.pendingDuplicates[pendingId] = {
                    excelName: originalFilename,
                    excelBaseName: outputBaseName,
                    sheetName: sheetName,
                    parseResult: sheetData,
                    matchResult: matchResult,
                    excelPath: req.file.path
                };

                allDuplicates.push({
                    sheetName: sheetName,
                    pendingId: pendingId,
                    duplicates: duplicates,
                    matched: matchedFiles.length,
                    total: sheetData.records.length
                });

                pendingSheets.push(sheetName);
            } else {
                // No duplicates for this sheet, process immediately
                let outputPath = null;
                let outputFilename = null;
                let downloadId = null;

                if (matchedFiles.length > 0) {
                    outputFilename = `${outputBaseName}.zip`;
                    outputPath = path.join(outputDir, `${sessionId}-${Date.now()}-${outputFilename}`);
                    createOutputZip(matchedFiles, outputPath);
                    downloadId = path.basename(outputPath, '.zip');
                }

                // Create file record
                const fileRecord = {
                    excelName: sheetResults.length > 1
                        ? `${originalFilename} (${sheetName})`
                        : originalFilename,
                    sheetName: sheetName,
                    outputPath: outputPath,
                    outputFilename: outputFilename,
                    matchedCount: matchedFiles.length,
                    totalCount: sheetData.records.length,
                    unmatchedCount: unmatchedPassengers.length,
                    unmatchedPassengers: unmatchedPassengers,
                    processedAt: Date.now(),
                    downloadId: downloadId
                };

                session.processedFiles.push(fileRecord);

                console.log(`处理完成: ${sheetName}, 匹配 ${matchedFiles.length}/${sheetData.records.length} 个PDF`);
            }
        }

        // If there are duplicates from any sheet, return them for user resolution
        if (allDuplicates.length > 0) {
            // For simplicity, process duplicates one sheet at a time
            const firstDup = allDuplicates[0];

            return res.json({
                success: true,
                hasDuplicates: true,
                pendingId: firstDup.pendingId,
                duplicates: firstDup.duplicates,
                matched: firstDup.matched,
                total: firstDup.total,
                sheetName: firstDup.sheetName,
                pendingSheetsCount: allDuplicates.length,
                message: `工作表 "${firstDup.sheetName}" 检测到${firstDup.duplicates.length}个重名乘客，请选择正确的船票`,
                allProcessed: session.processedFiles.map(f => ({
                    excelName: f.excelName,
                    matched: f.matchedCount,
                    total: f.totalCount,
                    unmatchedCount: f.unmatchedCount,
                    unmatchedPassengers: f.unmatchedPassengers,
                    downloadUrl: f.downloadId ? `/api/download/${f.downloadId}` : null,
                    downloadFilename: f.outputFilename
                }))
            });
        }

        // No duplicates, all sheets processed
        // Cleanup Excel file
        if (fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }

        // Calculate totals
        const totalMatched = session.processedFiles.reduce((sum, f) => sum + f.matchedCount, 0);
        const totalRecords = session.processedFiles.reduce((sum, f) => sum + f.totalCount, 0);

        res.json({
            success: true,
            matched: totalMatched,
            total: totalRecords,
            excelName: originalFilename,
            sheetsProcessed: sheetResults.length,
            message: sheetResults.length > 1
                ? `成功处理 ${sheetResults.length} 个工作表，共匹配 ${totalMatched} 个PDF文件`
                : `成功匹配 ${totalMatched} 个PDF文件`,
            allProcessed: session.processedFiles.map(f => ({
                excelName: f.excelName,
                matched: f.matchedCount,
                total: f.totalCount,
                unmatchedCount: f.unmatchedCount,
                unmatchedPassengers: f.unmatchedPassengers,
                downloadUrl: f.downloadId ? `/api/download/${f.downloadId}` : null,
                downloadFilename: f.outputFilename
            }))
        });

    } catch (error) {
        console.error('Excel processing error:', error);
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        res.status(500).json({ error: '处理Excel文件时出错: ' + error.message });
    }
});

// Resolve duplicate name selections
app.post('/api/resolve-duplicates', express.json(), async (req, res) => {
    try {
        const { sessionId, pendingId, selections } = req.body;

        if (!sessionId || !sessions.has(sessionId)) {
            return res.status(400).json({ error: '无效的会话ID' });
        }

        const session = sessions.get(sessionId);

        if (!session.pendingDuplicates || !session.pendingDuplicates[pendingId]) {
            return res.status(400).json({ error: '未找到待处理的重名数据' });
        }

        const pendingData = session.pendingDuplicates[pendingId];
        const { excelName, excelBaseName, matchResult } = pendingData;

        console.log(`\n处理重名选择: ${excelName}`);
        console.log(`用户选择:`, selections);

        // Build final matched files list
        const finalMatchedFiles = [...matchResult.matchedFiles]; // Start with already matched files

        // Add user-selected PDFs for duplicates
        selections.forEach(selection => {
            const duplicate = matchResult.duplicates.find(d =>
                d.name === selection.name && d.idCard === selection.idCard
            );

            if (duplicate) {
                const selectedOption = duplicate.options.find(opt =>
                    opt.filename === selection.selectedFilename
                );

                if (selectedOption) {
                    finalMatchedFiles.push({
                        path: selectedOption.path,
                        filename: selectedOption.filename,
                        room: selectedOption.room,
                        name: duplicate.name
                    });
                    console.log(`✓ 重名已解决: ${duplicate.name} (${duplicate.idCard}) -\u003e ${selectedOption.filename}`);
                }
            }
        });

        // Create output ZIP
        const outputFilename = `${excelBaseName}.zip`;
        const outputPath = path.join(outputDir, `${sessionId}-${Date.now()}-${outputFilename}`);
        createOutputZip(finalMatchedFiles, outputPath);

        // Create file record
        const fileRecord = {
            excelName: excelName,
            outputPath: outputPath,
            outputFilename: outputFilename,
            matchedCount: finalMatchedFiles.length,
            totalCount: pendingData.parseResult.records.length,
            unmatchedCount: matchResult.unmatchedPassengers.length,
            unmatchedPassengers: matchResult.unmatchedPassengers,
            processedAt: Date.now(),
            downloadId: path.basename(outputPath, '.zip')
        };

        // Add to session's processed files
        session.processedFiles.push(fileRecord);

        // Clean up pending data
        delete session.pendingDuplicates[pendingId];
        if (pendingData.excelPath && fs.existsSync(pendingData.excelPath)) {
            fs.unlinkSync(pendingData.excelPath);
        }

        console.log(`成功处理: ${excelName}, 匹配 ${finalMatchedFiles.length}/${pendingData.parseResult.records.length} 个PDF`);

        res.json({
            success: true,
            matched: finalMatchedFiles.length,
            total: pendingData.parseResult.records.length,
            excelName: excelName,
            unmatchedPassengers: matchResult.unmatchedPassengers,
            downloadUrl: `/api/download/${fileRecord.downloadId}`,
            downloadFilename: outputFilename,
            message: `成功匹配 ${finalMatchedFiles.length} 个PDF文件${matchResult.unmatchedPassengers.length > 0 ? `，${matchResult.unmatchedPassengers.length} 位游客未找到船票` : ''}`,
            allProcessed: session.processedFiles.map(f => ({
                excelName: f.excelName,
                matched: f.matchedCount,
                total: f.totalCount,
                unmatchedCount: f.unmatchedCount,
                unmatchedPassengers: f.unmatchedPassengers,
                downloadUrl: `/api/download/${f.downloadId}`,
                downloadFilename: f.outputFilename
            }))
        });

    } catch (error) {
        console.error('Duplicate resolution error:', error);
        res.status(500).json({ error: '处理重名选择时出错: ' + error.message });
    }
});

// Download matched files
app.get('/api/download/:downloadId', (req, res) => {
    try {
        const { downloadId } = req.params;

        // Find the file across all sessions
        let fileRecord = null;
        for (const [sessionId, session] of sessions) {
            if (session.processedFiles) {
                fileRecord = session.processedFiles.find(f => f.downloadId === downloadId);
                if (fileRecord) break;
            }
        }

        if (!fileRecord) {
            return res.status(404).json({ error: '文件不存在或已过期' });
        }

        if (!fs.existsSync(fileRecord.outputPath)) {
            return res.status(404).json({ error: '文件不存在' });
        }

        res.download(fileRecord.outputPath, fileRecord.outputFilename, (err) => {
            if (err) {
                console.error('Download error:', err);
            }
        });

    } catch (error) {
        console.error('Download error:', error);
        res.status(500).json({ error: '下载文件时出错' });
    }
});

// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'ok' });
});

app.listen(PORT, () => {
    console.log(`服务器运行在 http://localhost:${PORT}`);
});
