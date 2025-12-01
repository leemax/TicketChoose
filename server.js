const express = require('express');
const multer = require('multer');
const AdmZip = require('adm-zip');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { createExtractorFromFile } = require('node-unrar-js');

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

// Helper function to extract archive
async function extractArchive(archivePath, extractPath) {
    const ext = path.extname(archivePath).toLowerCase();

    if (ext === '.zip') {
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
}

// Helper function to parse Excel
function parseExcel(excelPath) {
    const workbook = XLSX.readFile(excelPath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    console.log('\n=== 开始解析Excel ===');
    console.log(`总行数: ${data.length}`);

    // Find column indices for 房号 and 中文姓名
    const headers = data[0];
    const roomIndex = headers.findIndex(h => h && h.toString().includes('房号'));
    const nameIndex = headers.findIndex(h => h && (h.toString().includes('中文姓名') || h.toString().includes('姓名')));

    console.log(`房号列索引: ${roomIndex}`);
    console.log(`姓名列索引: ${nameIndex}`);

    if (roomIndex === -1 || nameIndex === -1) {
        throw new Error('无法找到必需的列：房号 或 中文姓名');
    }

    // Extract room-name pairs (skip header row)
    // Handle merged cells by forward-filling room numbers
    const records = [];
    let lastRoom = null; // Track the last valid room number

    for (let i = 1; i < data.length; i++) {
        const row = data[i];

        // Get room number, use last valid room if current is empty (merged cell)
        let currentRoom = row[roomIndex];
        if (currentRoom !== undefined && currentRoom !== null && currentRoom !== '') {
            currentRoom = currentRoom.toString().trim();
            lastRoom = currentRoom; // Update last valid room
        } else if (lastRoom) {
            currentRoom = lastRoom; // Use last valid room for merged cells
        }

        const currentName = row[nameIndex];

        console.log(`行${i}: 房号=${currentRoom}, 姓名=${currentName}`);

        if (currentRoom && currentName) {
            records.push({
                room: currentRoom,
                name: currentName.toString().trim()
            });
        }
    }

    console.log(`解析出的记录数: ${records.length}`);
    console.log('=== Excel解析完成 ===\n');

    return records;
}

// Helper function to extract room and name from PDF filename
function extractInfoFromFilename(filename) {
    // Pattern: .*-(\d+)-([^-]+)--.*\.pdf
    const match = filename.match(/.*?-(\d+)-([^-]+)--.*\.pdf$/i);
    if (match) {
        return {
            room: match[1].trim(),
            name: match[2].trim()
        };
    }
    return null;
}

// Helper function to find matching PDFs
function findMatchingPDFs(pdfDir, excelRecords) {
    const allFiles = fs.readdirSync(pdfDir, { recursive: true });
    const matchedFiles = [];
    const matchedRecords = new Set(); // Track which records were matched

    console.log(`\n=== 开始匹配PDF文件 ===`);
    console.log(`Excel记录数: ${excelRecords.length}`);
    console.log(`Excel记录:`, excelRecords);
    console.log(`找到的文件总数: ${allFiles.length}`);

    let pdfCount = 0;
    allFiles.forEach(file => {
        if (file.toLowerCase().endsWith('.pdf')) {
            pdfCount++;
            const filename = path.basename(file);
            const info = extractInfoFromFilename(filename);

            console.log(`\n检查PDF: ${filename}`);
            console.log(`  提取信息:`, info);

            if (info) {
                // Check if this PDF matches any Excel record
                const matchIndex = excelRecords.findIndex(record =>
                    record.room === info.room && record.name === info.name
                );

                if (matchIndex !== -1) {
                    console.log(`  匹配结果: ✓ 匹配成功`);
                    matchedRecords.add(matchIndex); // Mark this record as matched
                    matchedFiles.push({
                        path: path.join(pdfDir, file),
                        filename: filename,
                        room: info.room,
                        name: info.name
                    });
                } else {
                    console.log(`  匹配结果: ✗ 未匹配`);
                }
            } else {
                console.log(`  ✗ 文件名格式不匹配正则表达式`);
            }
        }
    });

    // Find unmatched passengers
    const unmatchedPassengers = [];
    excelRecords.forEach((record, index) => {
        if (!matchedRecords.has(index)) {
            unmatchedPassengers.push({
                room: record.room,
                name: record.name
            });
        }
    });

    console.log(`\nPDF文件总数: ${pdfCount}`);
    console.log(`匹配成功的PDF数: ${matchedFiles.length}`);
    console.log(`未匹配的游客数: ${unmatchedPassengers.length}`);
    if (unmatchedPassengers.length > 0) {
        console.log(`未匹配的游客:`, unmatchedPassengers);
    }
    console.log(`=== 匹配完成 ===\n`);

    return {
        matchedFiles,
        unmatchedPassengers
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
    const maxAge = 3600000; // 1 hour
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
setInterval(cleanupOldFiles, 600000); // Every 10 minutes

// Routes

// Upload archive
app.post('/api/upload-archive', upload.single('archive'), async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ error: '请上传压缩文件' });
        }

        const sessionId = Date.now().toString();
        const extractPath = path.join(tempDir, sessionId);

        // Extract archive
        await extractArchive(req.file.path, extractPath);

        // Store session info
        sessions.set(sessionId, {
            archivePath: req.file.path,
            extractPath: extractPath,
            createdAt: Date.now()
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

        // Parse Excel
        const excelRecords = parseExcel(req.file.path);

        if (excelRecords.length === 0) {
            fs.unlinkSync(req.file.path); // Cleanup
            return res.status(400).json({ error: 'Excel文件中没有找到有效数据' });
        }

        // Find matching PDFs
        const matchResult = findMatchingPDFs(session.extractPath, excelRecords);
        const matchedFiles = matchResult.matchedFiles;
        const unmatchedPassengers = matchResult.unmatchedPassengers;

        if (matchedFiles.length === 0) {
            fs.unlinkSync(req.file.path); // Cleanup
            return res.json({
                success: true,
                matched: 0,
                total: excelRecords.length,
                excelName: originalFilename,
                unmatchedPassengers: unmatchedPassengers,
                message: '没有找到匹配的PDF文件'
            });
        }

        // Create output ZIP with Excel filename
        const outputFilename = `${excelBaseName}.zip`;
        const outputPath = path.join(outputDir, `${sessionId}-${Date.now()}-${outputFilename}`);
        createOutputZip(matchedFiles, outputPath);

        // Create file record
        const fileRecord = {
            excelName: originalFilename,
            outputPath: outputPath,
            outputFilename: outputFilename,
            matchedCount: matchedFiles.length,
            totalCount: excelRecords.length,
            unmatchedCount: unmatchedPassengers.length,
            unmatchedPassengers: unmatchedPassengers,
            processedAt: Date.now(),
            downloadId: path.basename(outputPath, '.zip')
        };

        // Add to session's processed files
        session.processedFiles.push(fileRecord);

        console.log(`成功处理: ${originalFilename}, 匹配 ${matchedFiles.length}/${excelRecords.length} 个PDF`);
        if (unmatchedPassengers.length > 0) {
            console.log(`⚠️  警告: ${unmatchedPassengers.length} 位游客未找到船票`);
        }

        res.json({
            success: true,
            matched: matchedFiles.length,
            total: excelRecords.length,
            excelName: originalFilename,
            unmatchedPassengers: unmatchedPassengers,
            downloadUrl: `/api/download/${fileRecord.downloadId}`,
            downloadFilename: outputFilename,
            message: `成功匹配 ${matchedFiles.length} 个PDF文件${unmatchedPassengers.length > 0 ? `，${unmatchedPassengers.length} 位游客未找到船票` : ''}`,
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

        // Cleanup Excel file
        fs.unlinkSync(req.file.path);

    } catch (error) {
        console.error('Excel processing error:', error);
        if (req.file && fs.existsSync(req.file.path)) {
            fs.unlinkSync(req.file.path);
        }
        res.status(500).json({ error: '处理Excel文件时出错: ' + error.message });
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
