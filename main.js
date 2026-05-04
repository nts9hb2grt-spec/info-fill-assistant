const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const mammoth = require('mammoth');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1000,
        height: 700,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });
    mainWindow.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

// 打开Excel文件
ipcMain.handle('open-excel', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
        properties: ['openFile']
    });
    if (result.canceled) return null;
    return result.filePaths[0];
});

// 打开Word文件
ipcMain.handle('open-word', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        filters: [{ name: 'Word', extensions: ['docx', 'doc'] }],
        properties: ['openFile']
    });
    if (result.canceled) return null;
    return result.filePaths[0];
});

// 处理文件
ipcMain.handle('process-files', async (event, { excelPath, wordPath }) => {
    try {
        // 读取Word文档
        const wordBuffer = fs.readFileSync(wordPath);
        const wordResult = await mammoth.extractRawText({ buffer: wordBuffer });
        let allText = wordResult.text;

        // 合并段落文本
        const paragraphs = allText.split(/\n+/).filter(p => p.trim());

        // 读取Excel
        const workbook = XLSX.readFile(excelPath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        let data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        // 获取第一列作为关键词，在第二列填充结果
        const results = [];
        let filledCount = 0;

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const keyword = String(row[0] || '').trim();

            if (keyword && keyword !== 'undefined' && keyword !== '') {
                const result = findContent(keyword, paragraphs, allText);
                results.push({ row: i, keyword, result });
                if (result !== '未找到') filledCount++;
            } else {
                results.push({ row: i, keyword: '', result: '' });
            }
        }

        // 写回Excel第二列
        for (const item of results) {
            if (item.row < data.length) {
                data[item.row][1] = item.result;
            }
        }

        // 保存结果
        const outputPath = excelPath.replace('.xlsx', '_结果.xlsx');
        const newSheet = XLSX.utils.aoa_to_sheet(data);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
        XLSX.writeFile(newWorkbook, outputPath);

        return { success: true, outputPath, totalCount: results.filter(r => r.keyword).length, filledCount };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

function findContent(keyword, paragraphs, tableText) {
    // 优先在段落中查找
    for (let i = 0; i < paragraphs.length; i++) {
        if (paragraphs[i].includes(keyword)) {
            let context = paragraphs[i];
            // 包含后面几段作为上下文
            for (let j = 1; j <= 2 && i + j < paragraphs.length; j++) {
                context += '\n' + paragraphs[i + j];
            }
            return context.trim();
        }
    }

    // 如果段落没找到，在表格文本中查找
    if (tableText && tableText.includes(keyword)) {
        const lines = tableText.split(/\n+/);
        for (const line of lines) {
            if (line.includes(keyword)) {
                return line.trim();
            }
        }
        return tableText.substring(0, 500) + '...';
    }

    return '未找到';
}

// 预览Excel内容
ipcMain.handle('preview-excel', async (event, filePath) => {
    try {
        const workbook = XLSX.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
        return { success: true, data: data.slice(0, 50), totalRows: data.length };
    } catch (error) {
        return { success: false, error: error.message };
    }
});

// 预览Word内容
ipcMain.handle('preview-word', async (event, filePath) => {
    try {
        const buffer = fs.readFileSync(filePath);
        const result = await mammoth.extractRawText({ buffer });
        return { success: true, text: result.value.substring(0, 5000) };
    } catch (error) {
        return { success: false, error: error.message };
    }
});