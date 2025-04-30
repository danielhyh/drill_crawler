/**
 * 台湾地质钻孔数据批量下载工具 - 高性能版
 * 优化重点：多线程并发、自适应延迟、文件存储精简、错误恢复机制
 */
const { Worker, isMainThread, parentPort, workerData } = require('worker_threads');
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const iconv = require('iconv-lite');
const os = require('os');

// 创建配置目录
const CONFIG_DIR = './drill_config';
const DATA_DIR = './drill_data';
// 添加全局文件路径定义
const STATUS_FILE = path.join(CONFIG_DIR, 'processing_status.json');
const ERROR_LOG_FILE = path.join(CONFIG_DIR, 'error_log.txt');
const PROJECTS_FILE = path.join(CONFIG_DIR, 'projects.json');
// 配置参数
const CONFIG = {
    // 输出目录
    outputDir: DATA_DIR,

    // 浏览器配置
    browser: {
        headless: false,  // 无头模式提高性能
        proxy: {
            server: 'http://127.0.0.1:7897'  // 代理地址
        }
    },

    // 搜索区域（WKT格式多边形）
    searchArea: 'POLYGON((12705872.421750566 2429796.7027784917,12705872.421750566 3003991.6592567354,14010805.368635094 3003991.6592567354,14010805.368635094 2429796.7027784917,12705872.421750566 2429796.7027784917))',

    // 自适应延迟参数
    minDelay: 100,      // 最小延迟(ms)
    maxDelay: 2000,     // 最大延迟(ms)
    initialDelay: 500,  // 初始延迟(ms)

    // 重试参数
    maxRetries: 3,      // 最大重试次数

    // 并发参数
    workerCount: Math.max(os.cpus().length - 1, 1), // 工作线程数（CPU核心数-1）
    projectsPerWorker: 50, // 每个工作线程处理的项目数

    // 处理范围
    resumeable: true,    // 是否支持恢复处理

    // 保存选项
    saveBaseData: true,
    saveTestData: true,
    saveChartImages: true,
    saveCoreImages: true,

    // 日志级别
    logLevel: 'debug'    // 'debug', 'info', 'error'
};

// 延迟函数（自适应延迟）
let currentDelay = CONFIG.initialDelay;
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms || currentDelay));

// 日志函数
const logger = (message, level = 'info') => {
    if ((level === 'error') || (CONFIG.logLevel === 'debug') ||
        (CONFIG.logLevel === 'info' && level !== 'debug')) {
        const timestamp = new Date().toISOString();
        const logMessage = `[${timestamp}][${level.toUpperCase()}] ${message}`;

        if (level === 'error') {
            fs.appendFileSync(ERROR_LOG_FILE, logMessage + '\n');
        }

        if (level === 'error' || CONFIG.logLevel === 'debug') {
            console.log(logMessage);
        }
    }
};

// Unicode转中文函数
function decodeUnicode(str) {
    if (!str) return '';
    return str.replace(/\\u([0-9a-fA-F]{4})/g, (match, hex) => {
        return String.fromCharCode(parseInt(hex, 16));
    }).replace(/&#([0-9]+);/g, (match, dec) => {
        return String.fromCharCode(dec);
    }).replace(/\u([0-9a-fA-F]{4})/g, (match, hex) => {
        return String.fromCharCode(parseInt(hex, 16));
    });
}

// 确保目录存在
function ensureDirectoryExists(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
        logger(`创建目录: ${dirPath}`, 'debug');
    }
    return dirPath;
}

// 安全的文件名生成函数
function getSafeFileName(name) {
    if (!name) return 'unknown';
    return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
        .replace(/\s+/g, '_')
        .replace(/__+/g, '_')
        .trim();
}

// 保存CSV文件
function saveCSV(data, filePath) {
    try {
        if (!data || (Array.isArray(data) && data.length === 0)) {
            return false;
        }

        // 确保目录存在
        ensureDirectoryExists(path.dirname(filePath));

        // 初始化 CSV 内容（不带 BOM 标记）
        let csvContent = '';

        // 如果是数组数据
        if (Array.isArray(data)) {
            if (typeof data[0] === 'object') {
                // 获取表头
                const headers = Object.keys(data[0]);

                // 创建 CSV 内容
                csvContent += headers.join(',') + '\n';

                // 添加数据行
                data.forEach(row => {
                    const csvRow = headers.map(header => {
                        let value = row[header] !== undefined ? row[header] : '';
                        // 转义 CSV 字段
                        if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) {
                            value = `"${value.replace(/"/g, '""')}"`;
                        }
                        return value;
                    }).join(',');
                    csvContent += csvRow + '\n';
                });
            }
        }
        // 如果是对象数据
        else if (typeof data === 'object' && !Array.isArray(data)) {
            csvContent += 'key,value\n';
            Object.entries(data).forEach(([key, value]) => {
                // 转义 CSV 字段
                if (typeof key === 'string' && (key.includes(',') || key.includes('"') || key.includes('\n'))) {
                    key = `"${key.replace(/"/g, '""')}"`;
                }
                if (typeof value === 'string' && (value.includes(',') || value.includes('"') || value.includes('\n'))) {
                    value = `"${value.replace(/"/g, '""')}"`;
                } else if (typeof value === 'object') {
                    value = `"${JSON.stringify(value).replace(/"/g, '""')}"`;
                }
                csvContent += `${key},${value}\n`;
            });
        }

        // 使用 iconv-lite 转换为 GBK 编码（解决 Excel 中文问题）
        const buffer = iconv.encode(csvContent, 'gbk');
        fs.writeFileSync(filePath, buffer);

        return true;
    } catch (error) {
        logger(`保存CSV文件失败: ${error.message}`, 'error');
        return false;
    }
}

// 保存图片数据函数
function saveBase64Image(base64Data, outputPath) {
    try {
        // 从base64字符串中提取数据部分
        const matches = base64Data.match(/data:image\/([a-zA-Z]+);base64,([^"]+)/);
        if (!matches || matches.length < 3) {
            logger(`无效的Base64图片数据`, 'debug');
            return false;
        }

        const imageData = matches[2];
        ensureDirectoryExists(path.dirname(outputPath));
        fs.writeFileSync(outputPath, Buffer.from(imageData, 'base64'));
        return true;
    } catch (error) {
        logger(`保存图片失败: ${error.message}`, 'error');
        return false;
    }
}

// 解析HTML表格数据
function parseHtmlTable(htmlString) {
    try {
        // 提取表格正则表达式
        const tableRegex = /<table[^>]*>([\s\S]*?)<\/table>/gi;
        const rowRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
        const cellRegex = /<t[hd][^>]*>([\s\S]*?)<\/t[hd]>/gi;

        let tables = [];

        // 提取所有表格
        let tableMatch;
        while ((tableMatch = tableRegex.exec(htmlString)) !== null) {
            const tableContent = tableMatch[1];
            const rows = [];
            const headers = [];
            let firstRow = true;

            // 提取表格行
            let rowMatch;
            while ((rowMatch = rowRegex.exec(tableContent)) !== null) {
                const rowContent = rowMatch[1];
                const cells = [];

                // 提取单元格内容
                let cellMatch;
                while ((cellMatch = cellRegex.exec(rowContent)) !== null) {
                    let cellContent = cellMatch[1].trim();
                    // 清理HTML
                    cellContent = cellContent.replace(/<[^>]*>/g, '')
                        .replace(/&nbsp;/g, ' ')
                        .replace(/&amp;/g, '&')
                        .replace(/&lt;/g, '<')
                        .replace(/&gt;/g, '>')
                        .replace(/&quot;/g, '"')
                        .replace(/&#39;/g, "'");
                    // 解码Unicode
                    cellContent = decodeUnicode(cellContent);
                    cells.push(cellContent);
                }

                if (firstRow) {
                    headers.push(...cells);
                    firstRow = false;
                } else if (cells.length > 0) {
                    if (cells.length === 2) {
                        // 键值对形式
                        rows.push({ key: cells[0], value: cells[1] });
                    } else {
                        // 多列数据
                        const rowData = {};
                        for (let i = 0; i < Math.min(headers.length, cells.length); i++) {
                            if (headers[i] && headers[i].trim() !== '') {
                                rowData[headers[i]] = cells[i];
                            }
                        }
                        rows.push(rowData);
                    }
                }
            }

            tables.push({
                headers: headers,
                rows: rows
            });
        }

        // 如果只有一个表格，直接返回其行数据
        if (tables.length === 1) {
            // 如果所有行都是键值对，转换为对象
            const table = tables[0];
            if (table.rows.length > 0 && table.rows.every(row => 'key' in row && 'value' in row)) {
                const resultObj = {};
                table.rows.forEach(row => {
                    resultObj[row.key] = row.value;
                });
                return resultObj;
            }
            return table.rows;
        }

        return tables;
    } catch (error) {
        logger(`解析HTML表格失败: ${error.message}`, 'error');
        return null;
    }
}

// 解析钻孔测试数据
function parseTestData(htmlString) {
    try {
        const results = [];
        // 查找所有测试面板
        const testBlockRegex = /<div class="panel panel-default">([\s\S]*?)<\/div>\s*<\/div>\s*<\/div>/gi;

        let testMatch;
        while ((testMatch = testBlockRegex.exec(htmlString)) !== null) {
            const testBlock = testMatch[1];

            // 提取主表（基本信息表）
            const mainTableRegex = /<table class='table table-bordered m-0'>([\s\S]*?)<\/table>/i;
            const mainTableMatch = testBlock.match(mainTableRegex);
            if (!mainTableMatch) continue;

            // 提取主表表头
            const mainHeadersRegex = /<tr class='bg-info[^>]*?>([\s\S]*?)<\/tr>/i;
            const mainHeadersMatch = mainTableMatch[1].match(mainHeadersRegex);
            if (!mainHeadersMatch) continue;

            // 解析主表表头
            const mainHeaders = [];
            const mainHeaderCellRegex = /<td[^>]*?>([\s\S]*?)<\/td>/gi;
            let headerMatch;
            while ((headerMatch = mainHeaderCellRegex.exec(mainHeadersMatch[1])) !== null) {
                mainHeaders.push(decodeUnicode(headerMatch[1].replace(/<[^>]*>/g, '').trim()));
            }

            // 提取主表数据行
            const mainDataRegex = /<tbody>([\s\S]*?)<\/tbody>/i;
            const mainDataMatch = mainTableMatch[1].match(mainDataRegex);
            if (!mainDataMatch) continue;

            // 解析主表数据行
            const mainData = {};
            const mainDataRowRegex = /<tr>([\s\S]*?)<\/tr>/i;
            const mainDataRowMatch = mainDataMatch[1].match(mainDataRowRegex);
            if (mainDataRowMatch) {
                const cellRegex = /<td[^>]*?>([\s\S]*?)<\/td>/gi;
                let cellIndex = 0;
                let cellMatch;
                while ((cellMatch = cellRegex.exec(mainDataRowMatch[1])) !== null && cellIndex < mainHeaders.length) {
                    mainData[mainHeaders[cellIndex]] = decodeUnicode(cellMatch[1].replace(/<[^>]*>/g, '').trim());
                    cellIndex++;
                }
            }

            // 提取子表（测试数据表）
            const subTableRegex = /<table class='table table-bordered'>([\s\S]*?)<\/table>/i;
            const subTableMatch = testBlock.match(subTableRegex);
            if (!subTableMatch) continue;

            // 提取子表表头
            const subHeadersRegex = /<tr class='bg-info'>([\s\S]*?)<\/tr>/i;
            const subHeadersMatch = subTableMatch[1].match(subHeadersRegex);
            if (!subHeadersMatch) continue;

            // 解析子表表头
            const subHeaders = [];
            const subHeaderCellRegex = /<th[^>]*?>([\s\S]*?)<\/th>/gi;
            let subHeaderMatch;
            while ((subHeaderMatch = subHeaderCellRegex.exec(subHeadersMatch[1])) !== null) {
                // 移除<br/>标签并清理空白
                let headerText = subHeaderMatch[1].replace(/<br\s*\/?>/gi, '').replace(/<[^>]*>/g, '').trim();
                subHeaders.push(decodeUnicode(headerText));
            }

            // 提取子表数据行
            const subDataRows = [];
            const subDataRowRegex = /<tr>([\s\S]*?)<\/tr>/gi;
            let subDataRowMatch;
            while ((subDataRowMatch = subDataRowRegex.exec(subTableMatch[1])) !== null) {
                const rowContent = subDataRowMatch[1].trim();
                if (rowContent === '') continue; // 跳过空行

                const rowCellRegex = /<td[^>]*?>([\s\S]*?)<\/td>/gi;
                const rowData = {};
                let cellIndex = 0;
                let rowCellMatch;

                // 检查是否是有效数据行
                const anyCellContent = rowContent.replace(/<[^>]*>/g, '').trim();
                if (anyCellContent === '') continue; // 跳过完全空的行

                while ((rowCellMatch = rowCellRegex.exec(subDataRowMatch[1])) !== null && cellIndex < subHeaders.length) {
                    rowData[subHeaders[cellIndex]] = decodeUnicode(rowCellMatch[1].replace(/<[^>]*>/g, '').trim());
                    cellIndex++;
                }

                // 确保行中至少有一个有效数据
                if (Object.keys(rowData).length > 0 && !Object.values(rowData).every(v => v === '')) {
                    subDataRows.push(rowData);
                }
            }

            // 创建完整的测试记录
            const testRecord = {
                基本信息: mainData,
                测试数据: subDataRows
            };

            results.push(testRecord);
        }

        return results;
    } catch (error) {
        console.error(`解析测试数据失败: ${error.message}`);
        return [];
    }
}
function saveTestDataToCSV(parsedTests, outputPath) {
    try {
        // CSV 内容（不带 BOM）
        let csvContent = '';

        for (let i = 0; i < parsedTests.length; i++) {
            const test = parsedTests[i];

            // 如果不是第一个测试，添加分隔行
            if (i > 0) {
                csvContent += '\n\n';
            }

            // 获取主表字段
            const basicInfo = test.基本信息;
            const mainHeaders = Object.keys(basicInfo);

            // 添加主表数据
            csvContent += mainHeaders.join(',') + '\n';
            csvContent += mainHeaders.map(header => {
                const value = basicInfo[header] || '';
                // 处理CSV特殊字符
                return value.includes(',') ? `"${value.replace(/"/g, '""')}"` : value;
            }).join(',') + '\n\n';

            // 添加子表数据
            if (test.测试数据 && test.测试数据.length > 0) {
                // 获取子表字段
                const subHeaders = Object.keys(test.测试数据[0]);

                // 添加子表表头
                csvContent += subHeaders.join(',') + '\n';

                // 添加子表数据行
                test.测试数据.forEach(row => {
                    csvContent += subHeaders.map(header => {
                        const value = row[header] || '';
                        // 处理CSV特殊字符
                        return value.includes(',') ? `"${value.replace(/"/g, '""')}"` : value;
                    }).join(',') + '\n';
                });
            }
        }

        // 转换为 GBK 编码
        const buffer = iconv.encode(csvContent, 'gbk');
        fs.writeFileSync(outputPath, buffer);

        return true;
    } catch (error) {
        logger(`保存CSV文件失败: ${error.message}`, 'error');
        return false;
    }
}
// 解析图片链接
function extractImagesFromHtml(htmlString) {
    const images = [];
    const imgRegex = /<a href="(data:image\/[^"]+)"[^>]*>/gi;

    let match;
    while ((match = imgRegex.exec(htmlString)) !== null) {
        images.push(match[1]);
    }

    return images;
}
// 从日志提取项目ID的函数
function extractCompletedProjectsFromLog(logFilePath) {
    if (!fs.existsSync(logFilePath)) {
        logger('日志文件不存在', 'error');
        return { completed: [], duplicates: [] };
    }

    const logContent = fs.readFileSync(logFilePath, 'utf8');
    const completedProjects = [];
    const duplicates = [];
    const idCounts = {}; // 用于跟踪每个ID出现的次数

    // 正则表达式匹配"处理项目 ID: XXXX"格式的行
    const regex = /处理项目\s+ID:\s+(\d+)/g;
    let match;

    while ((match = regex.exec(logContent)) !== null) {
        // 提取项目ID并将其转换为数字类型
        const projectId = parseInt(match[1], 10);

        // 跟踪ID出现次数
        if (!idCounts[projectId]) {
            idCounts[projectId] = 1;
            completedProjects.push(projectId); // 首次出现，添加到完成列表
        } else {
            idCounts[projectId]++;
            // 如果这是第二次出现，添加到重复列表
            if (idCounts[projectId] === 2) {
                duplicates.push(projectId);
            }
        }
    }

    logger(`从日志文件中提取了 ${completedProjects.length} 个已完成项目ID`, 'info');

    if (duplicates.length > 0) {
        logger(`检测到 ${duplicates.length} 个重复的项目ID，这些ID可能存在处理异常`, 'warn');
    }

    return { completed: completedProjects, duplicates };
}

// 创建单独的导入命令处理函数
function importFromLog(logPath) {
    logger(`开始从日志文件 ${logPath} 导入项目ID...`, 'info');

    // 确保配置目录存在
    ensureDirectoryExists(CONFIG_DIR);

    // 读取已有状态文件（如果存在）
    let processingStatus = {
        completed: [],
        failed: [],
        duplicates: [] // 添加重复ID列表
    };

    if (fs.existsSync(STATUS_FILE)) {
        try {
            const statusData = fs.readFileSync(STATUS_FILE, 'utf8');
            processingStatus = JSON.parse(statusData);
            // 确保duplicates字段存在
            processingStatus.duplicates = processingStatus.duplicates || [];
            logger(`读取到现有状态文件，已有 ${processingStatus.completed.length} 个已完成项目和 ${processingStatus.failed.length} 个失败项目`, 'info');
        } catch (error) {
            logger(`读取状态文件失败: ${error.message}，将创建新的状态文件`, 'error');
        }
    }

    // 从日志提取ID
    const { completed, duplicates } = extractCompletedProjectsFromLog(logPath);

    let newCompletedCount = 0;

    // 添加到已完成列表（去重）
    completed.forEach(id => {
        if (!processingStatus.completed.includes(id)) {
            processingStatus.completed.push(id);
            newCompletedCount++;
        }
    });

    // 添加重复ID到特殊列表
    duplicates.forEach(id => {
        if (!processingStatus.duplicates.includes(id)) {
            processingStatus.duplicates.push(id);
        }
    });

    // 保存更新后的状态
    fs.writeFileSync(STATUS_FILE, JSON.stringify(processingStatus, null, 2));

    logger(`导入完成：新增 ${newCompletedCount} 个已完成项目，发现 ${duplicates.length} 个重复项目ID`, 'info');
    logger(`当前状态：${processingStatus.completed.length} 个已完成项目，${processingStatus.failed.length} 个失败项目，${processingStatus.duplicates.length} 个可能异常的项目`, 'info');

    // 不继续执行爬虫，直接退出
    return true;
}
// 主线程函数
async function mainThread() {
    // 设置必要的路径
    ensureDirectoryExists(CONFIG_DIR);
    ensureDirectoryExists(DATA_DIR);

    // 检查是否为导入模式
    if (process.argv.includes('--import-log')) {
        const logPath = process.argv[process.argv.indexOf('--import-log') + 1];
        if (logPath) {
            return importFromLog(logPath);
        } else {
            logger('未指定日志文件路径，使用方法：node drill_crawler.js --import-log 日志文件路径', 'error');
            return false;
        }
    }

    // 如果不是导入模式，执行正常的爬虫流程
    logger('启动台湾地质钻孔数据批量下载工具 - 高性能版', 'info');

    // 确保输出目录存在
    ensureDirectoryExists(CONFIG.outputDir);

    try {
        // 获取项目列表
        logger('获取钻探项目列表...', 'info');
        const browser = await chromium.launch(CONFIG.browser);
        const context = await browser.newContext({
            userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36 Edg/135.0.0.0'
        });
        const page = await context.newPage();

        // 访问网站
        await page.goto('https://geotech.gsmma.gov.tw/Imoeagis/Home/Map', { timeout: 60000 });
        await page.waitForLoadState('networkidle');

        // 获取项目列表
        const projectsResponse = await page.evaluate(async (params) => {
            const timestamp = Date.now();
            const response = await fetch('https://geotech.gsmma.gov.tw/Imoeagis/api/DrillProjectsJson/GetDrillProjects', {
                method: 'POST',
                headers: {
                    'Referer': `https://geotech.gsmma.gov.tw/Imoeagis/js/WebWorker/GetProjectList.js?v=${timestamp}`,
                    'Content-Type': 'application/json; charset=utf-8',
                },
                body: JSON.stringify({
                    "sType": "wkt",
                    "sWKT": params.searchArea,
                    "sKeyWord": null,
                    "sBeginDepth": null,
                    "sEndDepth": null,
                    "sPlan": "",
                    "exename": "",
                    "orgname": "",
                    "userauth": "",
                    "status": null,
                    "pagecnt": "1"
                })
            });
            return await response.json();
        }, { searchArea: CONFIG.searchArea });

        // 关闭浏览器
        await browser.close();

        // 保存项目列表
        fs.writeFileSync(PROJECTS_FILE, JSON.stringify(projectsResponse, null, 2));

        logger(`获取到 ${projectsResponse.length} 个钻探项目`, 'info');

        // 提取所有项目ID列表
        const allProjectIds = projectsResponse.map(project => project.keyid);

        // 加载或初始化处理状态
        let processingStatus = {
            completed: [],
            failed: [],
            total: allProjectIds.length
        };

        if (CONFIG.resumeable && fs.existsSync(STATUS_FILE)) {
            try {
                const statusData = fs.readFileSync(STATUS_FILE, 'utf8');
                processingStatus = JSON.parse(statusData);
                // 确保 total 字段存在并正确
                processingStatus.total = allProjectIds.length;
                logger(`从上次中断恢复处理，已完成 ${processingStatus.completed.length}/${processingStatus.total} 个项目，失败 ${processingStatus.failed.length} 个项目`, 'info');
            } catch (error) {
                logger(`读取状态文件失败，将重新开始处理: ${error.message}`, 'error');
            }
        }
        // 动态计算未完成项目
        const pendingProjectIds = allProjectIds.filter(id =>
            !processingStatus.completed.includes(id) &&
            !processingStatus.failed.includes(id)
        );

        // 找出这些ID对应的完整项目数据
        const pendingProjects = projectsResponse.filter(project =>
            pendingProjectIds.includes(project.keyid)
        );

        logger(`待处理项目: ${pendingProjects.length}/${allProjectIds.length} 个`, 'info');

        // 如果所有项目已处理完成，直接退出
        if (pendingProjects.length === 0) {
            logger('所有项目已处理完成，无需再次运行', 'info');
            return;
        }

        // 计算工作批次数
        const batchSize = CONFIG.workerCount * CONFIG.projectsPerWorker;
        const batchCount = Math.ceil(pendingProjects.length / batchSize);

        // 处理每个批次
        for (let batchIndex = 0; batchIndex < batchCount; batchIndex++) {
            logger(`处理批次 ${batchIndex + 1}/${batchCount}`, 'info');

            // 获取当前批次的项目
            const startIndex = batchIndex * batchSize;
            const endIndex = Math.min(startIndex + batchSize, pendingProjects.length);
            const batchProjects = pendingProjects.slice(startIndex, endIndex);

            // 按工作线程数量分配项目
            const projectChunks = [];
            for (let i = 0; i < CONFIG.workerCount; i++) {
                const chunkStart = i * CONFIG.projectsPerWorker;
                const chunkEnd = Math.min(chunkStart + CONFIG.projectsPerWorker, batchProjects.length);

                if (chunkStart < batchProjects.length) {
                    projectChunks.push(batchProjects.slice(chunkStart, chunkEnd));
                }
            }

            // 启动工作线程
            const workers = [];
            const workerPromises = projectChunks.map((chunk, index) => {
                return new Promise((resolve, reject) => {
                    const worker = new Worker(__filename, {
                        workerData: {
                            config: CONFIG,
                            projects: chunk,
                            workerId: index
                        }
                    });

                    // 处理工作线程消息
                    worker.on('message', (message) => {
                        if (message.type === 'complete') {
                            // 工作线程完成所有任务
                            resolve(message.results);
                        } else if (message.type === 'project_completed') {
                            // 单个项目完成，立即更新状态文件
                            if (!processingStatus.completed.includes(message.projectId)) {
                                processingStatus.completed.push(message.projectId);
                                // 保存当前处理状态
                                fs.writeFileSync(STATUS_FILE, JSON.stringify(processingStatus));
                                logger(`项目 ${message.projectId} 处理完成，进度: ${processingStatus.completed.length}/${processingStatus.total}`, 'info');
                            }
                        } else if (message.type === 'project_failed') {
                            // 单个项目失败，立即更新状态文件
                            if (!processingStatus.failed.includes(message.projectId)) {
                                processingStatus.failed.push(message.projectId);
                                // 保存当前处理状态
                                fs.writeFileSync(STATUS_FILE, JSON.stringify(processingStatus));
                                logger(`项目 ${message.projectId} 处理失败，已失败: ${processingStatus.failed.length} 个`, 'info');
                            }
                        }
                    });

                    worker.on('error', (err) => {
                        logger(`工作线程 ${index} 错误: ${err.message}`, 'error');
                        reject(err);
                    });

                    worker.on('exit', (code) => {
                        if (code !== 0) {
                            reject(new Error(`工作线程 ${index} 异常退出，退出码: ${code}`));
                        }
                    });

                    workers.push(worker);
                });
            });

            // 等待所有工作线程完成
            const results = await Promise.allSettled(workerPromises);

            // 确保状态文件准确，合并工作线程结果
            results.forEach(result => {
                if (result.status === 'fulfilled') {
                    const { completed, failed } = result.value;

                    // 合并已完成项目
                    completed.forEach(id => {
                        if (!processingStatus.completed.includes(id)) {
                            processingStatus.completed.push(id);
                        }
                    });

                    // 合并失败项目
                    failed.forEach(id => {
                        if (!processingStatus.failed.includes(id)) {
                            processingStatus.failed.push(id);
                        }
                    });
                }
            });

            // 保存更新后的状态
            fs.writeFileSync(STATUS_FILE, JSON.stringify(processingStatus));
            logger(`批次 ${batchIndex + 1} 处理完成，总进度: ${processingStatus.completed.length}/${processingStatus.total}`, 'info');
        }

        logger(`所有项目处理完成，成功: ${processingStatus.completed.length}，失败: ${processingStatus.failed.length}`, 'info');

    } catch (error) {
        logger(`主线程执行失败: ${error.message}`, 'error');
        logger(error.stack, 'error');
    }
}

// 工作线程中添加项目完成/失败的实时通知
async function workerThread() {
    const { config, projects, workerId } = workerData;

    const completedProjects = [];
    const failedProjects = [];

    try {
        // 启动浏览器（保持原有代码）
        const browser = await chromium.launch(config.browser);
        const context = await browser.newContext({/* 保持原有代码 */});
        const page = await context.newPage();
        await ensureSessionValid(page);

        // 自适应延迟控制
        let currentDelay = config.initialDelay;
        let consecutiveErrors = 0;

        // 处理每个项目
        for (let i = 0; i < projects.length; i++) {
            const project = projects[i];

            try {
                const success = await processProject(page, project, config, currentDelay);

                if (success) {
                    completedProjects.push(project.keyid);
                    consecutiveErrors = 0;

                    // 实时向主线程报告项目完成
                    parentPort.postMessage({
                        type: 'project_completed',
                        projectId: project.keyid
                    });

                    // 降低延迟（保持原有代码）
                    currentDelay = Math.max(config.minDelay, currentDelay * 0.9);
                } else {
                    failedProjects.push(project.keyid);
                    consecutiveErrors++;

                    // 实时向主线程报告项目失败
                    parentPort.postMessage({
                        type: 'project_failed',
                        projectId: project.keyid
                    });

                    // 增加延迟（保持原有代码）
                    currentDelay = Math.min(config.maxDelay, currentDelay * 1.5);
                }

                logger(`线程 ${workerId}: 项目 ${project.keyid} 处理完成，状态: ${success ? '成功' : '失败'}`, 'debug');
            } catch (error) {
                logger(`处理项目 ${project.keyid} 失败: ${error.message}`, 'error');
                failedProjects.push(project.keyid);

                // 实时向主线程报告项目失败
                parentPort.postMessage({
                    type: 'project_failed',
                    projectId: project.keyid
                });

                consecutiveErrors++;
                currentDelay = Math.min(config.maxDelay, currentDelay * 2);
            }

            // 项目间延迟
            await delay(currentDelay);
        }

        // 关闭浏览器
        await browser.close();

        // 发送最终结果到主线程
        parentPort.postMessage({
            type: 'complete',
            results: {
                completed: completedProjects,
                failed: failedProjects
            }
        });

    } catch (error) {
        logger(`工作线程 ${workerId} 执行失败: ${error.message}`, 'error');

        // 发送结果到主线程
        parentPort.postMessage({
            type: 'complete',
            results: {
                completed: completedProjects,
                failed: [...failedProjects, ...projects.filter(p =>
                    !completedProjects.includes(p.keyid) && !failedProjects.includes(p.keyid)
                ).map(p => p.keyid)]
            }
        });
    }
}
async function ensureSessionValid(page) {
    try {
        // 访问网站首页并等待加载完成
        await page.goto('https://geotech.gsmma.gov.tw/Imoeagis/Home/Map', {
            waitUntil: 'networkidle',
            timeout: 60000
        });

        // 等待地图元素加载（关键界面元素）
        await page.waitForSelector('#map', { timeout: 30000 }).catch(() => {});

        // 确保所有脚本执行完毕
        await delay(3000);

        return true;
    } catch (error) {
        logger(`会话状态重建失败: ${error.message}`, 'error');
        return false;
    }
}

// 处理单个项目
async function processProject(page, project, config, currentDelay) {
    const projectId = project.keyid;

    // 跳过无效项目
    if (!projectId || project.drillHoleCount <= 0) {
        return true; // 视为成功处理
    }

    // 解码项目名称
    const projectName = decodeUnicode(project.projName) || '未命名项目';
    const projectHoleCount = project.drillHoleCount || 0;
    const formattedProjectName = `${projectName}(${projectHoleCount}钻孔)`;

    // 创建安全的目录名
    const safeDirName = getSafeFileName(formattedProjectName);

    // 创建项目目录
    const projectDir = path.join(config.outputDir, safeDirName);
    ensureDirectoryExists(projectDir);

    try {
        // 获取钻孔坐标数据
        let retry = 0;
        let holeResponse = null;

        while (retry < config.maxRetries) {
            try {
                holeResponse = await page.evaluate(async (params) => {
                    const timestamp = Date.now();
                    const response = await fetch(`https://geotech.gsmma.gov.tw/Imoeagis/api/DrCoordsJson/${params.projectId},,,`, {
                        method: 'GET',
                        headers: {
                            'Referer': `https://geotech.gsmma.gov.tw/Imoeagis/js/WebWorker/GetDrCoords.js?v=${timestamp}`,
                            'Content-Type': 'application/json; charset=utf-8',
                        }
                    });
                    return await response.json();
                }, { projectId });

                break;
            } catch (error) {
                retry++;
                if (retry === config.maxRetries) {
                    throw error;
                }
                await delay(currentDelay);
            }
        }

        if (!holeResponse || !Array.isArray(holeResponse)) {
            return false;
        }

        // 创建详细数据目录
        const detailsDir = path.join(projectDir, 'details');
        ensureDirectoryExists(detailsDir);

        // 处理每个钻孔
        let processedCount = 0;
        let successCount = 0;

        for (const hole of holeResponse) {
            // 最多处理前10个钻孔以加快速度
            if (processedCount >= 10) break;

            try {
                // 解析钻孔编号
                const holeNo = hole.holePointNo || '未知编号';

                // 创建安全的钻孔目录名
                const holeFolderName = `hole_${getSafeFileName(holeNo)}`;
                const holeDir = path.join(detailsDir, holeFolderName);

                // 处理钻孔数据
                const success = await processHole(page, hole, holeDir, config, currentDelay);
                if (success) {
                    successCount++;
                }
            } catch (error) {
                logger(`处理钻孔 ${hole.keyid} 时出错: ${error.message}`, 'error');
            }

            processedCount++;

            // 钻孔处理间添加短暂延迟
            await delay(currentDelay / 2);
        }

        return successCount > 0;
    } catch (error) {
        logger(`处理项目 ${projectId} 时出错: ${error.message}`, 'error');
        return false;
    }
}

// 处理单个钻孔
async function processHole(page, hole, outputDir, config, currentDelay) {
    const holeId = hole.keyid;
    const projectId = hole.projectKeyid;
    const holeNo = hole.holePointNo || '未知编号';

    // 创建钻孔目录
    ensureDirectoryExists(outputDir);

    let success = false;

    try {
        // 记录访问日志（模拟用户点击）
        await page.evaluate(async (params) => {
            try {
                await fetch('https://geotech.gsmma.gov.tw/Imoeagis/api/Syslogs', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json; charset=UTF-8',
                        'Referer': 'https://geotech.gsmma.gov.tw/Imoeagis/Home/Map'
                    },
                    body: JSON.stringify({
                        "OpType": "Hole_Click",
                        "Description": `ProjNo:${params.projectId},HoleNo:${params.holeNo}`
                    })
                });
            } catch (error) {
                // 忽略错误
            }
        }, { projectId, holeNo });

        let tasksCompleted = 0;

        // 获取基本资料
        if (config.saveBaseData) {
            const baseSuccess = await processBaseData(page, projectId, holeId, outputDir, config, currentDelay);
            if (baseSuccess) tasksCompleted++;
        }

// 获取试验资料
        if (config.saveTestData) {
            const testSuccess = await processTestData(page, projectId, holeId, outputDir, config, currentDelay);
            if (testSuccess) tasksCompleted++;
        }

        // 获取柱状图
        if (config.saveChartImages) {
            const chartSuccess = await processChartData(page, projectId, holeId, outputDir, config, currentDelay);
            if (chartSuccess) tasksCompleted++;
        }

        // 获取岩心照片
        if (config.saveCoreImages) {
            const coreSuccess = await processCoreImages(page, holeId, outputDir, config, currentDelay);
            if (coreSuccess) tasksCompleted++;
        }

        success = tasksCompleted > 0;
        return success;
    } catch (error) {
        logger(`处理钻孔 ${holeId} 时出错: ${error.message}`, 'error');
        return false;
    }
}

// 处理基本资料
async function processBaseData(page, projectId, holeId, outputDir, config, currentDelay) {
    // 获取基本资料
    const baseData = await getHoleDetail(page, 'BaseData', projectId, holeId, config, currentDelay);
    if (!baseData || !Array.isArray(baseData) || baseData.length === 0) {
        return false;
    }

    // 解析 HTML 表格数据
    const parsedData = parseHtmlTable(baseData[0]);
    if (!parsedData) {
        return false;
    }

    // 直接保存为 CSV
    return saveCSV(parsedData, path.join(outputDir, 'base_data.csv'));
}

// 处理试验资料 - CSV单工作表版本
async function processTestData(page, projectId, holeId, outputDir, config, currentDelay) {
    try {
        // 获取试验资料
        const testData = await getHoleDetail(page, 'Test', projectId, holeId, config, currentDelay);
        if (!testData || !Array.isArray(testData) || testData.length === 0) {
            return false;
        }

        // 解析测试数据
        const parsedTests = parseTestData(testData[0]);
        if (!parsedTests || parsedTests.length === 0) {
            return false;
        }

        // 直接保存到 test_data.csv，不保存中间文件
        return saveTestDataToCSV(parsedTests, path.join(outputDir, 'test_data.csv'));
    } catch (error) {
        logger(`处理试验资料时出错: ${error.message}`, 'error');
        return false;
    }
}

// 处理柱状图数据
async function processChartData(page, projectId, holeId, outputDir, config, currentDelay) {
    // 创建柱状图目录
    const chartDir = path.join(outputDir, 'chart_data');

    // 获取柱状图数据
    const chartData = await getHoleDetail(page, 'Chart', projectId, holeId, config, currentDelay);
    if (!chartData || !Array.isArray(chartData) || chartData.length === 0) {
        return false;
    }

    // 确保目录存在
    ensureDirectoryExists(chartDir);

    // 保存图片
    let savedCount = 0;

    // 处理每个chartData元素
    for (let dataIndex = 0; dataIndex < chartData.length; dataIndex++) {
        const chartDataItem = chartData[dataIndex];

        // 提取图片链接
        const images = extractImagesFromHtml(chartDataItem);
        if (images.length === 0) continue;

        // 保存图片
        for (let i = 0; i < images.length; i++) {
            const fileName = `柱状图_${dataIndex + 1}_${i + 1}.png`;
            const imagePath = path.join(chartDir, fileName);

            if (saveBase64Image(images[i], imagePath)) {
                savedCount++;
            }
        }
    }

    return savedCount > 0;
}

// 处理岩心照片
async function processCoreImages(page, holeId, outputDir, config, currentDelay) {
    try {
        // 创建岩心照片目录
        const coreDir = path.join(outputDir, 'core_data');

        // 发送API请求获取岩心照片数据
        const coreImagesResponse = await page.evaluate(async (params) => {
            const timestamp = Date.now();
            const response = await fetch('https://geotech.gsmma.gov.tw/Imoeagis/api/DrillImageJson', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json; charset=UTF-8',
                    'Referer': 'https://geotech.gsmma.gov.tw/Imoeagis/Home/Map',
                    'X-Requested-With': 'XMLHttpRequest'
                },
                body: JSON.stringify(params.holeId)
            });
            return await response.json();
        }, { holeId });

        if (!coreImagesResponse || !Array.isArray(coreImagesResponse) || coreImagesResponse.length === 0) {
            return false;
        }

        // 确保目录存在
        ensureDirectoryExists(coreDir);

        // 保存岩心照片
        let savedCount = 0;
        for (let i = 0; i < coreImagesResponse.length; i++) {
            const image = coreImagesResponse[i];

            // 生成文件名（基于深度）
            let fileName;
            if (image.topDepth !== undefined && image.bottomDepth !== undefined) {
                fileName = `深度${image.topDepth}-${image.bottomDepth}公尺.jpg`;
            } else {
                fileName = `岩心照片_${i + 1}.jpg`;
            }

            // 保存图片
            const imagePath = path.join(coreDir, fileName);
            if (image.imagePath && saveBase64Image(image.imagePath, imagePath)) {
                savedCount++;
            }
        }

        return savedCount > 0;
    } catch (error) {
        logger(`获取岩心照片失败: ${error.message}`, 'error');
        return false;
    }
}

// 获取钻孔详细信息（通用函数）- 优化版
async function getHoleDetail(page, mode, projectId, holeId, config, currentDelay) {
    // 自适应延迟
    let retryDelay = currentDelay;

    // 重试机制
    for (let retry = 0; retry < config.maxRetries; retry++) {
        try {
            const response = await page.evaluate(async (params) => {
                const timestamp = Date.now();
                const response = await fetch('https://geotech.gsmma.gov.tw/Imoeagis/api/GeoReport', {
                    method: 'POST',
                    headers: {
                        'Referer': `https://geotech.gsmma.gov.tw/Imoeagis/js/WebWorker/GetGeoReport.js?v=${timestamp}`,
                        'Content-Type': 'application/json; charset=utf-8',
                    },
                    body: JSON.stringify({
                        "Mode": params.mode,
                        "ProjectKeyId": params.projectId,
                        "KeyId": params.holeId
                    })
                });
                return await response.json();
            }, { mode, projectId, holeId });

            return response;
        } catch (error) {
            // 最后一次重试失败
            if (retry === config.maxRetries - 1) {
                throw error;
            }

            // 增加重试延迟
            retryDelay = Math.min(config.maxDelay, retryDelay * 1.5);
            await delay(retryDelay);
        }
    }

    return null;
}

// 程序入口点
if (isMainThread) {
    mainThread().catch(error => {
        logger(`程序执行失败: ${error.message}`, 'error');
    });
} else {
    workerThread().catch(error => {
        logger(`工作线程执行失败: ${error.message}`, 'error');
        process.exit(1);
    });
}
