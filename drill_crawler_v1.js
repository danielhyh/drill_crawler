/**
 * 台湾地质钻孔数据批量下载工具
 * 功能：自动下载钻探项目列表、钻孔坐标、基本资料、试验数据和岩心照片
 * 优化版：增强数据解析、完善目录结构、改进错误处理
 */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { execSync } = require('child_process');

// 配置参数
const CONFIG = {
    // 输出目录
    outputDir: './drill_data',

    // 浏览器配置
    browser: {
        headless: false,  // false表示显示浏览器窗口，调试时使用
        proxy: {
            server: 'http://127.0.0.1:7897'  // Clash代理地址，根据实际情况修改
        }
    },

    // 搜索区域（WKT格式多边形，默认时台湾高雄市某区域）
    // searchArea: 'POLYGON((13385362.09537984 2580873.7813668135,13385362.09537984 2589931.569218607,13405684.79027516 2589931.569218607,13405684.79027516 2580873.7813668135,13385362.09537984 2580873.7813668135))',
    // 下面这个区域应该是所有数据
    searchArea: 'POLYGON((12705872.421750566 2429796.7027784917,12705872.421750566 3003991.6592567354,14010805.368635094 3003991.6592567354,14010805.368635094 2429796.7027784917,12705872.421750566 2429796.7027784917))',

    // 请求间隔（毫秒）
    requestDelay: 2000,

    // 最大重试次数
    maxRetries: 3,

    // 每个项目最多处理的钻孔数（0表示不限制）
    maxHolesPerProject: 0,

    // 是否保存岩心照片
    saveCoreImages: true,

    // 是否保存柱状图
    saveChartImages: true,

    // 是否保存试验数据
    saveTestData: true,

    // 是否保存基本资料
    saveBaseData: true
};

// 确保输出目录存在
if (!fs.existsSync(CONFIG.outputDir)) {
    fs.mkdirSync(CONFIG.outputDir, { recursive: true });
}

// 日志函数
const logger = (message) => {
    const timestamp = new Date().toISOString();
    console.log(`[${timestamp}] ${message}`);
    fs.appendFileSync(path.join(CONFIG.outputDir, 'crawler_log.txt'), `[${timestamp}] ${message}\n`);
};

// 延迟函数
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

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

// 确保目录存在函数
function ensureDirectoryExists(dirPath) {
    if (!fs.existsSync(dirPath)) {
        fs.mkdirSync(dirPath, { recursive: true });
        logger(`创建目录: ${dirPath}`);
    }
    return dirPath;
}

// 保存图片数据函数
function saveBase64Image(base64Data, outputPath) {
    try {
        // 从base64字符串中提取数据部分
        const matches = base64Data.match(/data:image\/([a-zA-Z]+);base64,([^"]+)/);
        if (!matches || matches.length < 3) {
            logger(`无效的Base64图片数据: ${base64Data.substring(0, 50)}...`);
            return false;
        }

        const imageType = matches[1]; // jpg, png等
        const imageData = matches[2];

        // 确保输出目录存在
        const outputDir = path.dirname(outputPath);
        ensureDirectoryExists(outputDir);

        // 写入文件
        fs.writeFileSync(outputPath, Buffer.from(imageData, 'base64'));
        logger(`图片已保存至: ${outputPath}`);
        return true;
    } catch (error) {
        logger(`保存图片失败: ${error.message}`);
        return false;
    }
}

// 创建Excel文件 - 优化版
async function createExcel(data, outputPath) {
    try {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('数据');

        // 遍历数据添加到Excel
        if (Array.isArray(data)) {
            // 数组数据，每个元素作为一行
            if (data.length > 0 && typeof data[0] === 'object') {
                // 添加表头
                const headers = Object.keys(data[0]);
                worksheet.addRow(headers);

                // 添加数据行
                data.forEach(item => {
                    const row = headers.map(header => {
                        // 确保特殊值类型正确显示
                        if (item[header] === null || item[header] === undefined) return '';
                        if (typeof item[header] === 'object') return JSON.stringify(item[header]);
                        return item[header];
                    });
                    worksheet.addRow(row);
                });
            }
        } else if (typeof data === 'object') {
            // 对象数据，每个键值对作为一行
            worksheet.addRow(['键', '值']);
            Object.entries(data).forEach(([key, value]) => {
                let displayValue = value;
                if (value === null || value === undefined) displayValue = '';
                else if (typeof value === 'object') displayValue = JSON.stringify(value);
                worksheet.addRow([key, displayValue]);
            });
        }

        // 设置列宽
        worksheet.columns.forEach(column => {
            let maxLength = 0;
            column.eachCell({ includeEmpty: true }, cell => {
                if (cell.value) {
                    const length = cell.value.toString().length;
                    if (length > maxLength) {
                        maxLength = length;
                    }
                }
            });
            column.width = Math.min(Math.max(maxLength + 2, 10), 50); // 最小10，最大50
        });

        // 确保输出目录存在
        const outputDir = path.dirname(outputPath);
        ensureDirectoryExists(outputDir);

        // 保存Excel文件
        await workbook.xlsx.writeFile(outputPath);
        logger(`Excel文件已保存至: ${outputPath}`);
        return true;
    } catch (error) {
        logger(`创建Excel文件失败: ${error.message}`);
        logger(error.stack);
        return false;
    }
}

// 解析HTML表格数据 - 优化版，更好地处理子表格
function parseHtmlTable(htmlString) {
    try {
        // 正则表达式匹配表格
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
                    // 清理HTML标签
                    cellContent = cellContent.replace(/<[^>]*>/g, '');
                    // 解码HTML实体
                    cellContent = cellContent.replace(/&nbsp;/g, ' ')
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
                    // 第一行通常是表头
                    headers.push(...cells);
                    firstRow = false;
                } else {
                    // 根据表头创建数据对象
                    if (cells.length > 0) {
                        if (cells.length === 2) {
                            // 键值对形式的表格行
                            rows.push({ key: cells[0], value: cells[1] });
                        } else {
                            // 常规多列数据行
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
            }

            // 添加解析后的表格
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

        // 多个表格返回数组
        return tables;
    } catch (error) {
        logger(`解析HTML表格失败: ${error.message}`);
        logger(error.stack);
        return null;
    }
}

// 解析钻孔测试数据 - 优化版，正确处理多层结构
function parseTestData(htmlString) {
    try {
        const results = [];

        // 提取所有测试块
        const testBlockRegex = /<div class="panel panel-default">([\s\S]*?)<\/div>\s*<\/div>\s*<\/div>/gi;

        let testMatch;
        while ((testMatch = testBlockRegex.exec(htmlString)) !== null) {
            const testBlock = testMatch[1];

            // 提取试验信息
            const testInfoRegex = /<table[^>]*?class='table table-bordered m-0'[^>]*?>([\s\S]*?)<\/table>/i;
            const testInfoMatch = testBlock.match(testInfoRegex);

            if (!testInfoMatch) continue;

            // 解析试验信息表格
            const testInfoTable = parseHtmlTable(`<table>${testInfoMatch[1]}</table>`);

            if (!testInfoTable || !Array.isArray(testInfoTable) || testInfoTable.length === 0) {
                continue;
            }

            // 获取试验基本信息
            const testInfo = testInfoTable[0];

            // 提取试验数据表格
            const testDataRegex = /<table[^>]*?class='table table-bordered'[^>]*?>([\s\S]*?)<\/table>/i;
            const testDataMatch = testBlock.match(testDataRegex);

            if (!testDataMatch) continue;

            // 解析试验数据表格
            const testData = parseHtmlTable(`<table>${testDataMatch[1]}</table>`);

            // 创建完整的测试记录
            const testRecord = {
                试验编号: testInfo['試驗編號'] || '',
                试验名称: testInfo['試驗中文名稱'] || '',
                试验规范: testInfo['試驗規範'] || '',
                试验公司: testInfo['試驗公司'] || '',
                试验完成日期: testInfo['試驗完成日期'] || '',
                备注: testInfo['備註'] || '',
                数据: testData || []
            };

            results.push(testRecord);
        }

        return results;
    } catch (error) {
        logger(`解析测试数据失败: ${error.message}`);
        logger(error.stack);
        return [];
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

// 提取深度信息
function extractDepthInfo(data) {
    if (!data) return null;

    // 尝试从数组中提取深度信息
    if (Array.isArray(data) && data.length > 0) {
        // 检查topDepth和bottomDepth字段
        if (data[0].topDepth !== undefined && data[0].bottomDepth !== undefined) {
            return `深度${data[0].topDepth}-${data[0].bottomDepth}公尺`;
        }

        // 检查上限深度和下限深度字段
        if (data[0]['上限深度(公尺)'] && data[0]['下限深度(公尺)']) {
            return `深度${data[0]['上限深度(公尺)']}-${data[0]['下限深度(公尺)']}公尺`;
        }
    }

    // 从键值对对象中提取
    if (typeof data === 'object' && !Array.isArray(data)) {
        if (data['上限深度(公尺)'] && data['下限深度(公尺)']) {
            return `深度${data['上限深度(公尺)']}-${data['下限深度(公尺)']}公尺`;
        }
    }

    return null;
}

// 安全的文件名生成函数
function getSafeFileName(name) {
    if (!name) return 'unknown';

    // 移除不安全的文件名字符
    return name.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
        .replace(/\s+/g, '_')
        .replace(/__+/g, '_')
        .trim();
}

// 主函数
async function main() {
    logger('启动台湾地质钻孔数据批量下载工具');

    try {
        // 安装Playwright浏览器
        try {
            logger('检查浏览器是否已安装...');
            execSync('npx playwright install chromium', { stdio: 'inherit' });
            logger('浏览器安装/检查完成');
        } catch (error) {
            logger(`浏览器安装失败: ${error.message}`);
            logger('请尝试手动运行: npx playwright install chromium');
        }

        // 启动浏览器
        const browser = await chromium.launch(CONFIG.browser);
        logger('浏览器启动成功');

        const context = await browser.newContext({
            userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36 Edg/135.0.0.0'
        });

        // 创建新页面
        const page = await context.newPage();

        try {
            // 导航到目标网站
            logger('正在访问地质资料网站...');
            await page.goto('https://geotech.gsmma.gov.tw/Imoeagis/Home/Map', { timeout: 60000 });

            // 等待页面加载
            await page.waitForLoadState('networkidle');
            logger('页面加载完成');

            // 获取项目列表
            logger('发送钻探项目搜索请求...');
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

            logger(`获取到 ${projectsResponse.length} 个钻探项目`);

            // 保存项目数据
            fs.writeFileSync(
                path.join(CONFIG.outputDir, 'projects.json'),
                JSON.stringify(projectsResponse, null, 2)
            );

            // 创建项目信息数组
            const projectInfoArray = [];

            // 处理每个项目
            for (const project of projectsResponse) {
                // 跳过无效项目
                if (!project.keyid || project.drillHoleCount <= 0) {
                    logger(`跳过无效项目: ${project.keyid || '未知ID'}`);
                    continue;
                }

                try {
                    await processProject(page, project, projectInfoArray);
                } catch (error) {
                    logger(`处理项目时出错: ${error.message}`);
                    logger(error.stack);
                }

                // 项目间添加延迟
                await delay(CONFIG.requestDelay);
            }

            // 保存项目信息概览到Excel
            await createExcel(projectInfoArray, path.join(CONFIG.outputDir, '项目概览.xlsx'));

            logger('所有项目处理完成');

        } catch (error) {
            logger(`页面处理错误: ${error.message}`);
            logger(error.stack);
        } finally {
            // 关闭浏览器
            await browser.close();
            logger('浏览器已关闭');
        }
    } catch (error) {
        logger(`浏览器启动错误: ${error.message}`);
        logger(error.stack);
    }

    logger('爬虫执行完成');
}

// 处理单个项目
async function processProject(page, project, projectInfoArray) {
    const projectId = project.keyid;

    // 解码项目名称
    const projectName = decodeUnicode(project.projName) || '未命名项目';
    const projectHoleCount = project.drillHoleCount || 0;
    const formattedProjectName = `${projectName}(${projectHoleCount}钻孔)`;

    // 创建安全的目录名
    const safeDirName = getSafeFileName(formattedProjectName);

    logger(`处理项目 ID: ${projectId}, 名称: ${formattedProjectName}`);

    // 添加到项目信息数组
    projectInfoArray.push({
        项目ID: projectId,
        项目名称: projectName,
        项目编号: project.projNo,
        钻孔数量: projectHoleCount,
        执行单位: decodeUnicode(project.orgname) || '',
        状态: decodeUnicode(project.sStatus) || ''
    });

    // 创建项目目录（使用格式化后的项目名称）
    const projectDir = path.join(CONFIG.outputDir, safeDirName);
    ensureDirectoryExists(projectDir);

    // 保存项目基本信息
    fs.writeFileSync(
        path.join(projectDir, 'project_info.json'),
        JSON.stringify(project, null, 2)
    );

    // 获取钻孔坐标数据
    logger(`获取项目 ${projectId} 的钻孔坐标数据...`);
    let retry = 0;
    let holeResponse = null;

    while (retry < CONFIG.maxRetries) {
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
            logger(`获取钻孔数据失败，尝试重试 (${retry}/${CONFIG.maxRetries}): ${error.message}`);
            await delay(CONFIG.requestDelay);
        }
    }

    if (!holeResponse || !Array.isArray(holeResponse)) {
        logger(`无法获取项目 ${projectId} 的钻孔数据，跳过`);
        return;
    }

    logger(`获取到 ${holeResponse.length} 个钻孔数据`);

    // 保存钻孔坐标数据
    fs.writeFileSync(
        path.join(projectDir, 'drill_holes.json'),
        JSON.stringify(holeResponse, null, 2)
    );

    // 创建详细数据目录
    const detailsDir = path.join(projectDir, 'details');
    ensureDirectoryExists(detailsDir);

    // 处理每个钻孔
    let holesToProcess = holeResponse;
    if (CONFIG.maxHolesPerProject > 0 && holeResponse.length > CONFIG.maxHolesPerProject) {
        holesToProcess = holeResponse.slice(0, CONFIG.maxHolesPerProject);
        logger(`项目 ${projectId} 有 ${holeResponse.length} 个钻孔，但只处理前 ${CONFIG.maxHolesPerProject} 个以避免请求过多`);
    }

    for (const hole of holesToProcess) {
        try {
            // 解析钻孔编号
            const holeNo = hole.holePointNo || '未知编号';

            // 创建安全的钻孔目录名
            const holeFolderName = `hole_${getSafeFileName(holeNo)}`;

            // 处理钻孔数据
            await processHole(page, hole, path.join(detailsDir, holeFolderName));

            // 钻孔处理间添加延迟
            await delay(CONFIG.requestDelay);

        } catch (error) {
            logger(`处理钻孔 ${hole.keyid} 时出错: ${error.message}`);
            logger(error.stack);
        }
    }

    logger(`项目 ${projectId} 处理完成`);
}

// 处理单个钻孔 - 优化版
async function processHole(page, hole, outputDir) {
    const holeId = hole.keyid;
    const projectId = hole.projectKeyid;
    const holeNo = hole.holePointNo || '未知编号';

    logger(`处理钻孔 ID: ${holeId}, 编号: ${holeNo}`);

    // 创建钻孔目录
    ensureDirectoryExists(outputDir);

    // 保存钻孔基本信息
    fs.writeFileSync(
        path.join(outputDir, 'hole_info.json'),
        JSON.stringify(hole, null, 2)
    );

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
                console.error('记录日志失败', error);
            }
        }, { projectId, holeNo });

        // 获取基本资料
        if (CONFIG.saveBaseData) {
            await processBaseData(page, projectId, holeId, outputDir);
        }

        // 获取试验资料
        if (CONFIG.saveTestData) {
            await processTestData(page, projectId, holeId, outputDir);
        }

        // 获取柱状图
        if (CONFIG.saveChartImages) {
            await processChartData(page, projectId, holeId, outputDir);
        }

        // 获取岩心照片
        if (CONFIG.saveCoreImages) {
            await processCoreImages(page, holeId, outputDir);
        }

        logger(`钻孔 ${holeId} 处理完成`);
    } catch (error) {
        logger(`处理钻孔 ${holeId} 时出错: ${error.message}`);
        logger(error.stack);
        throw error;
    }
}

// 处理基本资料
async function processBaseData(page, projectId, holeId, outputDir) {
    logger(`获取钻孔 ${holeId} 的基本资料...`);

    // 创建基本资料目录
    const baseDataDir = path.join(outputDir, 'base_data');

    // 获取基本资料
    const baseData = await getHoleDetail(page, 'BaseData', projectId, holeId);
    if (!baseData || !Array.isArray(baseData) || baseData.length === 0) {
        logger(`钻孔 ${holeId} 没有基本资料`);
        return;
    }

    // 确保目录存在
    ensureDirectoryExists(baseDataDir);

    // 解析HTML表格数据
    const parsedData = parseHtmlTable(baseData[0]);
    if (!parsedData) {
        logger(`钻孔 ${holeId} 基本资料解析失败`);
        return;
    }

    // 保存为Excel
    await createExcel(parsedData, path.join(baseDataDir, 'base_data.xlsx'));

    // 保存原始JSON
    fs.writeFileSync(
        path.join(baseDataDir, 'base_data.json'),
        JSON.stringify(baseData, null, 2)
    );

    logger(`钻孔 ${holeId} 基本资料处理完成`);
}

// 处理试验资料
// 处理试验资料 - 修改版（单工作表）
async function processTestData(page, projectId, holeId, outputDir) {
    logger(`获取钻孔 ${holeId} 的试验资料...`);

    // 创建试验资料目录
    const testDataDir = path.join(outputDir, 'test_data');

    // 获取试验资料
    const testData = await getHoleDetail(page, 'Test', projectId, holeId);
    if (!testData || !Array.isArray(testData) || testData.length === 0) {
        logger(`钻孔 ${holeId} 没有试验资料`);
        return;
    }

    // 确保目录存在
    ensureDirectoryExists(testDataDir);

    // 解析测试数据
    const parsedTests = parseTestData(testData[0]);
    if (!parsedTests || parsedTests.length === 0) {
        logger(`钻孔 ${holeId} 试验资料解析失败`);

        // 保存原始数据以便调试
        fs.writeFileSync(
            path.join(testDataDir, 'test_data_raw.json'),
            JSON.stringify(testData, null, 2)
        );
        return;
    }

    // 创建Excel工作簿
    const workbook = new ExcelJS.Workbook();

    // 添加单个工作表
    const worksheet = workbook.addWorksheet('试验数据');

    // 当前行号，用于跟踪插入位置
    let currentRow = 1;

    // 处理每个测试
    for (let i = 0; i < parsedTests.length; i++) {
        const test = parsedTests[i];

        // 如果不是第一个测试，添加一个空行分隔
        if (i > 0) {
            worksheet.addRow([]);
            currentRow++;
        }

        // 添加测试基本信息头部
        const headerRow = worksheet.addRow(['试验编号', '试验名称', '试验公司', '试验完成日期', '备注']);
        currentRow++;

        // 设置基本信息头部样式
        headerRow.eachCell((cell) => {
            cell.font = { bold: true };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFD9D9D9' } // 浅灰色背景
            };
        });

        // 添加测试基本信息数据
        worksheet.addRow([
            test.试验编号 || '',
            test.试验名称 || '',
            test.试验公司 || '',
            test.试验完成日期 || '',
            test.备注 || ''
        ]);
        currentRow++;

        // 添加空行
        worksheet.addRow([]);
        currentRow++;

        // 处理测试数据
        const testDataArray = test.数据;
        if (Array.isArray(testDataArray) && testDataArray.length > 0) {
            // 获取所有列名
            const dataColumns = new Set();
            testDataArray.forEach(row => {
                if (row && typeof row === 'object') {
                    Object.keys(row).forEach(key => dataColumns.add(key));
                }
            });

            // 添加数据表头
            const dataHeaders = Array.from(dataColumns);
            const dataHeaderRow = worksheet.addRow(dataHeaders);
            currentRow++;

            // 设置数据表头样式
            dataHeaderRow.eachCell((cell) => {
                cell.font = { bold: true };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE6F0FF' } // 浅蓝色背景
                };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            // 添加数据行
            testDataArray.forEach(item => {
                if (item && typeof item === 'object') {
                    const rowData = dataHeaders.map(header => {
                        const value = item[header];
                        if (value === undefined || value === null) return '';
                        return value;
                    });
                    worksheet.addRow(rowData);
                    currentRow++;
                }
            });
        } else {
            // 没有数据的情况
            worksheet.addRow(['没有可用的测试数据']);
            currentRow++;
        }
    }

    // 设置列宽
    worksheet.columns.forEach(column => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, cell => {
            if (cell.value) {
                const length = cell.value.toString().length;
                if (length > maxLength) {
                    maxLength = length;
                }
            }
        });
        column.width = Math.min(Math.max(maxLength + 2, 12), 40);
    });

    // 保存Excel文件
    await workbook.xlsx.writeFile(path.join(testDataDir, 'test_data.xlsx'));
    logger(`试验资料已保存至: ${path.join(testDataDir, 'test_data.xlsx')}`);

    // 保存原始JSON
    fs.writeFileSync(
        path.join(testDataDir, 'test_data.json'),
        JSON.stringify(testData, null, 2)
    );

    // 以易读格式保存解析后的数据
    fs.writeFileSync(
        path.join(testDataDir, 'test_data_parsed.json'),
        JSON.stringify(parsedTests, null, 2)
    );

    logger(`钻孔 ${holeId} 试验资料处理完成，共处理了 ${parsedTests.length} 个试验项目`);
}

// 处理柱状图数据
async function processChartData(page, projectId, holeId, outputDir) {
    logger(`获取钻孔 ${holeId} 的柱状图数据...`);

    // 创建柱状图目录
    const chartDir = path.join(outputDir, 'chart_data');

    // 获取柱状图数据
    const chartData = await getHoleDetail(page, 'Chart', projectId, holeId);
    if (!chartData || !Array.isArray(chartData) || chartData.length === 0) {
        logger(`钻孔 ${holeId} 没有柱状图数据`);
        return;
    }

    // 确保目录存在
    ensureDirectoryExists(chartDir);

    // 保存所有图片数据
    let totalImageCount = 0;

    // 处理每个chartData元素
    for (let dataIndex = 0; dataIndex < chartData.length; dataIndex++) {
        const chartDataItem = chartData[dataIndex];

        // 提取图片链接
        const images = extractImagesFromHtml(chartDataItem);
        if (images.length === 0) {
            logger(`钻孔 ${holeId} 的第 ${dataIndex + 1} 组柱状图数据未找到图片`);
            continue;
        }

        // 保存图片
        for (let i = 0; i < images.length; i++) {
            // 生成图片文件名（使用dataIndex避免文件重名）
            const fileName = `柱状图_${dataIndex + 1}_${i + 1}.png`;
            const imagePath = path.join(chartDir, fileName);

            if (saveBase64Image(images[i], imagePath)) {
                totalImageCount++;
            }
        }
    }

    // 保存原始JSON
    fs.writeFileSync(
        path.join(chartDir, 'chart_data.json'),
        JSON.stringify(chartData, null, 2)
    );

    logger(`钻孔 ${holeId} 柱状图数据处理完成，共保存了 ${totalImageCount} 张图片`);
}

// 处理岩心照片
async function processCoreImages(page, holeId, outputDir) {
    logger(`获取钻孔 ${holeId} 的岩心照片...`);

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
            logger(`钻孔 ${holeId} 没有岩心照片数据`);
            return;
        }

        // 确保目录存在
        ensureDirectoryExists(coreDir);

        // 保存岩心照片
        let successCount = 0;
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
            if (image.imagePath) {
                const success = saveBase64Image(image.imagePath, imagePath);
                if (success) successCount++;
            }
        }

        // 保存原始JSON
        fs.writeFileSync(
            path.join(coreDir, 'core_images.json'),
            JSON.stringify(coreImagesResponse, null, 2)
        );

        logger(`钻孔 ${holeId} 岩心照片处理完成，成功保存了 ${successCount} 张照片`);

    } catch (error) {
        logger(`获取岩心照片失败: ${error.message}`);
        logger(error.stack);
    }
}

// 获取钻孔详细信息（通用函数）
async function getHoleDetail(page, mode, projectId, holeId) {
    logger(`获取钻孔 ${holeId} 的 ${mode} 数据...`);

    // 重试机制
    let retry = 0;
    while (retry < CONFIG.maxRetries) {
        try {
            // 添加短暂延迟避免请求过快
            await delay(CONFIG.requestDelay / 2);

            const response = await page.evaluate(async (params) => {
                // 生成时间戳（模拟浏览器行为）
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
            retry++;
            logger(`获取钻孔 ${holeId} 的 ${mode} 数据失败，尝试重试 (${retry}/${CONFIG.maxRetries}): ${error.message}`);

            // 增加重试延迟
            await delay(CONFIG.requestDelay * retry);
        }
    }

    logger(`获取钻孔 ${holeId} 的 ${mode} 数据失败，已达最大重试次数`);
    return null;
}

// 检查NPM依赖并安装
async function checkAndInstallDependencies() {
    try {
        logger('检查并安装必要的NPM依赖...');

        // 检查package.json是否存在
        if (!fs.existsSync('package.json')) {
            logger('创建package.json文件...');
            execSync('npm init -y', { stdio: 'inherit' });
        }

        // 检查所需依赖
        const dependencies = ['exceljs', 'playwright'];
        const missingDeps = [];

        for (const dep of dependencies) {
            try {
                require.resolve(dep);
                logger(`${dep} 已安装`);
            } catch (e) {
                logger(`${dep} 未安装，添加到安装列表`);
                missingDeps.push(dep);
            }
        }

        // 安装缺失的依赖
        if (missingDeps.length > 0) {
            logger(`安装缺失的依赖: ${missingDeps.join(', ')}...`);
            execSync(`npm install ${missingDeps.join(' ')}`, { stdio: 'inherit' });
        }

        logger('依赖检查完成');
        return true;
    } catch (error) {
        logger(`依赖检查失败: ${error.message}`);
        logger('请手动安装依赖: npm install exceljs playwright');
        return false;
    }
}

// 入口点
async function start() {
    try {
        // 检查依赖
        await checkAndInstallDependencies();

        // 执行主函数
        await main();
    } catch (error) {
        logger(`爬虫执行失败: ${error.message}`);
        logger(error.stack);
    }
}

// 启动程序
start();
