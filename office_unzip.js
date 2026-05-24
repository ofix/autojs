const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { glob } = require('glob');
const xml2js = require('xml2js');
const { Builder } = require('xml2js');

// 支持的 Office 文档扩展名（都是 ZIP 格式）
const SUPPORTED_EXTENSIONS = {
    // Word 文档
    word: ['.docx', '.docm', '.dotx', '.dotm'],
    // PowerPoint 文档
    powerpoint: ['.pptx', '.pptm', '.potx', '.potm', '.ppsx', '.ppsm'],
    // Excel 文档
    excel: ['.xlsx', '.xlsm', '.xltx', '.xltm', '.xlsb'],
    // 其他支持的格式
    other: ['.vsdx', '.pub', '.vdw'] // Visio, Publisher 等
};

// 所有支持的扩展名
const ALL_SUPPORTED_EXTENSIONS = [
    ...SUPPORTED_EXTENSIONS.word,
    ...SUPPORTED_EXTENSIONS.powerpoint,
    ...SUPPORTED_EXTENSIONS.excel,
    ...SUPPORTED_EXTENSIONS.other
];

// XML 格式化选项
const XML_FORMAT_OPTIONS = {
    indent: '  ',
    newline: '\n',
    pretty: true,
    cdata: true
};

/**
 * 获取文档类型
 * @param {string} filePath 文件路径
 * @returns {string} 文档类型 (word/powerpoint/excel/other/unknown)
 */
function getDocumentType(filePath) {
    const ext = path.extname(filePath).toLowerCase();

    if (SUPPORTED_EXTENSIONS.word.includes(ext)) return 'word';
    if (SUPPORTED_EXTENSIONS.powerpoint.includes(ext)) return 'powerpoint';
    if (SUPPORTED_EXTENSIONS.excel.includes(ext)) return 'excel';
    if (SUPPORTED_EXTENSIONS.other.includes(ext)) return 'other';
    return 'unknown';
}

/**
 * 检查文件是否为支持的 Office 文档
 * @param {string} filePath 文件路径
 * @returns {boolean} 是否为支持的文档
 */
function isOfficeDocument(filePath) {
    const ext = path.extname(filePath).toLowerCase();
    return ALL_SUPPORTED_EXTENSIONS.includes(ext);
}

/**
 * 获取文档类型的图标/标识
 */
function getDocumentTypeIcon(docType) {
    const icons = {
        word: '📝',
        powerpoint: '📊',
        excel: '📈',
        other: '📄',
        unknown: '❓'
    };
    return icons[docType] || '📄';
}

/**
 * 获取文档类型的中文名称
 */
function getDocumentTypeName(docType) {
    const names = {
        word: 'Word',
        powerpoint: 'PowerPoint',
        excel: 'Excel',
        other: 'Office',
        unknown: '未知'
    };
    return names[docType] || 'Office文档';
}

/**
 * 获取目录统计信息
 */
function getDirectoryStats(dirPath) {
    let files = 0;
    let dirs = 0;

    try {
        const items = fs.readdirSync(dirPath);
        for (const item of items) {
            const itemPath = path.join(dirPath, item);
            if (fs.statSync(itemPath).isDirectory()) {
                dirs++;
                const subStats = getDirectoryStats(itemPath);
                files += subStats.files;
                dirs += subStats.dirs;
            } else {
                files++;
            }
        }
    } catch (error) {
        console.error(`统计目录失败: ${dirPath}`, error.message);
    }

    return { files, dirs };
}

/**
 * 格式化单个 XML 文件
 */
async function formatXmlFile(filePath) {
    try {
        const content = fs.readFileSync(filePath, 'utf-8');

        // 解析 XML
        const parser = new xml2js.Parser({
            explicitArray: false,
            mergeAttrs: false,
            normalize: false,
            trim: false
        });

        const parsed = await parser.parseStringPromise(content);

        // 构建格式化后的 XML
        const builder = new Builder({
            xmldec: { version: '1.0', encoding: 'UTF-8', standalone: null },
            renderOpts: { pretty: true, indent: '  ', newline: '\n' },
            headless: false,
            allowSurrogateChars: true,
            cdata: true
        });

        const formattedXml = builder.buildObject(parsed);
        fs.writeFileSync(filePath, formattedXml, 'utf-8');

        return { success: true, message: `✓ 已格式化: ${path.basename(filePath)}` };
    } catch (error) {
        // 如果解析失败，尝试简单格式化
        try {
            const content = fs.readFileSync(filePath, 'utf-8');
            const simpleFormatted = simpleXmlFormat(content);
            fs.writeFileSync(filePath, simpleFormatted, 'utf-8');
            return { success: true, message: `✓ 已格式化(简单): ${path.basename(filePath)}` };
        } catch (simpleError) {
            return { success: false, message: `✗ 格式化失败: ${path.basename(filePath)}` };
        }
    }
}

/**
 * 简单的 XML 格式化
 */
function simpleXmlFormat(xmlStr) {
    let formatted = '';
    let indent = 0;
    let inTag = false;
    let inCData = false;

    for (let i = 0; i < xmlStr.length; i++) {
        const char = xmlStr[i];

        if (char === '<' && xmlStr.substr(i, 9) === '<![CDATA[') {
            inCData = true;
            formatted += char;
            continue;
        }
        if (inCData && char === '>' && xmlStr.substr(i - 2, 3) === ']]>') {
            inCData = false;
            formatted += char;
            continue;
        }

        if (inCData) {
            formatted += char;
            continue;
        }

        if (char === '<') {
            if (xmlStr[i + 1] === '/') {
                if (!inTag) {
                    indent = Math.max(0, indent - 1);
                    if (formatted.length > 0 && formatted[formatted.length - 1] !== '\n') {
                        formatted += '\n';
                    }
                    formatted += '  '.repeat(indent);
                }
            } else if (xmlStr[i + 1] !== '?' && xmlStr[i + 1] !== '!') {
                if (!inTag && formatted.length > 0 && formatted[formatted.length - 1] !== '\n') {
                    formatted += '\n';
                    formatted += '  '.repeat(indent);
                }
                if (xmlStr[i + 1] !== '/' && xmlStr[i + 1] !== '?') {
                    indent++;
                }
            }
            inTag = true;
            formatted += char;
        }
        else if (char === '>') {
            inTag = false;
            formatted += char;
            if (xmlStr[i - 1] === '/') {
                indent = Math.max(0, indent - 1);
            }
            if (i + 1 < xmlStr.length && xmlStr[i + 1] !== '<') {
                formatted += '\n';
            }
        }
        else {
            formatted += char;
        }
    }

    return formatted.replace(/\n\s*\n/g, '\n');
}

/**
 * 格式化目录中的所有 XML 文件
 */
async function formatXmlFilesInDirectory(dirPath, options = {}) {
    const results = { formatted: 0, failed: 0, total: 0, details: [] };
    const xmlFiles = [];

    function findXmlFiles(dir) {
        const items = fs.readdirSync(dir);
        for (const item of items) {
            const fullPath = path.join(dir, item);
            const stat = fs.statSync(fullPath);
            if (stat.isDirectory()) {
                findXmlFiles(fullPath);
            } else if (item.endsWith('.xml') || item.endsWith('.rels')) {
                xmlFiles.push(fullPath);
            }
        }
    }

    findXmlFiles(dirPath);
    results.total = xmlFiles.length;

    if (xmlFiles.length === 0) {
        if (!options.quiet) {
            console.log(`📭 目录中没有 XML 文件: ${dirPath}`);
        }
        return results;
    }

    if (!options.quiet) {
        console.log(`\n🎨 开始格式化 XML 文件 (共 ${xmlFiles.length} 个)...`);
    }

    const concurrency = options.concurrency || 10;
    for (let i = 0; i < xmlFiles.length; i += concurrency) {
        const chunk = xmlFiles.slice(i, i + concurrency);
        const promises = chunk.map(async (xmlFile) => {
            const result = await formatXmlFile(xmlFile);
            if (result.success) {
                results.formatted++;
            } else {
                results.failed++;
            }
            if (!options.quiet) {
                console.log(result.message);
            }
            results.details.push(result);
        });
        await Promise.all(promises);
    }

    if (!options.quiet) {
        console.log(`\n✅ XML 格式化完成: ${results.formatted} 成功, ${results.failed} 失败`);
    }

    return results;
}

/**
 * 解压单个 Office 文档
 */
async function extractOfficeDocument(filePath, outputDir = null, options = {}) {
    const formatXml = options.formatXml !== false;
    const quiet = options.quiet || false;

    try {
        if (!fs.existsSync(filePath)) {
            return {
                success: false,
                message: `文件不存在: ${filePath}`,
                outputPath: null,
                filePath: filePath,
                docType: null,
                xmlFormatResult: null
            };
        }

        const docType = getDocumentType(filePath);
        const docTypeName = getDocumentTypeName(docType);
        const docIcon = getDocumentTypeIcon(docType);

        if (docType === 'unknown') {
            return {
                success: false,
                message: `⚠ 不支持的文档类型: ${path.basename(filePath)}`,
                outputPath: null,
                filePath: filePath,
                docType: docType,
                xmlFormatResult: null
            };
        }

        if (!outputDir) {
            const baseName = path.basename(filePath, path.extname(filePath));
            const dirName = path.dirname(filePath);
            outputDir = path.join(dirName, `${baseName}_extracted`);
        }

        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        // 读取并解压文档
        const fileBuffer = fs.readFileSync(filePath);
        const zip = await JSZip.loadAsync(fileBuffer);

        let fileCount = 0;
        const promises = [];

        zip.forEach((relativePath, zipEntry) => {
            if (!zipEntry.dir) {
                const fullPath = path.join(outputDir, relativePath);
                const fullDir = path.dirname(fullPath);

                if (!fs.existsSync(fullDir)) {
                    fs.mkdirSync(fullDir, { recursive: true });
                }

                const promise = zipEntry.async('nodebuffer').then(content => {
                    fs.writeFileSync(fullPath, content);
                    fileCount++;
                });
                promises.push(promise);
            }
        });

        await Promise.all(promises);

        let xmlFormatResult = null;

        // 格式化 XML 文件
        if (formatXml) {
            if (!quiet) {
                console.log(`\n📝 格式化 ${path.basename(filePath)} 中的 XML 文件...`);
            }
            xmlFormatResult = await formatXmlFilesInDirectory(outputDir, { quiet });
        }

        const stats = getDirectoryStats(outputDir);

        let message = `${docIcon} 解压成功: ${path.basename(filePath)} (${docTypeName}) → ${stats.files} 个文件, ${stats.dirs} 个文件夹`;
        if (xmlFormatResult) {
            message += ` | XML: ${xmlFormatResult.formatted}/${xmlFormatResult.total}`;
        }

        return {
            success: true,
            message: message,
            outputPath: outputDir,
            filePath: filePath,
            docType: docType,
            xmlFormatResult: xmlFormatResult
        };
    } catch (error) {
        return {
            success: false,
            message: `✗ 解压失败: ${path.basename(filePath)} - ${error.message}`,
            outputPath: null,
            filePath: filePath,
            docType: getDocumentType(filePath),
            xmlFormatResult: null
        };
    }
}

/**
 * 处理单个文件
 */
async function processFile(filePath, results, options = {}) {
    if (isOfficeDocument(filePath)) {
        if (!options.quiet) {
            const docType = getDocumentType(filePath);
            const icon = getDocumentTypeIcon(docType);
            console.log(`${icon} 正在处理: ${path.basename(filePath)} (${getDocumentTypeName(docType)})`);
        }
        const result = await extractOfficeDocument(filePath, null, options);
        results.push(result);
        if (!options.quiet) {
            console.log(result.message);
        }
    } else {
        const msg = `⚠ 跳过非 Office 文档: ${path.basename(filePath)} (支持: ${ALL_SUPPORTED_EXTENSIONS.join(', ')})`;
        if (!options.quiet) {
            console.log(msg);
        }
        results.push({
            success: false,
            message: msg,
            outputPath: null,
            filePath: filePath,
            docType: null,
            xmlFormatResult: null
        });
    }
}

/**
 * 处理目录
 */
async function processDirectory(dirPath, results, options = {}) {
    if (!fs.existsSync(dirPath)) {
        const msg = `❌ 目录不存在: ${dirPath}`;
        console.error(msg);
        results.push({
            success: false,
            message: msg,
            outputPath: null,
            filePath: dirPath,
            docType: null,
            xmlFormatResult: null
        });
        return;
    }

    const stats = fs.statSync(dirPath);
    if (!stats.isDirectory()) {
        await processFile(dirPath, results, options);
        return;
    }

    if (!options.quiet) {
        console.log(`\n📁 扫描目录: ${dirPath}`);
    }

    const files = fs.readdirSync(dirPath);
    let officeCount = 0;

    for (const file of files) {
        const fullPath = path.join(dirPath, file);
        const fileStats = fs.statSync(fullPath);

        if (fileStats.isDirectory() && options.recursive !== false) {
            await processDirectory(fullPath, results, options);
        } else if (fileStats.isFile() && isOfficeDocument(fullPath)) {
            officeCount++;
            await processFile(fullPath, results, options);
        }
    }

    if (officeCount === 0 && !options.quiet) {
        console.log(`📭 目录中没有找到 Office 文档`);
    }
}

/**
 * 处理多个路径
 */
async function processMultiplePaths(inputPaths, options = {}) {
    const defaultOptions = {
        recursive: true,
        parallel: true,
        maxParallel: 5,
        quiet: false,
        formatXml: true,
        xmlConcurrency: 10
    };

    const opts = { ...defaultOptions, ...options };
    const results = [];

    // 收集所有需要处理的文件
    const allFiles = [];

    for (const inputPath of inputPaths) {
        if (!fs.existsSync(inputPath)) {
            console.error(`❌ 路径不存在: ${inputPath}`);
            results.push({
                success: false,
                message: `路径不存在: ${inputPath}`,
                outputPath: null,
                filePath: inputPath,
                docType: null,
                xmlFormatResult: null
            });
            continue;
        }

        const stats = fs.statSync(inputPath);

        if (stats.isFile()) {
            if (isOfficeDocument(inputPath)) {
                allFiles.push(inputPath);
            } else if (!opts.quiet) {
                console.log(`⚠ 跳过非 Office 文档: ${path.basename(inputPath)}`);
            }
        } else if (stats.isDirectory()) {
            const findOfficeFiles = (dir) => {
                const files = fs.readdirSync(dir);
                for (const file of files) {
                    const fullPath = path.join(dir, file);
                    const fileStats = fs.statSync(fullPath);
                    if (fileStats.isDirectory() && opts.recursive) {
                        findOfficeFiles(fullPath);
                    } else if (fileStats.isFile() && isOfficeDocument(fullPath)) {
                        allFiles.push(fullPath);
                    }
                }
            };
            findOfficeFiles(inputPath);
        }
    }

    if (allFiles.length === 0) {
        console.log('📭 没有找到任何 Office 文档');
        return { total: 0, success: 0, fail: 0, results: [], xmlStats: null, typeStats: {} };
    }

    // 统计文档类型
    const typeStats = {
        word: 0,
        powerpoint: 0,
        excel: 0,
        other: 0
    };

    allFiles.forEach(file => {
        const docType = getDocumentType(file);
        if (typeStats[docType] !== undefined) {
            typeStats[docType]++;
        }
    });

    console.log(`\n📊 共找到 ${allFiles.length} 个 Office 文档待处理:`);
    console.log(`   📝 Word: ${typeStats.word} 个`);
    console.log(`   📊 PowerPoint: ${typeStats.powerpoint} 个`);
    console.log(`   📈 Excel: ${typeStats.excel} 个`);
    if (typeStats.other > 0) {
        console.log(`   📄 其他: ${typeStats.other} 个`);
    }
    console.log('');

    // 处理文件
    if (opts.parallel) {
        const chunks = [];
        for (let i = 0; i < allFiles.length; i += opts.maxParallel) {
            chunks.push(allFiles.slice(i, i + opts.maxParallel));
        }

        for (const chunk of chunks) {
            const chunkPromises = chunk.map(file => processFile(file, results, opts));
            await Promise.all(chunkPromises);
        }
    } else {
        for (const file of allFiles) {
            await processFile(file, results, opts);
        }
    }

    // 统计结果
    const successCount = results.filter(r => r.success).length;
    const failCount = results.filter(r => !r.success).length;

    // 统计 XML 格式化结果
    let totalXmlFiles = 0;
    let formattedXmlFiles = 0;
    let failedXmlFiles = 0;

    results.forEach(r => {
        if (r.xmlFormatResult) {
            totalXmlFiles += r.xmlFormatResult.total;
            formattedXmlFiles += r.xmlFormatResult.formatted;
            failedXmlFiles += r.xmlFormatResult.failed;
        }
    });

    return {
        total: allFiles.length,
        success: successCount,
        fail: failCount,
        results: results,
        typeStats: typeStats,
        xmlStats: {
            total: totalXmlFiles,
            formatted: formattedXmlFiles,
            failed: failedXmlFiles
        }
    };
}

/**
 * 批量解压主函数
 */
async function batchExtractOfficeDocuments(inputPaths, options = {}) {
    const paths = Array.isArray(inputPaths) ? inputPaths : [inputPaths];

    console.log('\n' + '='.repeat(80));
    console.log('🚀 批量解压 Office 文档工具 (支持 Word/PPT/Excel)');
    console.log('='.repeat(80));
    console.log(`📂 输入路径 (${paths.length} 个):`);
    paths.forEach((p, i) => {
        console.log(`   ${i + 1}. ${p}`);
    });
    console.log(`🔄 递归子目录: ${options.recursive !== false ? '是' : '否'}`);
    console.log(`⚡ 并行处理: ${options.parallel !== false ? '是' : '否'}`);
    console.log(`🎨 XML格式化: ${options.formatXml !== false ? '是' : '否'}`);
    console.log(`📁 支持格式: ${ALL_SUPPORTED_EXTENSIONS.join(', ')}`);
    console.log('='.repeat(80) + '\n');

    const startTime = Date.now();

    try {
        const stats = await processMultiplePaths(paths, options);

        const endTime = Date.now();
        const duration = ((endTime - startTime) / 1000).toFixed(2);

        console.log('\n' + '='.repeat(80));
        console.log('📊 执行完成统计');
        console.log('='.repeat(80));
        console.log(`📄 文档总数: ${stats.total}`);
        console.log(`✅ 解压成功: ${stats.success} 个`);
        console.log(`❌ 解压失败: ${stats.fail} 个`);

        if (stats.typeStats) {
            console.log(`\n📋 文档类型分布:`);
            console.log(`   📝 Word: ${stats.typeStats.word} 个`);
            console.log(`   📊 PowerPoint: ${stats.typeStats.powerpoint} 个`);
            console.log(`   📈 Excel: ${stats.typeStats.excel} 个`);
            if (stats.typeStats.other > 0) {
                console.log(`   📄 其他: ${stats.typeStats.other} 个`);
            }
        }

        if (stats.xmlStats && stats.xmlStats.total > 0) {
            console.log(`\n🎨 XML 格式化统计:`);
            console.log(`   📄 XML 文件总数: ${stats.xmlStats.total}`);
            console.log(`   ✅ 格式化成功: ${stats.xmlStats.formatted}`);
            console.log(`   ❌ 格式化失败: ${stats.xmlStats.failed}`);
        }

        console.log(`⏱️  总耗时: ${duration} 秒`);

        if (stats.fail > 0) {
            console.log('\n失败列表:');
            stats.results.filter(r => !r.success).forEach(r => {
                console.log(`  ${r.message}`);
            });
        }

        console.log('\n' + '='.repeat(80) + '\n');

        return stats;
    } catch (error) {
        console.error('❌ 执行过程中发生错误:', error);
        throw error;
    }
}

/**
 * 命令行接口
 */
async function main() {
    const args = process.argv.slice(2);

    if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
        console.log(`
使用方法:
  node extract-office.js <路径1> [路径2] [路径3] ... [选项]

参数:
  <路径>        一个或多个 Office 文件路径或包含 Office 文件的目录路径

选项:
  --no-format    不解压后格式化 XML（默认会格式化）
  --no-recursive 不递归子目录（默认会递归）
  --serial       串行处理（默认并行）
  --quiet        静默模式（减少输出）
  --format-only <目录>  只格式化已解压的目录中的 XML

支持的格式:
  📝 Word:       .docx, .docm, .dotx, .dotm
  📊 PowerPoint: .pptx, .pptm, .potx, .potm, .ppsx, .ppsm
  📈 Excel:      .xlsx, .xlsm, .xltx, .xltm, .xlsb
  📄 其他:       .vsdx (Visio), .pub (Publisher)

示例:
  # 解压单个文件
  node extract-office.js document.docx

  # 解压多个不同类型的文件
  node extract-office.js word.docx presentation.pptx data.xlsx

  # 解压目录下所有 Office 文件
  node extract-office.js ./documents

  # 解压多个目录
  node extract-office.js ./docs ./slides ./sheets

  # 混合使用文件和目录
  node extract-office.js ./documents special.pptx ./spreadsheets

  # 只格式化已解压的文件夹
  node extract-office.js --format-only ./document_extracted

  # 使用通配符
  node extract-office.js "./docs/*.docx" "./slides/*.pptx"
        `);
        return;
    }

    // 检查是否为只格式化模式
    if (args[0] === '--format-only') {
        if (args.length < 2) {
            console.error('❌ 请指定要格式化的目录');
            return;
        }
        const dirPath = args[1];
        await formatXmlFilesInDirectory(dirPath, { quiet: args.includes('--quiet') });
        return;
    }

    // 解析选项
    const options = {
        recursive: !args.includes('--no-recursive'),
        parallel: !args.includes('--serial'),
        quiet: args.includes('--quiet'),
        formatXml: !args.includes('--no-format')
    };

    // 过滤掉选项参数，保留路径
    let paths = args.filter(arg => !arg.startsWith('--'));

    // 通配符支持
    const hasGlobPattern = paths.some(p => p.includes('*') || p.includes('?'));

    if (hasGlobPattern) {
        try {
            const { glob } = require('glob');
            let allFiles = [];
            for (const pattern of paths) {
                const files = await glob(pattern, { absolute: true });
                allFiles.push(...files);
            }
            paths = [...new Set(allFiles)];

            if (paths.length === 0) {
                console.log('📭 没有找到匹配的文件');
                return;
            }

            console.log(`📊 通过通配符匹配到 ${paths.length} 个文件`);
        } catch (error) {
            console.error('❌ 需要安装 glob 模块来支持通配符: npm install glob');
            return;
        }
    }

    await batchExtractOfficeDocuments(paths, options);
}

// 导出函数
module.exports = {
    extractOfficeDocument,
    batchExtractOfficeDocuments,
    processMultiplePaths,
    formatXmlFilesInDirectory,
    isOfficeDocument,
    getDocumentType,
    getDocumentTypeName,
    SUPPORTED_EXTENSIONS
};

// 如果直接运行脚本
if (require.main === module) {
    main();
}