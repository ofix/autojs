const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { exec } = require('child_process');
const util = require('util');
const execPromise = util.promisify(exec);

// 支持的 Word 文件扩展名
const SUPPORTED_EXTENSIONS = ['.docx', '.docm', '.dotx', '.dotm'];

/**
 * 检查文件是否为 Word 文档
 * @param {string} filePath 文件路径
 * @returns {boolean} 是否为支持的 Word 文档
 */
function isWordDocument(filePath) {
    const ext = path.extname(filePath).toLowerCase();
    return SUPPORTED_EXTENSIONS.includes(ext);
}

/**
 * 获取目录统计信息
 * @param {string} dirPath 目录路径
 * @returns {{files: number, dirs: number}}
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
 * 解压单个 Word 文档
 * @param {string} wordFilePath Word 文件路径
 * @param {string} outputDir 输出目录（如果未指定，则解压到同目录下的同名文件夹）
 * @returns {Promise<{success: boolean, message: string, outputPath: string, filePath: string}>}
 */
async function extractWordDoc(wordFilePath, outputDir = null) {
    try {
        // 检查文件是否存在
        if (!fs.existsSync(wordFilePath)) {
            return {
                success: false,
                message: `文件不存在: ${wordFilePath}`,
                outputPath: null,
                filePath: wordFilePath
            };
        }

        // 确定输出目录
        if (!outputDir) {
            const baseName = path.basename(wordFilePath, path.extname(wordFilePath));
            const dirName = path.dirname(wordFilePath);
            outputDir = path.join(dirName, `${baseName}_extracted`);
        }

        // 创建输出目录（如果不存在）
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }

        // 读取 Word 文件（本质是 ZIP）
        const fileBuffer = fs.readFileSync(wordFilePath);
        
        // 使用 JSZip 解压
        const zip = await JSZip.loadAsync(fileBuffer);
        
        let fileCount = 0;
        
        // 遍历 ZIP 中的所有文件并解压
        const promises = [];
        zip.forEach((relativePath, zipEntry) => {
            if (!zipEntry.dir) {
                const fullPath = path.join(outputDir, relativePath);
                const fullDir = path.dirname(fullPath);
                
                // 确保父目录存在
                if (!fs.existsSync(fullDir)) {
                    fs.mkdirSync(fullDir, { recursive: true });
                }
                
                // 异步解压文件
                const promise = zipEntry.async('nodebuffer').then(content => {
                    fs.writeFileSync(fullPath, content);
                    fileCount++;
                });
                promises.push(promise);
            }
        });
        
        await Promise.all(promises);
        
        const stats = getDirectoryStats(outputDir);
        
        return {
            success: true,
            message: `✓ 解压成功: ${path.basename(wordFilePath)} → ${stats.files} 个文件, ${stats.dirs} 个文件夹`,
            outputPath: outputDir,
            filePath: wordFilePath
        };
    } catch (error) {
        return {
            success: false,
            message: `✗ 解压失败: ${path.basename(wordFilePath)} - ${error.message}`,
            outputPath: null,
            filePath: wordFilePath
        };
    }
}

/**
 * 处理单个文件
 * @param {string} filePath 文件路径
 * @param {Array} results 结果数组
 * @param {Object} options 选项
 */
async function processFile(filePath, results, options = {}) {
    if (isWordDocument(filePath)) {
        if (!options.quiet) {
            console.log(`📄 正在处理 Word 文件: ${path.basename(filePath)}`);
        }
        const result = await extractWordDoc(filePath);
        results.push(result);
        if (!options.quiet) {
            console.log(result.message);
        }
    } else {
        const msg = `⚠ 跳过非 Word 文件: ${path.basename(filePath)}`;
        if (!options.quiet) {
            console.log(msg);
        }
        results.push({
            success: false,
            message: msg,
            outputPath: null,
            filePath: filePath
        });
    }
}

/**
 * 处理目录：遍历所有 Word 文件
 * @param {string} dirPath 目录路径
 * @param {Array} results 结果数组
 * @param {Object} options 选项
 */
async function processDirectory(dirPath, results, options = {}) {
    if (!fs.existsSync(dirPath)) {
        const msg = `❌ 目录不存在: ${dirPath}`;
        console.error(msg);
        results.push({
            success: false,
            message: msg,
            outputPath: null,
            filePath: dirPath
        });
        return;
    }

    const stats = fs.statSync(dirPath);
    if (!stats.isDirectory()) {
        // 如果是文件，直接处理
        await processFile(dirPath, results, options);
        return;
    }

    if (!options.quiet) {
        console.log(`\n📁 扫描目录: ${dirPath}`);
    }
    
    const files = fs.readdirSync(dirPath);
    let wordCount = 0;
    
    for (const file of files) {
        const fullPath = path.join(dirPath, file);
        const fileStats = fs.statSync(fullPath);
        
        if (fileStats.isDirectory() && options.recursive !== false) {
            // 递归处理子目录
            await processDirectory(fullPath, results, options);
        } else if (fileStats.isFile() && isWordDocument(fullPath)) {
            wordCount++;
            await processFile(fullPath, results, options);
        }
    }
    
    if (wordCount === 0 && !options.quiet) {
        console.log(`📭 目录中没有找到 Word 文件: ${dirPath}`);
    }
}

/**
 * 处理多个路径（文件或目录）
 * @param {Array<string>} inputPaths 输入路径数组
 * @param {Object} options 选项
 * @returns {Promise<Object>} 处理结果统计
 */
async function processMultiplePaths(inputPaths, options = {}) {
    const defaultOptions = {
        recursive: true,      // 是否递归子目录
        parallel: true,       // 是否并行处理多个路径
        maxParallel: 5,       // 最大并行数
        quiet: false,         // 是否静默模式
        outputDir: null       // 指定输出目录（仅对单文件有效）
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
                filePath: inputPath
            });
            continue;
        }
        
        const stats = fs.statSync(inputPath);
        
        if (stats.isFile()) {
            if (isWordDocument(inputPath)) {
                allFiles.push(inputPath);
            } else if (!opts.quiet) {
                console.log(`⚠ 跳过非 Word 文件: ${path.basename(inputPath)}`);
            }
        } else if (stats.isDirectory()) {
            // 递归查找所有 Word 文件
            const findWordFiles = (dir) => {
                const files = fs.readdirSync(dir);
                for (const file of files) {
                    const fullPath = path.join(dir, file);
                    const fileStats = fs.statSync(fullPath);
                    if (fileStats.isDirectory() && opts.recursive) {
                        findWordFiles(fullPath);
                    } else if (fileStats.isFile() && isWordDocument(fullPath)) {
                        allFiles.push(fullPath);
                    }
                }
            };
            findWordFiles(inputPath);
        }
    }
    
    if (allFiles.length === 0) {
        console.log('📭 没有找到任何 Word 文件');
        return { total: 0, success: 0, fail: 0, results: [] };
    }
    
    console.log(`\n📊 共找到 ${allFiles.length} 个 Word 文件待处理\n`);
    
    // 处理文件
    if (opts.parallel) {
        // 并行处理，控制并发数
        const chunks = [];
        for (let i = 0; i < allFiles.length; i += opts.maxParallel) {
            chunks.push(allFiles.slice(i, i + opts.maxParallel));
        }
        
        for (const chunk of chunks) {
            const chunkPromises = chunk.map(file => processFile(file, results, opts));
            await Promise.all(chunkPromises);
        }
    } else {
        // 串行处理
        for (const file of allFiles) {
            await processFile(file, results, opts);
        }
    }
    
    // 统计结果
    const successCount = results.filter(r => r.success).length;
    const failCount = results.filter(r => !r.success).length;
    
    return {
        total: allFiles.length,
        success: successCount,
        fail: failCount,
        results: results
    };
}

/**
 * 批量解压主函数（支持多路径）
 * @param {string|Array<string>} inputPaths 输入路径（文件或目录，可以是字符串或数组）
 * @param {Object} options 选项
 */
async function batchExtractWordDocuments(inputPaths, options = {}) {
    // 统一转换为数组
    const paths = Array.isArray(inputPaths) ? inputPaths : [inputPaths];
    
    console.log('\n' + '='.repeat(70));
    console.log('🚀 批量解压 Word 文档工具 (多路径支持)');
    console.log('='.repeat(70));
    console.log(`📂 输入路径 (${paths.length} 个):`);
    paths.forEach((p, i) => {
        console.log(`   ${i + 1}. ${p}`);
    });
    console.log(`🔄 递归子目录: ${options.recursive !== false ? '是' : '否'}`);
    console.log(`⚡ 并行处理: ${options.parallel !== false ? '是' : '否'}`);
    console.log('='.repeat(70) + '\n');
    
    const startTime = Date.now();
    
    try {
        const stats = await processMultiplePaths(paths, options);
        
        const endTime = Date.now();
        const duration = ((endTime - startTime) / 1000).toFixed(2);
        
        console.log('\n' + '='.repeat(70));
        console.log('📊 执行完成统计');
        console.log('='.repeat(70));
        console.log(`📄 总文件数: ${stats.total}`);
        console.log(`✅ 成功: ${stats.success} 个`);
        console.log(`❌ 失败: ${stats.fail} 个`);
        console.log(`⏱️  总耗时: ${duration} 秒`);
        
        if (stats.fail > 0) {
            console.log('\n失败列表:');
            stats.results.filter(r => !r.success).forEach(r => {
                console.log(`  ${r.message}`);
            });
        }
        
        console.log('\n' + '='.repeat(70) + '\n');
        
        return stats;
    } catch (error) {
        console.error('❌ 执行过程中发生错误:', error);
        throw error;
    }
}

/**
 * 生成文件列表（支持通配符模式）
 * @param {Array<string>} patterns 文件模式数组
 * @returns {Array<string>} 匹配的文件列表
 */
function globFiles(patterns) {
    const { glob } = require('glob');
    const allFiles = [];
    
    for (const pattern of patterns) {
        const files = glob.sync(pattern, { absolute: true });
        allFiles.push(...files);
    }
    
    return [...new Set(allFiles)]; // 去重
}

/**
 * 命令行接口
 */
async function main() {
    // 获取命令行参数
    const args = process.argv.slice(2);
    
    if (args.length === 0) {
        console.log(`
使用方法:
  node extract-word.js <路径1> [路径2] [路径3] ... [选项]

参数:
  <路径>        一个或多个 Word 文件路径或包含 Word 文件的目录路径
  --no-recursive  不递归子目录（默认会递归）
  --serial        串行处理（默认并行）
  --quiet         静默模式（减少输出）
  --help, -h      显示帮助信息

通配符支持:
  node extract-word.js "./docs/*.docx"
  node extract-word.js "./**/*.docx"

示例:
  # 解压单个 Word 文件
  node extract-word.js document.docx

  # 解压多个文件
  node extract-word.js file1.docx file2.docx file3.docx

  # 解压目录下所有 Word 文件
  node extract-word.js ./documents ./reports

  # 混合使用文件和目录
  node extract-word.js ./documents special.docx ./reports

  # 使用通配符
  node extract-word.js "./docs/*.docx" "./**/report.docx"

  # 不递归子目录
  node extract-word.js ./documents --no-recursive

  # 串行处理（避免资源占用过高）
  node extract-word.js ./docs ./reports --serial

支持的 Word 格式: .docx, .docm, .dotx, .dotm
        `);
        return;
    }
    
    // 解析选项
    const options = {
        recursive: !args.includes('--no-recursive'),
        parallel: !args.includes('--serial'),
        quiet: args.includes('--quiet')
    };
    
    // 过滤掉选项参数，保留路径
    let paths = args.filter(arg => 
        !arg.startsWith('--') && arg !== '-h' && arg !== '--help'
    );
    
    // 检查是否需要使用 glob
    const hasGlobPattern = paths.some(p => p.includes('*') || p.includes('?'));
    
    if (hasGlobPattern) {
        try {
            // 尝试加载 glob 模块
            require.resolve('glob');
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
            console.log('💡 或者直接使用文件路径而不是通配符');
            return;
        }
    }
    
    await batchExtractWordDocuments(paths, options);
}

// 导出函数供其他模块使用
module.exports = {
    extractWordDoc,
    batchExtractWordDocuments,
    processMultiplePaths,
    isWordDocument,
    globFiles
};

// 如果直接运行脚本
if (require.main === module) {
    main();
}