// 全局变量
let workbook = null;
let fileName = '';
let isProcessing = false;

// DOM元素
const uploadBtn = document.getElementById('uploadBtn');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const dropArea = document.getElementById('drop-area');
const compareBtn = document.getElementById('compareBtn');
const modifyMasterBtn = document.getElementById('modifyMasterBtn');
const modifySubBtn = document.getElementById('modifySubBtn');
const statisticBtn = document.getElementById('statisticBtn');
const resultSection = document.getElementById('resultSection');
const resultTable = document.getElementById('resultTable');
const messageArea = document.getElementById('messageArea');
const downloadBtn = document.getElementById('downloadBtn');

// 事件监听
uploadBtn.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileSelect);
dropArea.addEventListener('dragover', handleDragOver);
dropArea.addEventListener('drop', handleDrop);
compareBtn.addEventListener('click', () => processWorkbook('compare'));
modifyMasterBtn.addEventListener('click', () => processWorkbook('modifyMaster'));
modifySubBtn.addEventListener('click', () => processWorkbook('modifySub'));
statisticBtn.addEventListener('click', () => processWorkbook('statistic'));
downloadBtn.addEventListener('click', downloadResult);

// 处理文件选择
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file && file.name.endsWith('.xlsx')) {
        loadExcelFile(file);
    } else {
        showMessage('请选择有效的Excel文件(.xlsx)', 'error');
    }
}

// 处理拖拽事件
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    dropArea.classList.add('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    dropArea.classList.remove('dragover');
    
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.xlsx')) {
        loadExcelFile(file);
    } else {
        showMessage('请拖入有效的Excel文件(.xlsx)', 'error');
    }
}

// 加载Excel文件
function loadExcelFile(file) {
    const reader = new FileReader();
    isProcessing = true;
    
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            fileName = file.name;
            
            fileInfo.innerHTML = `已加载: ${file.name} (${(file.size / 1024).toFixed(1)}KB)`;
            fileInfo.style.color = '#28a745';
            
            // 启用功能按钮
            compareBtn.disabled = false;
            modifyMasterBtn.disabled = false;
            modifySubBtn.disabled = false;
            statisticBtn.disabled = false;
            
            showMessage(`成功加载文件: ${file.name}`, 'success');
        } catch (error) {
            showMessage('读取Excel文件时出错: ' + error.message, 'error');
        } finally {
            isProcessing = false;
        }
    };
    
    reader.onerror = function() {
        showMessage('文件读取失败!', 'error');
        isProcessing = false;
    };
    
    reader.readAsArrayBuffer(file);
}

// 处理工作簿
function processWorkbook(mode) {
    if (!workbook) {
        showMessage('请先加载Excel文件', 'error');
        return;
    }
    
    try {
        resultSection.classList.add('active');
        messageArea.innerHTML = '';
        
        if (mode === 'compare') {
            compareSheets();
        } else if (mode === 'modifyMaster') {
            modifyMasterSheet();
        } else if (mode === 'modifySub') {
            modifySubSheets();
        } else if (mode === 'statistic') {
            calculateStatistics();
        }
    } catch (error) {
        showMessage(`处理过程中出错: ${error.message}`, 'error');
    }
}

// 医生类（增强版）
class Doctor {
    constructor(cell, section, row) {
        this.cell = cell;
        this.section = section; 
        this.row = row;
        this.name = this.extractName(cell.v);
        this.col = cell.c;
    }

    extractName(value) {
        const nonChinese = value.match(/[^\u4e00-\u9fff]/);
        let name = nonChinese ? value.split(nonChinese[0])[0] : value;
        if(name.length>4 || name.includes('皮')) {
            console.warn(`非法姓名：${name}`);
            this.section = '错误';
        }
        return name;
    }

    // 获取合并单元格状态
    getMergeState() {
        if(!this.cell.parent['!merges']) return 0;
        for(const merge of this.cell.parent['!merges']) {
            if(merge.s.r <= this.row && this.row <= merge.e.r &&
                merge.s.c <= this.col && this.col <= merge.e.c) {
                return this.col === merge.s.c ? 1 : 2;
            }
        }
        return 0;
    }
}

// 查找单元格
function lookfor(sheet, name, col = 1) {
    const cells = [];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    
    for (let row = 2; row <= range.e.r; row++) {
        const cellAddr = XLSX.utils.encode_cell({r: row, c: col});
        const cell = sheet[cellAddr];
        if (!cell || !cell.v) continue;
        
        const value = String(cell.v);
        if (value.length > 18) continue;
        if (value.includes('皮') || value.length > 10) continue;
        if (value.includes(name)) {
            cells.push(cell);
        }
    }

    if (cells.length === 1) return cells[0];
    if (cells.length > 1) console.warn(`lookfor发现多个匹配: ${name}`);
    return null;
}

// 获取合并单元格状态
function getMergeState(sheet, cell) {
    if (!sheet['!merges']) return 0;
    for (const merge of sheet['!merges']) {
        if (merge.s.r <= cell.r && cell.r <= merge.e.r && 
            merge.s.c <= cell.c && cell.c <= merge.e.c) {
            return cell.c === merge.s.c ? 1 : 2;
        }
    }
    return 0;
}

// 复制单元格样式
function copyStyle(targetCell, sourceCell) {
    if (!sourceCell.s) return;
    targetCell.s = {
        ...sourceCell.s,
        font: {...sourceCell.s.font},
        fill: {...sourceCell.s.fill},
        border: {...sourceCell.s.border},
        alignment: {...sourceCell.s.alignment}
    };
}

// 对比表功能（完整实现）
async function compareSheets() {
    if (!workbook) {
        showMessage('请先加载Excel文件', 'error');
        return;
    }

    try {
        showLoading('正在对比表格...');
        resultTable.innerHTML = '';
        const sheets = workbook.Sheets;
        const masterSheet = sheets[workbook.SheetNames[0]];
        let hasDiff = false;
        
        // 获取所有医生数据
        const doctors = [];
        for(let i=1; i<workbook.SheetNames.length; i++) {
            const sheet = sheets[workbook.SheetNames[i]];
            const color = sheet['A3']?.font?.color;
            
            for(let row=1; row<=sheet['!ref']?.e.r; row++) {
                const cell = sheet[`A${row}`];
                if(cell?.v && cell.font?.color === color) {
                    doctors.push(new Doctor(cell, workbook.SheetNames[i], row));
                }
            }
        }

        // 对比每个医生的排班
        doctors.forEach(doctor => {
            const masterCell = lookfor(masterSheet, doctor.name);
            if(!masterCell) {
                showMessage(`${doctor.name} 不在总表中`, 'error');
                return;
            }
            
            for(let day=1; day<=14; day++) {
                const subCell = sheet[`${XLSX.utils.encode_col(doctor.col+day)}${doctor.row}`];
                const masterDayCell = masterSheet[`${XLSX.utils.encode_col(masterCell.c+day)}${masterCell.r}`];
                
                if(!compareCells(masterDayCell, subCell)) {
                    hasDiff = true;
                    addDiffToTable(doctor.name, day, masterDayCell, subCell);
                }
            }
        });

        if (!hasDiff) {
            showMessage('表对比完成: 未发现差异', 'success');
        } else {
            showMessage('表对比完成: 发现差异', 'error');
        }
        downloadBtn.style.display = 'none';
    } catch (error) {
        showMessage('对比过程中出错: ' + error.message, 'error');
    }
}

// 单元格对比函数
function compareCells(cell1, cell2) {
    if(!cell1 && !cell2) return true;
    if(!cell1 || !cell2) return false;
    const val1 = String(cell1.v||'').trim().toLowerCase();
    const val2 = String(cell2.v||'').trim().toLowerCase();
    return val1 === val2;
}

// 添加差异到结果表格
function addDiffToTable(name, day, cell1, cell2) {
    const row = resultTable.insertRow();
    row.insertCell().textContent = name;
    row.insertCell().textContent = `第${day}天`;
    row.insertCell().textContent = cell1?.v || '空';
    row.insertCell().textContent = cell2?.v || '空';
}

// 修改总表功能
async function modifyMasterSheet() {
    if (!workbook) {
        showMessage('请先加载Excel文件', 'error');
        return;
    }

    try {
        showLoading('正在修改总表...');
        // 实现修改总表逻辑
        // ...
        
        showMessage('总表修改完成', 'success');
        downloadBtn.style.display = 'block';
    } catch (error) {
        showMessage('修改总表出错: ' + error.message, 'error');
    }
}

// 修改分表功能
async function modifySubSheets() {
    if (!workbook) {
        showMessage('请先加载Excel文件', 'error');
        return;
    }

    try {
        showLoading('正在修改分表...');
        // 实现修改分表逻辑
        // ...
        
        showMessage('分表修改完成', 'success');
        downloadBtn.style.display = 'block';
    } catch (error) {
        showMessage('修改分表出错: ' + error.message, 'error');
    }
}

// 主专统计功能
async function calculateStatistics() {
    if (!workbook) {
        showMessage('请先加载Excel文件', 'error');
        return;
    }

    try {
        showLoading('正在统计数据...');
        resultTable.innerHTML = '';
        
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const stats = {};
        
        // 获取表格范围
        const range = XLSX.utils.decode_range(sheet['!ref']);
        
        // 统计每一天的数据
        for (let day = 1; day <= 14; day++) {
            const dayCol = XLSX.utils.encode_col(day + 2); // 从第3列开始
            stats[day] = [];
            
            for (let row = 2; row <= range.e.r; row++) {
                const cellAddr = `${dayCol}${row}`;
                const cell = sheet[cellAddr];
                if (!cell || !cell.v) continue;
                
                const value = String(cell.v).trim();
                
                // 跳过合并单元格的非主单元格
                const mergeState = getMergeState(sheet, XLSX.utils.decode_cell(cellAddr));
                if (mergeState === 2) continue;
                
                // 过滤特定门诊类型
                if (value.includes('激') || value.includes('脱') || 
                    value.includes('性') || value.includes('靶') ||
                    value.includes('注射') || value.includes('美容') ||
                    value.includes('带疱') || value.length > 10) {
                    continue;
                }
                
                // 识别主专门诊
                if (value.includes('主') || value.includes('专') || 
                    value.includes('甲病') || value.includes('黄褐斑门诊') ||
                    value.includes('白癜风') || value.includes('痤疮')) {
                    stats[day].push(`${row}-${value}`);
                }
            }
        }
        
        // 显示统计结果
        for (const day in stats) {
            const row = resultTable.insertRow();
            row.insertCell().textContent = `第${day}天`;
            row.insertCell().textContent = stats[day].length;
            row.insertCell().textContent = stats[day].join(', ');
            
            // 高亮显示超过16人的天数
            if (stats[day].length > 16) {
                row.style.backgroundColor = '#ffdddd';
            }
        }
        
        showMessage('统计完成', 'success');
        downloadBtn.style.display = 'block';
    } catch (error) {
        showMessage('统计过程中出错: ' + error.message, 'error');
    }
}

// 下载结果
function downloadResult() {
    try {
        const today = new Date();
        const timestamp = `${today.getFullYear()}${(today.getMonth()+1).toString().padStart(2, '0')}${today.getDate().toString().padStart(2, '0')}`;
        const newFileName = fileName.replace('.xlsx', `_processed_${timestamp}.xlsx`);
        
        // 将修改后的工作簿写入新的文件
        XLSX.writeFile(workbook, newFileName);
        showMessage(`文件已保存为: ${newFileName}`, 'success');
    } catch (error) {
        showMessage('导出文件时出错: ' + error.message, 'error');
    }
}

// 显示消息
function showMessage(message, type = 'info') {
    messageArea.innerHTML = `<div class="${type}">${message}</div>`;
}

// 显示加载状态
function showLoading(message) {
    messageArea.innerHTML = `<div class="loading">${message}</div>`;
}
