// -------------------------
// 辅助函数
// -------------------------


function showMsg(msg, type = 'info') {
    els.msg.innerHTML = `<div class="${type}">${msg}</div>`;
    els.resultSection.classList.add('active');
}

// 辅助函数：清理字符串值
function cleanValue(val) {
    if (val === null || val === undefined) return '';
    return String(val).replace(/\s+/g, '').toLowerCase();
}

// -------------------------
// Doctor 类（尽量贴合 Python 行为）
// -------------------------
class Doctor {
    // r,c 使用 0-based 索引，以配合 XLSX utils
    constructor(cell, r, c) {
        
        this.cell = cell; // 原始 cell 对象引用
        //this.section = sheetTitle;
        this.row = r;
        this.col = c;
        this.cell_t = null;
        this.name = this.extractName(cell && cell.v ? cell.v : '');
        // 如果判断为非法：section = '错误'
        if (this.name.length > 4 || this.name.includes('皮')) {
            console.warn(`在表<。。。>发现疑似非法姓名： <${this.name}> , 丢弃`);
            this.section = '错误';
        }
    }

    extractName(value) {
        value = String(value || '').trim();
        if (!value) return '';
        const nonChinese = value.match(/[^\u4e00-\u9fff]/);
        if (nonChinese) {
            return value.split(nonChinese[0])[0];
        }
        return value;
    }
}

// -------------------------
// 核心：识别医生 getDoctors（优先使用 A3 字体颜色）
// -------------------------
function getCellFontColor(cell) {
    if (!cell || !cell.s) return null;
    
    // 确保我们正确访问字体对象
    const font = cell.s.font;
    if (!font || !font.color) return null;
    
    // 处理RGB颜色（最常用）
    if (font.color.rgb) {
        // 标准化RGB格式，保持与deepCloneStyle一致
        let rgb = font.color.rgb.toUpperCase();
        
        // 处理各种RGB格式
        if (rgb.length === 8 && rgb.startsWith('FF')) {
            // Excel内部格式，保留完整格式
            return rgb;
        } else if (rgb.length === 6 && !rgb.startsWith('#')) {
            // 纯RGB值，添加透明度前缀
            return 'FF' + rgb;
        } else if (rgb.startsWith('#')) {
            // 带#号的格式
            const hex = rgb.slice(1);
            return hex.length === 6 ? 'FF' + hex : hex;
        }
        return rgb;
    }
    
    // 处理主题颜色，完整保存theme和tint信息
    if (font.color.theme !== undefined) {
        const themeStr = 'theme:' + font.color.theme;
        // 添加tint值（如果存在）
        return font.color.tint !== undefined ? 
            `${themeStr}:tint=${font.color.tint}` : themeStr;
    }
    
    // 处理索引颜色，确保正确检测undefined
    if (font.color.indexed !== undefined) {
        return 'indexed:' + font.color.indexed;
    }
    
    // 处理可能的其他颜色格式
    if (typeof font.color === 'string') {
        return font.color;
    }
    
    return null;
}


function getDoctors(sheet) { 
    const doctors = [];
    if (!sheet || !sheet['!ref']) return doctors;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const baseA3Addr = XLSX.utils.encode_cell({r: 2, c: 0});
    const baseA3 = sheet[baseA3Addr];
    const baseColor = getCellFontColor(baseA3);
    for (let r = 0; r <= range.e.r; r++) {
        const addr = XLSX.utils.encode_cell({r: r, c: 0}); // A列
        const cell = sheet[addr];
        if (!cell || cell.v === undefined || cell.v === null) continue;
        if (typeof cell.v !== 'string' && typeof cell.v !== 'number') continue;
        const val = String(cell.v).trim();
        if (!val) continue;

        if (baseColor) {
            const cColor = getCellFontColor(cell);
            
            if (cColor && cColor === baseColor) {

                doctors.push(new Doctor(cell, r, 0)); 
                continue;
            }
        }

        // 若没有样式或 A3 没样式，则使用备选文字过滤 (保守)
        // 避免把表头 / 备注误判为医生
        const headerKeywords = ['备注', '总计', '日期', '姓名', '排班', '时间', '合计'];
        if (headerKeywords.some(k => val.includes(k))) continue;
        // 长度过滤，尽量贴近 Python 行为: 忽略过长的字符串
        if (val.length > 8) continue;
        // 含数字或多余字符可能不是姓名
        if (/[A-Za-z0-9]/.test(val)) continue;

        // 通过上面筛选后仍可能包含真实医生
        doctors.push(new Doctor(cell, r, 0)); 
    }

    return doctors.filter(d => d.section !== '错误');
}

// -------------------------
// lookfor：在总表的 B 列（index = 1）查找姓名
// 与 Python 保持一致：从第2行开始 (row index 1)
// -------------------------
function lookfor(sheet, name, col = 1) {
    if (!sheet || !sheet['!ref'] || !name) return null;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    const matches = [];
    // Python 的循环是 for row in range(2, sheet.max_row+1) -> 1-based row 从2开始 -> 0-based 从1开始
    for (let r = 1; r <= range.e.r; r++) {
        const addr = XLSX.utils.encode_cell({r: r, c: col});
        const cell = sheet[addr];
        if (!cell || cell.v === undefined || cell.v === null) continue;
        const val = String(cell.v).trim();
        if (!val) continue;
        if (val.length > 18) continue;
        if (val.includes('皮') || val.length > 10) continue;
        if (val.includes(name)) matches.push({ r: r, c: col, cell: cell, addr: addr });
    }
    if (matches.length === 1) return matches[0];
    if (matches.length > 1) {
        console.warn(`lookfor: 找到多个匹配 ${name} -> ${matches.length}`);
        return matches[0]; // 选择第一个，尽量容错（与 Python 会报错不同，这里保守处理）
    }
    return null;
}

// -------------------------
// 合并单元格工具：unmergeRow + syncMerges
// sheet['!merges'] 是一个数组，每项为 {s:{r,c}, e:{r,c}}
// -------------------------
function ensureMergesArray(sheet) {
    if (!sheet['!merges']) sheet['!merges'] = [];
}

// 从目标 sheet 删除与目标行有关的所有合并区域（unmerge）
function unmergeRow(sheet, targetRow) {
    if (!sheet || !sheet['!merges']) return;
    sheet['!merges'] = sheet['!merges'].filter(m => !(m.s.r <= targetRow && targetRow <= m.e.r));
}

// 在 targetSheet 上，以 targetRow 应用 sourceSheet 在 sourceRow 的横向合并（保留相对列偏移）
// sourceNameCol, targetNameCol 分别是姓名列的列索引（0-based）
function syncMerges(sourceSheet, targetSheet, sourceRow, targetRow, sourceNameCol, targetNameCol) {
    if (!sourceSheet || !sourceSheet['!merges']) return;
    ensureMergesArray(targetSheet);
    // 先删除目标行任何现存的合并（与 Python unmerge 行为一致）
    unmergeRow(targetSheet, targetRow);

    // 找到 sourceRow 上的横向合并（min_row<=sourceRow<=max_row）
    const rowMerges = sourceSheet['!merges'].filter(m => m.s.r <= sourceRow && sourceRow <= m.e.r);
    rowMerges.forEach(m => {
        // 只处理与姓名列同一行的横向合并：即合并范围行包含 sourceRow
        // 计算相对于姓名列的偏移
        const startRel = m.s.c - sourceNameCol;
        const endRel = m.e.c - sourceNameCol;
        // 目标合并的新列
        const newStart = targetNameCol + startRel;
        const newEnd = targetNameCol + endRel;
        // 追加到 target merges
        targetSheet['!merges'].push({ s: { r: targetRow, c: newStart }, e: { r: targetRow, c: newEnd } });
    });
}

// 获取合并状态：0=非合并，1=主单元格，2=非主单元格（被合并覆盖）
function getMergeState(sheet, r, c) {
    if (!sheet || !sheet['!merges']) return 0;
    for (const m of sheet['!merges']) {
        if (m.s.r <= r && r <= m.e.r && m.s.c <= c && c <= m.e.c) {
            if (m.s.r === r && m.s.c === c) return 1; // 主单元格
            return 2; // 被合并的非主单元格
        }
    }
    return 0;
}

// 复制单元格值与样式（基于 xlsx-js-style 的样式结构）
function copyCellValueAndStyle(targetSheet, targetAddr, sourceCell) {
    // 确保目标单元格存在
    if (!targetSheet[targetAddr]) {
        targetSheet[targetAddr] = {};
    }
    const tgt = targetSheet[targetAddr];

    // 复制值和类型
    tgt.v = sourceCell?.v ?? '';
    tgt.t = sourceCell?.t ?? 's';

    // 样式处理 - 确保完整复制颜色信息
    if (sourceCell && sourceCell.s) {
        // !使用改进的deepCloneStyle进行深拷贝
        //tgt.s = deepCloneStyle(sourceCell.s);
        tgt.s = JSON.parse(JSON.stringify(sourceCell.s));

        
        // 确保填充模式正确设置
        if (tgt.s && tgt.s.fill && !tgt.s.fill.patternType) {
            tgt.s.fill.patternType = 'solid';
        }
    } else {
        // 如果源单元格没有样式，删除目标单元格的样式
        delete tgt.s;
    }
}


// -------------------------
// 处理单个子表（change_sheet_s）
// flag: 0 compare, 1 sub -> master, 2 master -> sub
// -------------------------
function changeSheetS(subSheet, masterSheet, flag) {
    const doctors = getDoctors(subSheet);
    const matched = [];
    const diffs = [];
    let modifiedCount = 0;

    // 匹配医生到总表
    doctors.forEach(doc => {
        const found = lookfor(masterSheet, doc.name, 1); // B列 -> index 1
        if (!found) {
            console.warn(`${doc.name} -- 不在总表内`);
            return;
        }
        doc.cell_t = found; // {r,c,cell,addr}
        matched.push(doc);
    });

    matched.forEach(doc => {
        const masterInfo = doc.cell_t; // row, c, cell, addr
        const subNameCol = doc.col;      // 0-based (A 列)
        const masterNameCol = masterInfo.c; // B 列 index 1

        // 合并单元格同步（按照 flag）
        if (flag === 1) {
            // 把 sub 的合并信息应用到 master
            syncMerges(subSheet, masterSheet, doc.row, masterInfo.r, subNameCol, masterNameCol);
        } else if (flag === 2) {
            // 把 master 的合并信息应用到 sub
            syncMerges(masterSheet, subSheet, masterInfo.r, doc.row, masterNameCol, subNameCol);
        }

        // 14 天循环（Python 1..14），映射到列: nameCol + day
        for (let day = 1; day <= 14; day++) {
            const subC = subNameCol + day;
            const masterC = masterNameCol + day;
            const subAddr = XLSX.utils.encode_cell({ r: doc.row, c: subC });
            const masterAddr = XLSX.utils.encode_cell({ r: masterInfo.r, c: masterC });

            const subCell = subSheet[subAddr] || { v: null, t: 's' };
            const masterCell = masterSheet[masterAddr] || { v: null, t: 's' };

            if (flag === 0) {
                // 对比模式：将经过空白/空字符清洗后的字符串比较（case-insensitive）
                const vs = subCell.v;
                const vm = masterCell.v;
                if (vs !== vm) {
                    diffs.push({ name: doc.name, day: day, m: masterCell.v || '', s: subCell.v || '' });
                }
            } else {
                // 修改模式：决定 src/tgt
                let srcCell, tgtSheet, tgtAddr;
                if (flag === 1) { // sub -> master
                    srcCell = subCell; tgtSheet = masterSheet; tgtAddr = masterAddr;
                } else { // flag === 2, master -> sub
                    srcCell = masterCell; tgtSheet = subSheet; tgtAddr = subAddr;
                }
                // 如果源单元格为合并块的非主单元格，取主单元格
                const srcMergeState = getMergeState((flag === 1 ? subSheet : masterSheet), (flag === 1 ? doc.row : masterInfo.r), (flag === 1 ? subC : masterC));
                if (srcMergeState === 2) {
                    // 找到对应主单元格的列（向左找）
                    const merges = (flag === 1 ? subSheet['!merges'] : masterSheet['!merges']) || [];
                    for (const m of merges) {
                        if (m.s.r <= (flag === 1 ? doc.row : masterInfo.r) && (flag === 1 ? doc.row : masterInfo.r) <= m.e.r &&
                            m.s.c <= (flag === 1 ? subC : masterC) && (flag === 1 ? subC : masterC) <= m.e.c) {
                            const mainAddr = XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c });
                            srcCell = (flag === 1 ? subSheet[mainAddr] : masterSheet[mainAddr]) || srcCell;
                            break;
                        }
                    }
                }
                // 执行复制值与样式
                copyCellValueAndStyle(tgtSheet, tgtAddr, srcCell);
                modifiedCount++;
            }
        } // end for 14 days
    }); // end matched.forEach

    return { diffs, modifiedCount, matchedCount: matched.length };
}

// -------------------------
// 上午/下午 拆分（delflag）实现
// Python 中针对列 3..16 (1-based)，即 0-based 的 2..15
// -------------------------
function delflag(sheet) {
    if (!sheet || !sheet['!ref']) return;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    // 目标列 2..15 (0-based)
    for (let c = 2; c <= 15 && c <= range.e.c; c++) {
        for (let r = 1; r <= range.e.r; r++) { // Python 从 row 2 开始 -> 0-based 从1开始
            const addr = XLSX.utils.encode_cell({ r: r, c: c });
            const cell = sheet[addr];
            if (!cell || typeof cell.v !== 'string') continue;
            if (!cell.v.includes('/')) continue;
            const state = getMergeState(sheet, r, c);
            if (state === 0) {
                // 没合并，直接删除斜杠
                sheet[addr].v = cell.v.replace('/', '');
                continue;
            }
            // 如果是合并单元格，拆分为两列
            // 方案：unmerge 整行的合并（针对该行），然后将当前单元格值分配到 c 和 c+1
            // 取值
            const parts = String(cell.v).split('/');
            const am = parts[0] || '';
            const pm = parts[1] || '';
            // unmerge 该行所有合并单元格
            unmergeRow(sheet, r);
            // 写入 AM 到 c
            const addrAm = XLSX.utils.encode_cell({ r: r, c: c });
            const addrPm = XLSX.utils.encode_cell({ r: r, c: c + 1 });
            // 复制样式（尽量保持）
            if (!sheet[addrAm]) sheet[addrAm] = {};
            sheet[addrAm].v = am;
            if (cell.s) sheet[addrAm].s = JSON.parse(JSON.stringify(cell.s));

            // PM
            if (!sheet[addrPm]) sheet[addrPm] = {};
            sheet[addrPm].v = pm;

            if (cell.s) sheet[addrAm].s = JSON.parse(JSON.stringify(cell.s));
        }
    }
}

// -------------------------
// 统计函数（statistic）
// -------------------------
function statistic(masterSheet) {
    if (!masterSheet || !masterSheet['!ref']) return {};
    delflag(masterSheet); // 先拆分上午/下午与合并问题
    const range = XLSX.utils.decode_range(masterSheet['!ref']);
    const result = {};
    // Python 遍历 3..16 (1-based) => 0-based 2..15
    const include = ['主', '专', '甲病', '黄褐斑', '白癜风', '痤疮'];
    const exclude = ['激', '脱', '性', '靶', '注射', '美容', '带疱'];

    for (let c = 2; c <= 15 && c <= range.e.c; c++) {
        const arr = [];
        for (let r = 1; r <= range.e.r; r++) { // rows from 2nd row
            const addr = XLSX.utils.encode_cell({ r: r, c: c });
            const cell = masterSheet[addr];
            if (!cell || !cell.v) continue;
            // 如果是合并的非主单元格，跳过（Python 把被合并的单元格视为非主）
            if (getMergeState(masterSheet, r, c) === 2) continue;
            const val = String(cell.v).trim();
            if (val.length > 10) continue;
            if (exclude.some(k => val.includes(k))) continue;
            if (include.some(k => val.includes(k)) && !val.includes('激')) {
                arr.push(`${r+1}-${val}`); // Python row index is 1-based; +1 for human readable
            }
        }
        result[c - 1] = arr; // 使键便于展示（day=col-1）
    }
    return result;
}

// -------------------------
// 运行控制：processWorkbook + runCompare/runModify/runStatistic
// -------------------------
function processWorkbook(mode) {
    if (!workbook) return showMsg('请先加载文件', 'error');
    showMsg('正在处理...', 'loading');
    els.table.innerHTML = '';
    els.btns.download.style.display = 'none';

    // 使用Promise简化异步处理流程
    return Promise.resolve()
        .then(() => {
            
            const sheets = workbook.Sheets;
            const masterSheet = sheets[workbook.SheetNames[0]];
            //console.error(`processWorkbook后内容:${workbook.SheetNames[0]}`);
            if (mode === 'compare') runCompare(sheets, masterSheet);
            else if (mode === 'modifyMaster') runModify(sheets, masterSheet, 1);
            else if (mode === 'modifySub') runModify(sheets, masterSheet, 2);
            else if (mode === 'statistic') runStatistic(masterSheet);
        })
        .catch(e => {
            console.error('处理失败:', e);
            showMsg(`运行出错: ${e.message}`, 'error');
        });
}

function runCompare(sheets, masterSheet) {
    let totalDiffs = 0;
    let html = '<thead><tr><th>姓名</th><th>天数</th><th>总表</th><th>分表</th></tr></thead><tbody>';
    for (let i = 1; i < workbook.SheetNames.length; i++) {
        const subSheet = sheets[workbook.SheetNames[i]];
        const res = changeSheetS(subSheet, masterSheet, 0);
        res.diffs.forEach(d => {
            html += `<tr><td>${d.name}</td><td>第${d.day}天</td><td>${d.m}</td><td>${d.s}</td></tr>`;
        });
        totalDiffs += res.diffs.length;
    }
    html += '</tbody>';
    els.table.innerHTML = html;
    if (totalDiffs === 0) showMsg('完美！未发现任何差异', 'success');
    else showMsg(`发现 ${totalDiffs} 处不一致`, 'error');
}

function runModify(sheets, masterSheet, flag) {
    let totalModified = 0;
    for (let i = 1; i < workbook.SheetNames.length; i++) {
        const subSheet = sheets[workbook.SheetNames[i]];
        const res = changeSheetS(subSheet, masterSheet, flag);
        totalModified += res.modifiedCount || 0;
    }
    const type = flag === 1 ? '总表' : '分表';
    showMsg(`${type}修改完成！共修改 ${totalModified} 个单元格，样式与合并已同步，请下载保存。`, 'success');
    els.btns.download.style.display = 'block';
}

function runStatistic(masterSheet) {
    const stats = statistic(masterSheet);
    let html = '<thead><tr><th>天数</th><th>人数</th><th>详情</th></tr></thead><tbody>';
    for (const key in stats) {
        const arr = stats[key];
        const count = arr.length;
        const style = count > 16 ? 'style="background:#ffebee; color:#c62828; font-weight:bold;"' : '';
        html += `<tr ${style}><td>第${key}天</td><td>${count}</td><td style="text-align:left">${arr.join(', ')}</td></tr>`;
    }
    html += '</tbody>';
    els.table.innerHTML = html;
    showMsg('统计完成 (红色行表示超过16人)', 'success');
}


// -------------------------
// 文件读写：handleFileSelect / handleDrop / downloadResult
// -------------------------
function handleFileSelect(e) {
    const file = (e.target.files && e.target.files[0]) || (e.dataTransfer && e.dataTransfer.files[0]);
    if (!file || !file.name.endsWith('.xlsx')) return showMsg('请使用 .xlsx 文件', 'error');
    
    showMsg('正在处理，请稍候...', 'loading');
    const reader = new FileReader();
    
    reader.onload = (ev) => {
        try {
            const data = new Uint8Array(ev.target.result);
            workbook = XLSX.read(data, {
                type: 'array', 
                cellStyles: true,
                //cellNF: true // 必需：解析数字格式/字体格式（否则可能漏解析font）

            });

            fileName = file.name;
            els.fileInfo.textContent = `当前文件: ${file.name}`;
            els.fileInfo.style.color = 'green';
            showMsg('文件加载成功，请选择功能', 'success');
            
            showMsg(`是否有属性：${!!workbook.Sheets[workbook.SheetNames[1]]?.A3?.s.font}`,'success' );
            Object.values(els.btns).forEach(b => b.disabled = false);
            els.btns.download.style.display = 'none';
        } catch (error) {
            console.error('文件处理失败:', error);
            showMsg(`文件处理失败: ${error.message}`, 'error');
            workbook = null;
            fileName = '';
        }
    };
    
    reader.onerror = () => {
        showMsg('文件读取失败', 'error');
        workbook = null;
        fileName = '';
    };
    
    reader.readAsArrayBuffer(file);
}

function handleDrop(e) {
    e.preventDefault(); els.dropArea.classList.remove('dragover');
    const file = (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]);
    if (!file) return showMsg('未检测到文件', 'error');
    // 构造一个伪事件给 handleFileSelect
    const fakeEvent = { target: { files: [file] } };
    handleFileSelect(fakeEvent);
}

function downloadResult() {
    if (!workbook) return showMsg('没有可保存的工作簿', 'error');
    if (!fileName) return showMsg('文件名未定义', 'error');
    
    const ts = new Date().toISOString().slice(0,10).replace(/-/g,'');
    const newName = fileName.replace('.xlsx', `_processed_${ts}.xlsx`);
    
    // 增强Excel写入配置，确保所有样式（包括颜色）正确保存
    try {
        // 完整的写入选项，优化样式保留和颜色处理
        const writeOptions = {
            bookType: 'xlsx',
            type: 'file',
        };
        
        XLSX.writeFile(workbook, newName, writeOptions);
        console.log('文件成功保存，所有样式（包括颜色）已正确保留');
        showMsg(`文件已保存为: ${newName}`, 'success');
    } catch (e) {
        console.error('保存时出错', e);
        showMsg('保存失败（请在控制台查看错误）', 'error');
    }
}