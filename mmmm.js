const matched = [];
let doctors = [];

function showMsg(msg, type = 'info') {//要显示的消息。
    if (!els || !els.msg) {
        console.log(`[${type}] ${msg}`);
        return;
    }
    els.msg.innerHTML = `<div class="${type}">${msg}</div>`;
    if (els.resultSection) els.resultSection.classList.add('active');
}

function cleanValue(val) {//清洗字符串值，去除空白并转换为小写。
    if (val === null || val === undefined) return '';
    return String(val).replace(/\s+/g, '').toLowerCase();
}

function colLetterToIndex(letter) {  //将 Excel 列字母转换为 0 基础索引。
    // A -> 1, B -> 2 ... Z -> 26, AA -> 27 ...
    let col = 0;
    for (let i = 0; i < letter.length; i++) {
        col = col * 26 + (letter.charCodeAt(i) - 64);
    }
    return col - 1; // return 0-based
}

function indexToColLetter(index) {  //将 0 基础索引转换为 Excel 列字母。
    // 0 -> A
    let n = index + 1;
    let s = '';
    while (n > 0) {
        let m = (n - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

function a1ToRC(a1) {// 将 Excel A1 格式转换为行和列索引。
    // like "B12" -> {r: 11 (0-based), c: 1}
    const m = a1.match(/^([A-Z]+)(\d+)$/i);
    if (!m) return null;
    const col = colLetterToIndex(m[1].toUpperCase());
    const row = parseInt(m[2], 10) - 1;
    return { r: row, c: col };
}

function rcToA1(r, c) {  //将行和列索引转换为 Excel A1 格式。
    return `${indexToColLetter(c)}${r + 1}`;
}

function decodeRange(rangeStr) {  //解码 Excel 范围字符串。
    // "A1:C3" -> {s:{r,c}, e:{r,c}}

    if (!rangeStr.includes(':')) {
        const a = a1ToRC(rangeStr);
        return { s: a, e: a };
    }
    const parts = rangeStr.split(':');
    const s = a1ToRC(parts[0]);
    const e = a1ToRC(parts[1]);
    
    return { s, e };
}

function getWorksheetMergeRanges(ws) {
    // 获取工作表中的所有合并范围，返回数组：rangeStr，如 ["A1:C1", "E2:E3", ...]
    if (!ws) {
        console.warn('getWorksheetMergeRanges：工作表 ws 不存在');
        return [];
    }

    try {
        
        const mergedRanges = ws.model.merges;
        // 将 MergeRange 对象转为范围字符串（如 MergeRange → "A1:C1"）

        return mergedRanges;
        return mergedRanges.map(range => range.address);
    } catch (e) {
        // 🌟 修正：输出具体错误日志，方便排查
        console.error('getWorksheetMergeRanges：获取合并范围失败', e.message);
        return [];
    }
}

function isCellInRange(r, c, rangeStr) {   //检查单元格是否在指定范围内。
    const range = decodeRange(rangeStr);
    return (r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c);
}

function isCellMasterInRange(r, c, rangeStr) {   //检查单元格是否是指定范围的主单元格。
    const range = decodeRange(rangeStr);
    return (r === range.s.r && c === range.s.c);
}

function getMergeState(ws, r, c) {   //获取单元格的合并状态。
    const ranges = getWorksheetMergeRanges(ws);
    for (const range of ranges) {
        if (isCellInRange(r, c, range)) {
            if (isCellMasterInRange(r, c, range)) return 1;
            return 2;
        }
    }
    return 0;
}


function deepClone(obj) {// 通用深度克隆函数（必须保留，否则样式嵌套对象会浅复制）
    if (obj === null || typeof obj !== "object") return obj;
    if (obj instanceof Date) return new Date(obj.getTime());
    if (obj instanceof Array) return obj.map(item => deepClone(item));
    const cloneObj = {};
    for (const key in obj) {
        if (obj.hasOwnProperty(key)) {
            cloneObj[key] = deepClone(obj[key]);
        }
    }
    return cloneObj;
}

function copyCellValueAndStyleExcelJS(targetCell, sourceCell) {//单元格复制，核心函数。
    // 边界校验：目标单元格不存在直接返回
    if (!targetCell) return;

    // 🌟 第一步：深度复制「值」（按类型处理，重点支持富文本）
    if (!sourceCell) {
        // 源单元格不存在：清空目标单元格的值和所有样式
        targetCell.value = null;
        ['font', 'fill', 'border', 'alignment', 'numFmt'].forEach(key => delete targetCell[key]);
        return;
    }

    const sourceVal = sourceCell.value;
    if (sourceVal === undefined || sourceVal === null) {
        targetCell.value = null;
    } else {
        // 按值类型针对性复制，保留原数据结构
        if (sourceVal.richText && Array.isArray(sourceVal.richText)) {
            // 1. 富文本：深度克隆 richText 数组及内部 font 样式（保留原有正确逻辑）
            targetCell.value = {
                richText: sourceVal.richText.map(segment => ({
                    ...segment, // 克隆文本及其他段落属性
                    font: segment.font ? deepClone(segment.font) : undefined // 段落级字体样式
                }))
            };
        } else if (sourceVal.formula) {
            // 2. 公式：克隆 formula 和 result（保留可计算性）
            targetCell.value = deepClone(sourceVal);
        } else if (sourceVal instanceof Date) {
            // 3. 日期：克隆时间戳（避免引用冲突）
            targetCell.value = new Date(sourceVal.getTime());
        } else if (typeof sourceVal === 'object') {
            // 4. 其他对象/数组：深度克隆
            targetCell.value = deepClone(sourceVal);
        } else {
            // 5. 基础类型：直接赋值
            targetCell.value = sourceVal;
        }
    }

    // 🌟 第二步：补充「单元格全局样式」深度复制（核心修正：新增这部分）
    const globalStyles = ['fill', 'border', 'alignment', 'numFmt', 'font'];
    globalStyles.forEach(styleKey => {
        const sourceStyle = sourceCell[styleKey];
        if (sourceStyle) {
            // 深度克隆样式（避免引用冲突，numFmt是字符串/数字，直接赋值即可）
            targetCell[styleKey] = styleKey === 'numFmt' 
                ? sourceStyle 
                : deepClone(sourceStyle);
        } else {
            // 源单元格无该样式：删除目标单元格的旧样式（避免残留）
            delete targetCell[styleKey];
        }
    });
}

function unmergeRowExcelJS(ws, targetRow) {// 在目标工作表上删除包含目标行的所有合并。
    if (!ws) return;
    const rowNumber = targetRow + 1;
    const ranges = getWorksheetMergeRanges(ws);
    for (const range of ranges) {
        const dec = decodeRange(range);
        if (rowNumber >= dec.s.r + 1 && rowNumber <= dec.e.r + 1) {
            try {
                ws.unMergeCells(range);
            } catch (e) {
                // ignore
            }
        }
    }
}

/**
 * 将源工作表中的源行的横向合并复制到目标工作表的目标行。
 * @param {object} sourceSheet - 源工作表对象。
 * @param {object} targetSheet - 目标工作表对象。
 * @param {number} sourceRow - 0 基础源行索引。
 * @param {number} targetRow - 0 基础目标行索引。
 * @param {number} sourceNameCol - 源工作表中姓名列的 0 基础列索引。
 * @param {number} targetNameCol - 目标工作表中姓名列的 0 基础列索引。
 */
function syncMergesExcelJS(sourceSheet, targetSheet, sourceRow, targetRow, sourceNameCol, targetNameCol) {//同步 合并。
    if (!sourceSheet || !sourceSheet._merges) return;
    // 清理目标行上的合并
    unmergeRowExcelJS(targetSheet, targetRow);

    const ranges = getWorksheetMergeRanges(sourceSheet);//获取工作表 合并范围
    for (const rangeStr of ranges) {
        const dec = decodeRange(rangeStr);
        // 检查 sourceRow 是否处于该合并区间的行范围
        if (sourceRow >= dec.s.r && sourceRow <= dec.e.r) {
            // 计算相对于姓名列的偏移
            const startRel = dec.s.c - sourceNameCol;
            const endRel = dec.e.c - sourceNameCol;
            const newStartCol = targetNameCol + startRel;
            const newEndCol = targetNameCol + endRel;
            // ExcelJS mergeCells 参数是 (top,left,bottom,right) with 1-based indexes
            try {
                targetSheet.mergeCells(targetRow + 1, newStartCol + 1, targetRow + 1, newEndCol + 1);
            } catch (e) {
                console.warn('mergeCells failed', e, rangeStr);
            }
        }
    }
}

class Doctor {//医生类
    constructor(cell, r) {
        this.row = r;

        //cellString单元格分表
        this.cellString = cell && cell.value !== undefined && cell.value !== null ? String(cell.value).trim() : '';
        this.cell_s=cell;
        this.name = this.extractName(this.cellString);
        this.cell_m = null; //cell_m单元格总表
        this.section = cell.worksheet.name;
        if (this.name.length > 4 || this.name.includes('皮')) {
            console.warn(`在表<${cell.worksheet.name}>发现疑似非法姓名： <${this.name}> , 丢弃`);
            this.section = '错误';
        }
    }
    extractName(value) {//去除非中文后的姓名
        value = String(value || '').trim();
        if (!value) return '';
        const nonChinese = value.match(/[^\u4e00-\u9fff]/);
        if (nonChinese) {
            return value.split(nonChinese[0])[0];
        }
        return value;
    }
}

function isTextBoldLikeMarker(val) {
    if (!val) return false;
    const s = String(val).trim();
    return s.startsWith('【') && s.endsWith('】');
}

function getDoctorsExcelJS(worksheet) {
    if (!worksheet) {
    console.warn('获取医生列表失败：工作表不存在', 'error');
    return;
}
    const rowCount = worksheet.rowCount || worksheet.actualRowCount || 0;
    // 检查 A3 (r=2,c=0) 是否作为基准（原逻辑: A3 bold）
    const baseA3 = worksheet.getRow(3).getCell(1); // ExcelJS: getRow(3) is row 3 (1-based)
    const baseIsBold = isTextBoldLikeMarker(baseA3.value);

    for (let r = 1; r <= rowCount; r++) {
        const cell = worksheet.getRow(r).getCell(1);
        const v = cell && cell.value !== undefined && cell.value !== null ? String(cell.value).trim() : '';
        if (!v) continue;
        if (baseIsBold) {
            if (!isTextBoldLikeMarker(v)) continue;
            doctors.push(new Doctor(cell, r - 1));
            continue;
        }
        const headerKeywords = ['备注', '总计', '日期', '姓名', '排班', '时间', '合计'];
        if (headerKeywords.some(k => v.includes(k))) continue;
        if (v.length > 8) continue;
        if (/[A-Za-z0-9]/.test(v)) continue;

        doctors.push(new Doctor(cell, r - 1));
    }

    return doctors.filter(d => d.section !== '错误');
}

function lookforExcelJS(worksheet, name, col = 1) {
    if (!worksheet || !name) return null;
    const rowCount = worksheet.rowCount || worksheet.actualRowCount || 0;

    const matches = [];
    for (let r = 2; r <= rowCount; r++) {
        const cell = worksheet.getRow(r).getCell(col + 1); 
        const v = (cell && cell.value !== undefined && cell.value !== null) ? String(cell.value).trim() : '';
        if (!v) continue;
        if (v.length > 18) continue;
        if (v.includes('皮') || v.length > 10) continue;
        if (v.includes(name)) matches.push({ r: r - 1, c: col, cell, addr: rcToA1(r - 1, col) });
    }
    if (matches.length === 1) return matches[0].cell;
    if (matches.length > 1) {
        console.warn(`lookfor: 找到多个匹配 ${name} -> ${matches.length}`);
        return matches[0];
    }
    return null;
}

const getCellSafeValue = (cellObj) => {//获取cell的值（安全的）
                
            // 1. 先判断单元格是否存在（避免 cellObj 为 null/undefined）
            if (!cellObj || cellObj.value === undefined) return null;
            const value = cellObj.value;

            // 🌟 新增：优先处理「富文本格式」（核心修复）
            if (value?.richText && Array.isArray(value.richText)) {
                // 遍历富文本数组，提取每段的 text 并拼接（忽略格式信息）
                return value.richText.map(segment => segment.text || '').join('');
            }

            // 2. 处理对象类型（排除 null，避免 JS 历史 bug）
            if (typeof value === 'object' && value !== null) {
                // 处理日期对象（转可读格式）
                if (value instanceof Date) {
                return value.toLocaleDateString(); // 如 "2025/12/01"，可按需调整
                }
                // 处理 Excel 公式对象（可选：优先取计算结果，无结果则取公式）
                if (value.formula) {
                return value.result || value.formula;
                }
                // 其他普通对象/数组（转 JSON 字符串，保留结构）
                return JSON.stringify(value);
            }

            // 3. 基础类型（字符串、数字、布尔）：直接返回（保持原类型）
            return value;
};

function changeSheetS_ExcelJS(flag) {//核心函数，对比与修改sheet。
    masterSheet=workbook.worksheets[0];
    const diffs = [];
    let modifiedCount = 0;  //修改计数
    matched.forEach(doc => {//对每个匹配成功的医生进行处理。
        subSheet=doc.cell_s.worksheet;
        const masterInfo = doc.cell_m;
        const subNameCol = doc.cell_s.col; // 0-based
        const masterNameCol = masterInfo.col;
        
        // 处理合并单元格同步
        if (flag === 1) {// sub -> master 合并复制。
            syncMergesExcelJS(subSheet, masterSheet, doc.row, masterInfo.row, subNameCol, masterNameCol);
        } else if (flag === 2) {// master -> sub合并复制。
            syncMergesExcelJS(masterSheet, subSheet, masterInfo.row, doc.row, masterNameCol, subNameCol);
        }

        for (let day = 1; day <= 14; day++) {//合并复制
            const subC = subNameCol + day;
            const masterC = masterNameCol + day;

            //获取主、分表班次单元格对象
            
            const subCellObj = subSheet.getRow(doc.cell_s.row).getCell(subC);
            const masterCellObj = masterSheet.getRow(doc.cell_m.row).getCell(masterC);
            
            // 调用函数获取cell的值（安全的获取）
            
            const subVal = getCellSafeValue(subCellObj);
            const masterVal = getCellSafeValue(masterCellObj);
            //const subVal = subCellObj.value
            //const masterVal = masterCellObj.value
            
            if (flag === 0) {
                // compare - 清洗空白并比较（case-insensitive）
                const vs = (subVal === null || subVal === undefined) ? '' : String(subVal).trim().replace(/[^\u4e00-\u9fa5]/g, '');
                const vm = (masterVal === null || masterVal === undefined) ? '' : String(masterVal).trim().replace(/[^\u4e00-\u9fa5]/g, '');
                if (vs !== vm) {
                    diffs.push({ name: doc, day, m: vm, s: vs ,cel:doc.cell_s});
                    //console.log(`发现差异：姓名<${doc.name}> <${getstart(day)}> 总表<${vm}> 分表<${vs}> 在表<${subSheet.name}>的单元格地址为 <${rcToA1(doc.row, subC)}>`);
                }
                else {
                    //console.log(`匹配成功：姓名<${doc.name}> <${getstart(day)}> 总表<${vm}> 分表<${vs}> 在表<${subSheet.name}>的单元格地址为 <${rcToA1(doc.row, subC)}>`);
                }
            } else {
                // 修改
                let srcCell = (flag === 1) ? subCellObj : masterCellObj;
                let tgtCell = (flag === 1) ? masterCellObj : subCellObj;

                // 如果源为合并区域的非主单元格，寻找主单元格
                const srcSheet = (flag === 1) ? subSheet : masterSheet;
                const rIndex = (flag === 1) ? doc.row : masterInfo.r;
                const cIndex = (flag === 1) ? subC : masterC;
                const srcMergeState = getMergeState(srcSheet, rIndex, cIndex);
                if (srcMergeState === 2) {
                    // 找到合并区间并使用主单元格
                    const ranges = getWorksheetMergeRanges(srcSheet);
                    for (const range of ranges) {
                        if (isCellInRange(rIndex, cIndex, range)) {
                            const mainRC = decodeRange(range).s; // 主单元格坐标
                            srcCell = srcSheet.getRow(mainRC.r + 1).getCell(mainRC.c + 1);
                            break;
                        }
                    }
                }

                // 执行复制（值 + 样式）
                copyCellValueAndStyleExcelJS(tgtCell, srcCell);

                modifiedCount++;
            }
        }
    });
    return { diffs, modifiedCount, matchedCount: matched.length };//diffs:差异列表，modifiedCount:修改计数，matchedCount:匹配医生计数。
}

function delflagExcelJS(ws){//？？删除斜杠，复制样式拆分AM PM。
    if (!ws) return;
    const rowCount = ws.rowCount || ws.actualRowCount || 0;
    // 目标列 2..15 (0-based)
    for (let c = 2; c <= 15; c++) {
        for (let r = 2; r <= rowCount; r++) { // 从第2行开始（1-based -> r=2）
            const cell = ws.getRow(r).getCell(c + 1);
            if (!cell || typeof cell.value !== 'string') continue;
            if (!cell.value.includes('/')) continue;

            const state = getMergeState(ws, r - 1, c);
            const parts = String(cell.value).split('/');
            const am = parts[0] || '';
            const pm = parts[1] || '';

            if (state === 0) {
                // 直接替换当前单元格为去斜杠的值
                cell.value = String(cell.value).replace('/', '');
                continue;
            }

            // 若为合并，unmerge 整行的合并（针对该行）
            unmergeRowExcelJS(ws, r - 1);

            // 写入 AM 到 c, PM 到 c+1 （注意创建单元格）
            const rowObj = ws.getRow(r);
            const addrAm = rowObj.getCell(c + 1);
            const addrPm = rowObj.getCell(c + 2);

            // 复制样式
            addrAm.value = typeof am === 'object' ? JSON.stringify(am) : am;
            if (cell.font) addrAm.font = deepClone(cell.font);
            if (cell.alignment) addrAm.alignment = deepClone(cell.alignment);
            if (cell.fill) addrAm.fill = deepClone(cell.fill);

            addrPm.value = typeof pm === 'object' ? JSON.stringify(pm) : pm;
            if (cell.font) addrPm.font = deepClone(cell.font);
            if (cell.alignment) addrPm.alignment = deepClone(cell.alignment);
            if (cell.fill) addrPm.fill = deepClone(cell.fill);
        }
    }
}

function statisticExcelJS(masterSheet) {//统计目标sheet的主专，返回统计列表。
    if (!masterSheet) return {};
    //showMsg('正在统计，稍后。。。', 'success');
    delflagExcelJS(masterSheet);
    const rowCount = masterSheet.rowCount || masterSheet.actualRowCount || 0;
    const result = {};

    const include = ['主', '专', '甲病', '黄褐斑', '白癜风', '痤疮'];
    const exclude = ['激', '脱', '性', '靶', '注射', '美容', '带疱'];

    for (let c = 2; c <= 15; c++) {
        const arr = [];
        for (let r = 2; r <= rowCount; r++) {
            const cell = masterSheet.getRow(r).getCell(c + 1);
            if (!cell || !cell.value) continue;
            // 如果是合并的非主单元格，跳过
            if (getMergeState(masterSheet, r - 1, c) === 2) continue;
            const val = String(cell.value).trim();
            if (val.length > 10) continue;
            if (exclude.some(k => val.includes(k))) continue;
            if (include.some(k => val.includes(k)) && !val.includes('激')) {
                arr.push(`${r}-${val}`);
            }
        }
        result[c - 1] = arr;
    }

    return result;
}

function getstart(num) {//获取星期几与上下午，根据0-14数字。
  // 1. 参数校验：确保是1-14之间的有效数字（排除非数字、NaN、超出区间值）
  const isQualified = 
    typeof num === 'number' && 
    !isNaN(num) && 
    num >= 1 && 
    num <= 14;

  if (!isQualified) {
    return [];
  }

  // 2. 核心计算：被除数+1 → 得到周数（1-7）→ 判断am/pm
  const adjustedDividend = num + 1; // 被除数先+1
  const weekNum = Math.floor(adjustedDividend / 2); // 周数（1-7，无需额外加减）
  const period = adjustedDividend % 2 === 0 ? 'Am' : 'Pm'; // 上下午标识

  // 3. 拼接目标格式：周X_Xm（例：周1_am、周7_pm）
  return `周${weekNum}_${period}`;
}

function runCompareExcelJS() {//对比doctor对。
    
    let totalDiffs = 0;
    let html = '<thead><tr><th>姓名</th><th>日期</th><th>总表</th><th>分表</th><th>科室</th></tr></thead><tbody>';

    const worksheets = workbook.worksheets;
    const masterSheet = worksheets[0];
    showMsg('正在对比，稍后。。。', 'success');
    for (let i = 1; i < worksheets.length; i++) {
        const subSheet = worksheets[i];
        console.log(`正在对比：表<${subSheet.name}> 与表 <${masterSheet.name}>`);
        const res = changeSheetS_ExcelJS(0);
        res.diffs.forEach(d => {
            html += `<tr><td>${d.name.name}</td><td>${getstart(d.day)}</td><td>${d.m}</td><td>${d.s}</td><td>${d.cel.worksheet.name} _ ${d.cel.address}</td></tr>`;
        });
        totalDiffs += res.diffs.length;
    }
    html += '</tbody>';
    if (els && els.table) els.table.innerHTML = html;
    if (totalDiffs === 0) showMsg('完美！未发现任何差异', 'success');
    else showMsg(`发现 ${totalDiffs} 处不一致`, 'error');
}

function runModifyExcelJS(flag) {//使用doctor数据修改表。
    let totalModified = 0;
    const worksheets = workbook.worksheets;
    if (!worksheets || worksheets.length === 0) return showMsg('工作簿没有任何工作表', 'error');
    showMsg('正在修改，稍后。。。', 'success');
    for (let i = 1; i < worksheets.length; i++) {
        const subSheet = worksheets[i];
        const res = changeSheetS_ExcelJS(flag);
        totalModified += res.modifiedCount || 0;
    }
    const type = flag === 1 ? '总表' : '分表';
    showMsg(`${type}修改完成！共修改 ${totalModified} 个单元格，样式与合并已同步，请下载保存。`, 'success');
    if (els && els.btns && els.btns.download) els.btns.download.style.display = 'block';
}

function runStatisticExcelJS() {//调用统计->整合输出。
    const stats = statisticExcelJS(workbook.worksheets[0]);
    let html = '<thead><tr><th>日期</th><th>人数</th><th>详情</th></tr></thead><tbody>';
    for (const key in stats) {
        const arr = stats[key];
        const count = arr.length;
        const style = count > 16 ? 'style="background:#ffebee; color:#c62828; font-weight:bold;"' : '';
        html += `<tr ${style}><td>${getstart(Number(key))}</td><td>${count}</td><td style="text-align:left">${arr.join(', ')}</td></tr>`;
    }
    html += '</tbody>';
    if (els && els.table) els.table.innerHTML = html;
    showMsg('统计完成 (红色行表示超过16人)', 'success');
}

function init(){    //初始化匹配医生列表
    workbook.worksheets.forEach((sheet, index) => {
      if (index === 0) return; // 跳过第一个Sheet（索引0）
    doctors = getDoctorsExcelJS(sheet);    //获取医生列表
    })
    // 关键排查：打印 doctors 的值和类型
    doctors.forEach(doc => {    //匹配医生到总表
        const found = lookforExcelJS(workbook.worksheets[0], doc.name, 1);
        if (!found) {
            console.warn(`<${doc.section}>科室内的<${doc.name}> -- 不在总表内`);
            return;
        }
        doc.cell_m = found;
        matched.push(doc);  //记录匹配成功的医生
    });
    console.log(`共匹配成功 ${matched.length} 位医生`);
}
