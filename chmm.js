
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

function isCellMasterInRange(r, c, rangeStr) {   //检查单元格是否是指定范围的主单元格。
    const range = decodeRange(rangeStr);
    return (r === range.s.r && c === range.s.c);
}

function getMR(ws) {//获取sheet的合并范围。
    // 获取工作表中的所有合并范围，返回数组：rangeStr，如 ["A1:C1", "E2:E3", ...]
    if (!ws) {
        console.warn('getMR：工作表 ws 不存在');
        return [];
    }
    try {
        const mergedRanges = ws.model.merges;
        // 将 MergeRange 对象转为范围字符串（如 MergeRange → "A1:C1"）
        return mergedRanges;
        //return mergedRanges.map(range => range.address);
    } catch (e) {
        // 🌟 修正：输出具体错误日志，方便排查
        console.error('getMR：获取合并范围失败', e.message);
        return [];
    }
}

function isCellInRange(cell, rangeS) {   //检查单元格是否在指定范围内。
    const range = decodeRange(rangeS);
    r = decodeRange(cell.address).s.r;
    c = decodeRange(cell.address).s.c;
    return (r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c);
}

function getMergeState(cell) {   //获取单元格的合并状态。
    const ranges = getMR(cell.worksheet);
    r = decodeRange(cell.address).s.r;
    c = decodeRange(cell.address).s.c;
    for (const range of ranges) {
        if (isCellInRange(cell, range)) {
            //0:非合并单元格 1: 主单元格，2：非主单元格
            const ran = decodeRange(range);
            if (r === ran.s.r && c === ran.s.c) return 1;
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

    const ranges = getMR(sourceSheet);//获取工作表 合并范围
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



function unmergeRowExcelJS(ws, targetRow) {// 在目标工作表上删除包含目标行的所有合并。
    if (!ws) return;
    const rowNumber = targetRow + 1;
    const ranges = getMR(ws);
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

function changeSheetS_ExcelJS(flag) {//核心函数，对比与修改sheet。
    masterSheet=workbook.worksheets[0];
    const diffs = [];
    let modifiedCount = 0;  //修改计数
    diffs.forEach(doc => {//对每个匹配成功的医生进行处理。
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
                    const ranges = getMR(srcSheet);
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
    );
    return { diffs, modifiedCount, matchedCount: matched.length };//diffs:差异列表，modifiedCount:修改计数，matchedCount:匹配医生计数。
}

/**
 * 在 ExcelJS（浏览器版）中复制一个矩形范围到指定左上角单元格。
 *
 * @param {Worksheet} ws        - 工作表对象
 * @param {string} srcRange     - 源范围，如 "B2:D6"
 * @param {string} dstStart     - 目标左上角，如 "F10"
 */
function copyRange() {
    const ExcelJS = window.ExcelJS;
    ws = workbook.worksheets[0]
    //const src = ExcelJS.Workbook.xlsx.decodeRange(srcRange); // {top, left, bottom, right}
    //const dst = ExcelJS.Workbook.xlsx.decodeAddress(dstStart); // {row, col}
    const src = {top:2,left:'B',bottom:8,right:'P'}
    const dst = {row:2, col:'T'}
    
    const rowOffset = dst.row - src.top;
    const colOffset = dst.col - src.left;

    /* 1. 复制单元格内容与样式
    for (let r = src.top; r <= src.bottom; r++) {
        for (let c = src.left; c <= src.right; c++) {

            const srcCell = ws.getCell(r, c);
            const dstCell = ws.getCell(r + rowOffset, c + colOffset);

            // 拷贝值（支持公式）
            dstCell.value = srcCell.value;

            // 拷贝所有样式
            dstCell.style = JSON.parse(JSON.stringify(srcCell.style));
        }
    }*/

    // 2. 复制合并单元格
    ws.model.merges.forEach(m => {
        console.log('1111')
        if (
            m.top    >= src.top    && m.bottom <= src.bottom &&
            m.left   >= src.left   && m.right  <= src.right
        ) {
            const newMerge = {
                top:    m.top    + rowOffset,
                bottom: m.bottom + rowOffset,
                left:   m.left   + colOffset,
                right:  m.right  + colOffset
            };
            ws.mergeCells(
                newMerge.top, newMerge.left,
                newMerge.bottom, newMerge.right
            );
        }
    });
}