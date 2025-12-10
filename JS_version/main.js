//------全局变量----------------------------------------------------
const matched = new Set();  //匹配成功医生列表
const doctors = new Set();   //获医生列表
let diffs = new Set();  //记录有差异的医生列表
//------单元格合并----------------------------------------------------------------------

function isCellInRange(cell, rangeS) {   //检查单元格是否在指定范围内。
    const range = decodeRange(rangeS);
    r = decodeRange(cell.address).s.r;
    c = decodeRange(cell.address).s.c;
    return (r >= range.s.r && r <= range.e.r && c >= range.s.c && c <= range.e.c);
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

function getMergeState(cell) {   //获取单元格的合并状态。
    //0:非合并单元格 1: 主单元格，2：非主单元格
    const ranges = getMR(cell.worksheet);
    r = decodeRange(cell.address).s.r;
    c = decodeRange(cell.address).s.c;
    for (const range of ranges) {
        if (isCellInRange(cell, range)) {
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

function deepcopy(sourceCell, targetCell) {//单元格复制，核心函数。
    if (!targetCell) return;
    // 🌟 第一步：深度复制「值」（按类型处理，重点支持富文本）
    if (!sourceCell) {
        // 源单元格不存在：清空目标单元格的值和所有样式
        targetCell.value = null;
        ['font', 'fill', 'border', 'alignment', 'numFmt'].forEach(key => delete targetCell[key]);
        return;
    }
    if(sourceCell.value === targetCell.value) return;
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
    return true;
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

function changeSheetS(flag) {//核心函数，对比与修改sheet。
    //flag = 0 主标覆盖分表；flag = 1 分表覆盖主标
    count=0;
    bug=[];
    diffs.forEach(doc =>{
        
       doc.dif.forEach(dif => {
            const targetCell = flag ? dif.subcell : dif.mastercell;
            const sourceCell = flag ? dif.mastercell : dif.subcell;
            //console.log(`${flag?'分表':'总表'}<${doc.name}>条目<${targetCell.address}>修改中...`);
            try{
                //compareMerge(dif,flag) ; //先处理合并单元格
                compareMerge(sourceCell, targetCell,dif) ; //先处理合并单元格
                deepcopy(sourceCell, targetCell); //再复制值与样式
                count++;
            }catch (e) {
                bug.push([`修改<${targetCell.worksheet.name}>的<${doc.name}>条目 ${targetCell.address} 时遇到问题:${e.message}。`]); 
                console.error(`修改<${targetCell.worksheet.name}>的<${doc.name}>条目<${targetCell.address}>时遇到问题`, e.message);
        }});
    })
    return [count,bug];
}
function compareMerge(sou_cell, tar_cell,dif){//cell合并状态，根据diffs中doctor的dif列表。
//function compareMerge(dif,flag){//cell合并状态，根据diffs中doctor的dif列表。
    //const [tar_cell, sou_cell] = flag ? [dif['subcell'], dif['mastercell']] : [dif['mastercell'] , dif['subcell']];
    tar_sheet = tar_cell.worksheet;
    if (getMergeState(sou_cell)===getMergeState(tar_cell)) return;
    if (getMergeState(sou_cell)===0) {tar_sheet.unMergeCells(tar_cell.address);return;}//源单元格被标记为 分散 状态
    row_se = tar_cell.row;
    const [col_s,col_e] = dif['day'] % 2 === 0 ? [tar_cell.col+1,tar_cell.col] : [tar_cell.col,tar_cell.col-1];
    tar_sheet.mergeCells(row_se,col_s,row_se,col_e);  
    }

//-----查与改-----------------------------------------------------------------------
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

function showMsg(msg, type = 'info') {//要显示的消息。
    if (!els || !els.msg) {
        console.log(`[${type}] ${msg}`);
        return;
    }
    els.msg.innerHTML = `<div class="${type}">${msg}</div>`;
    if (els.resultSection) els.resultSection.classList.add('active');
}

class Doctor {//医生类
    constructor(cell) {
        this.row = cell.row;

        //cellString单元格分表
        this.cellString = cell && cell.value !== undefined && cell.value !== null ? String(cell.value).trim() : '';
        this.cell_s=cell;
        this.cell_m = null; //cell_m单元格总表
        this.name = this.extractName(this.cellString);
        this.section = cell.worksheet.name;
        this.dif = [];
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

function IsName(val,sheet) {// 基于既定规则，判断文本是人名
    //排除为姓名的规则
    const Keywords = ['备注', '总计', '日期', '姓名', '排班', '时间', '合计','专家','黑专','普门','皮','说明','补充'];
    if (!val || Keywords.some(k => val.includes(k))) {
        //console.warn(`在表<${sheet.name}>发现疑似非法姓名： <${val}> , 丢弃.原因:包含关键词`);
        return false;
    }
    if (val.length > 15||val.length < 2) {
        console.warn(`在表<${sheet.name}>发现疑似非法姓名： <${val}> , 丢弃.原因:长度超过15`);
        return false;
    }
    return true;
}

function getDoctors(worksheet) {//在指定sheet中，找到并压入Doctor。
    if (!worksheet) {
    console.warn('获取医生列表失败：工作表不存在', 'error');
    return;
}
    const rowCount = worksheet.rowCount || worksheet.actualRowCount || 0;
    // 检查 A3 (r=2,c=0) 是否作为基准（原逻辑: A3 bold）
    //const baseA3 = worksheet.getRow(3).getCell(1); // ExcelJS: getRow(3) is row 3 (1-based)
    for (let r = 1; r <= rowCount; r++) {
        const cell = worksheet.getRow(r).getCell(1);
        const val = cell && cell.value !== undefined && cell.value !== null ? String(cell.value).trim() : '';
        if(!val) continue;
        if (IsName(val,worksheet)) doctors.add(new Doctor(cell));
    }
}

function lookfor(worksheet, name, col = 1) {   //从总表中找到对应的行。
    if (!worksheet || !name) return null;
    const rowCount = worksheet.rowCount || worksheet.actualRowCount || 0;

    const matches = [];
    for (let r = 2; r <= rowCount; r++) {
        const cell = worksheet.getRow(r).getCell(col + 1); 
        const v = (cell && getCellText(cell) !== undefined && cell.value !== null) ? String(cell.value).trim() : '';
        if (!v) continue;
        if (v.includes('皮') || v.length > 10) continue;
        if (v.includes(name)) matches.push(cell);
    }
    if (matches.length === 1) return matches[0];
    if (matches.length > 1) {
        console.warn(`lookfor: 找到多个匹配 ${name} -> ${matches.length}`);
        return matches[0];
    }
    return null;
}

function getCellText(cell) {//获取单元格文本内容（多种情况处理）。
    const v = cell.value;
    if (v == null) return "";

    // 情况 1：普通文本或数字、布尔值
    if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") {
        return String(v);
    }

    // 情况 2：富文本 { richText: [...] }
    if (v.richText) {
        return v.richText.map(part => part.text).join("");
    }

    // 情况 3：超链接 { text: "...", hyperlink:"..." }
    if (v.text) {
        return v.text;
    }

    // 情况 4：公式单元格 { formula: "...", result: ... }
    if (v.formula != null) {
        // 一般用于比对文本，应比对 result
        if (v.result != null) return String(v.result);
        return ""; // 没有 result 时返回空
    }

    // 情况 5：日期
    if (v instanceof Date) {
        return v.toISOString();
    }

    // 兜底
    return String(v);
}

function Compare(){//对比总表与分表医生班次。
    //masterSheet=workbook.worksheets[0];
    diffs.clear();
    matched.forEach(doc => {//对每个匹配成功的医生进行处理。
        subSheet=doc.cell_s.worksheet;
        const subNameCol = doc.cell_s.col; // 0-based
        const masterNameCol = doc.cell_m.col;
        doc.dif.length=0;
        for (let day = 1; day <= 14; day++) {//合并复制
            const subC = subNameCol + day;
            const masterC = masterNameCol + day;
            //获取主、分表班次单元格对象
            const subCellObj = subSheet.getRow(doc.cell_s.row).getCell(subC);
            const masterCellObj = workbook.worksheets[0].getRow(doc.cell_m.row).getCell(masterC);
            // 调用函数获取cell的值（安全的获取）
            const subVal = getCellText(subCellObj);
            const masterVal = getCellText(masterCellObj);
            // compare - 清洗空白并比较（case-insensitive）
            const vs = (subVal === null || subVal === undefined) ? '' : String(subVal).trim().replace(/[^\u4e00-\u9fa5]/g, '');
            const vm = (masterVal === null || masterVal === undefined) ? '' : String(masterVal).trim().replace(/[^\u4e00-\u9fa5]/g, '');
            if (vs !== vm)  {
                doc.dif.push({ d: day, mastercell: masterCellObj, subcell: subCellObj});
                diffs.add(doc);
            }
        }
    });
    els.btns.download.style.display = '';
}

function splitBySlash(str, num) {//按斜杠分割字符串并根据数字返回对应段落。
  // 容错：确保参数1为字符串类型
  const targetStr = String(str);
  
  // 1. 不包含 / 则原样返回
  if (!targetStr.includes('/')) {
    return targetStr;
  }

  // 2. 包含 / 则分割为两段（即使有多个 /，仅取前两段；末尾/分割后空字符串也保留）
  const [firstSegment, secondSegment = ''] = targetStr.split('/');

  // 3. 容错处理参数2：转为数字，非数字则按非偶数处理
  const targetNum = Number(num);
  const isEven = !isNaN(targetNum) && targetNum % 2 !== 0;

  // 4. 偶数返回前一段，非偶数返回后一段
  return isEven ? firstSegment : secondSegment;
}

function statisticExcelJS() {//统计sheet主专，返回结果列表。
    const masterSheet = workbook.worksheets[0];
    const rowCount = masterSheet.rowCount || masterSheet.actualRowCount || 0;
    const result = {};
    const include = ['主', '专', '甲病', '黄褐斑', '白癜风', '痤疮'];
    const exclude = ['激', '脱', '性', '靶', '注射', '美容', '带疱'];

    for (let col = 3; col <= 16; col++) {//从第二列开始
        
        const arr = [];
        for (let row = 2; row <= rowCount; row++) {
            const cell = masterSheet.getRow(row).getCell(col);
            if (!cell || !cell.value) continue;
            const value = splitBySlash(getCellText (cell),col);
            if (value.length > 15) continue;
            if (exclude.some(k => value.includes(k))) continue;
            if (include.some(k => value.includes(k)) && !value.includes('激')) {//激专不算
                arr.push(`${row}-${value}\t  `);
            }
        }
        result[col-2] = arr;
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

//-----DOM接口-----------------------------------------
function runCompareExcelJS() {//对比doctor对。
    showMsg('正在对比，稍后。。。', 'success');
    let html = '<thead><tr><th>姓名</th><th>日期</th><th>总表</th><th>分表</th><th>对应位置</th></tr></thead><tbody>';
    Compare();
    diffs.forEach(d => {
        d.dif.forEach(diff => {  
            html += `<tr><td>${d.name}</td><td>${getstart(diff.d)}</td><td>${getCellText(diff.mastercell)}</td><td>${getCellText(diff.subcell)}</td><td>总表${d.cell_m.row}行 : 分表 ${d.cell_s.worksheet.name}_${d.cell_s.row}行</td></tr>`;
        });
    });
    html += '</tbody>';
    if (els && els.table) els.table.innerHTML = html;
    if (diffs.size === 0) showMsg('完美！未发现任何差异', 'success');
    else showMsg(`发现 ${diffs.size} 人不一致`, 'error');
    els.btns.download.style.display = 'block';
}

function runModifyExcelJS(flag) {//改总\分表。
    //flag=0为改总表，flag=1为改分表
    showMsg(`正在修改，请稍后。。。`, 'success');
    setTimeout(() => { 
        const worksheets = workbook.worksheets;
        if (!worksheets || worksheets.length === 0) return showMsg('工作簿没有任何工作表', 'error');
        if (diffs.size === 0 ) Compare();
        const[count,bug] = changeSheetS(flag);
        const type = flag? '分表' : '总表';
        showMsg(`${type}修改完成！共修改${count}处，请下载保存。`, 'success');
        if(bug.length!==0){
            let html = `<thead><tr><th>异常条目数：${bug.length}注意手动处理</th></tr></thead><tbody>`;
            for (const key in bug) {
                html += `<tr><td>${bug[key]}</td></tr>`;
            }
            html += '</tbody>';
            if (els && els.table) els.table.innerHTML = html;
        }
        els.btns.download.style.display = 'block';
    }, 0);
}

function runStatisticExcelJS() {//调用统计->整合输出。
    const stats = statisticExcelJS();
    let html = '<thead><tr><th>日期</th><th>人数</th><th>详情（行号-门诊类型）</th></tr></thead><tbody>';
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

function init(){    //初始化匹配医生列表。
    let notinMsheet =[];
    workbook.worksheets.forEach((sheet, index) => {
    if (index === 0) return; // 跳过第一个Sheet（索引0）
    getDoctors(sheet);    //获取医生列表
    })
    // 关键排查：打印 doctors 的值和类型
    doctors.forEach(doc => {    //匹配医生到总表
        const found = lookfor(workbook.worksheets[0], doc.name, 1);
        if (!found) {
            notinMsheet.push({name:doc.name,section:doc.section,row:doc.cell_s.row,reason:'不在总表内'});
            return;
        }
        doc.cell_m = found;
        matched.add(doc);  //记录匹配成功的医生
    });
    return notinMsheet;
}