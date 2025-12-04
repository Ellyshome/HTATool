const matched = new Set();
const doctors = new Set();    //获取医生列表
let diffs = new Set();
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
        this.name = this.extractName(this.cellString);
        this.cell_m = null; //cell_m单元格总表
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
    const Keywords = ['备注', '总计', '日期', '姓名', '排班', '时间', '合计','专家','黑专','普门','皮'];
    if (!val || Keywords.some(k => val.includes(k))) {
        console.warn(`在表<${sheet.name}>发现疑似非法姓名： <${val}> , 丢弃.原因:包含关键词`);
        return false;
    }
    if (val.length > 15) {
        console.warn(`在表<${sheet.name}>发现疑似非法姓名： <${val}> , 丢弃.原因:长度超过15`);
        return false;
    }
    return true;
}

function getDoctorsExcelJS(worksheet) {//在指定sheet中，找到并压入Doctor。
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
        if (IsName(val,worksheet)) doctors.add(new Doctor(cell));
    }
}

function lookforExcelJS(worksheet, name, col = 1) {//从总表中找到对应的行。
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
            const subVal = getCellSafeValue(subCellObj);
            const masterVal = getCellSafeValue(masterCellObj);
            // compare - 清洗空白并比较（case-insensitive）
            const vs = (subVal === null || subVal === undefined) ? '' : String(subVal).trim().replace(/[^\u4e00-\u9fa5]/g, '');
            const vm = (masterVal === null || masterVal === undefined) ? '' : String(masterVal).trim().replace(/[^\u4e00-\u9fa5]/g, '');
            if (vs !== vm)  {
                doc.dif.push({ d: day, m: vm, s: vs});
                diffs.add(doc);
            }
        }
    }   );
    els.btns.download.style.display = 'block';
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
    let html = '<thead><tr><th>姓名</th><th>总表行号</th><th>日期</th><th>总表</th><th>分表</th><th>分表位置</th></tr></thead><tbody>';
    showMsg('正在对比，稍后。。。', 'success');
    Compare();
    diffs.forEach(d => {
        d.dif.forEach(diff => {
            html += `<tr><td>${d.name}</td><td>${d.cell_m.row}</td><td>${getstart(diff.d)}</td><td>${diff.m}</td><td>${diff.s}</td><td>${d.cell_s.worksheet.name} _ ${d.cell_s.address}</td></tr>`;
        });
        //html += `<tr><td>${d.name}</td><th>${d.cell_m.row}</th><td>${getstart(d.dif['d'])}</td><td>${d.dif['m']}</td><td>${d.dif['s']}</td><td>${d.cell_s.worksheet} _ ${d.cell_s.address}</td></tr>`;
    });
    console.log(`共发现 ${diffs.size} 人不一致`);
    html += '</tbody>';
    if (els && els.table) els.table.innerHTML = html;
    if (diffs.size === 0) showMsg('完美！未发现任何差异', 'success');
    else showMsg(`发现 ${diffs.size} 人不一致`, 'error');
}

function runModifyExcelJS(flag) {//改总\分表。
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

function init(){    //初始化匹配医生列表。
    workbook.worksheets.forEach((sheet, index) => {
      if (index === 0) return; // 跳过第一个Sheet（索引0）
    getDoctorsExcelJS(sheet);    //获取医生列表
    })
    // 关键排查：打印 doctors 的值和类型
    doctors.forEach(doc => {    //匹配医生到总表
        const found = lookforExcelJS(workbook.worksheets[0], doc.name, 1);
        if (!found) {
            console.warn(`<${doc.section}>科室内的<${doc.name}> -- 不在总表内`);
            return;
        }
        doc.cell_m = found;
        matched.add(doc);  //记录匹配成功的医生
    });
    
    console.log(`共匹配成功 ${matched.size} 位医生`);
}