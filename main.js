

const els = {
    uploadBtn: document.getElementById('uploadBtn'),
    fileInput: document.getElementById('fileInput'),
    fileInfo: document.getElementById('fileInfo'),
    dropArea: document.getElementById('drop-area'),
    btns: {
        compare: document.getElementById('compareBtn'),
        modMaster: document.getElementById('modifyMasterBtn'),
        modSub: document.getElementById('modifySubBtn'),
        stat: document.getElementById('statisticBtn'),
        download: document.getElementById('downloadBtn')
    },
    resultSection: document.getElementById('resultSection'),
    table: document.getElementById('resultTable'),
    msg: document.getElementById('messageArea')
};

els.uploadBtn.addEventListener('click', () => els.fileInput.click());
els.fileInput.addEventListener('change', handleFileSelectExcelJS);
els.dropArea.addEventListener('dragover', (e) => { e.preventDefault(); els.dropArea.classList.add('dragover'); });
els.dropArea.addEventListener('drop', handleDropExcelJS);

els.btns.compare.addEventListener('click', () => processWorkbook('compare'));
els.btns.modMaster.addEventListener('click', () => processWorkbook('modifyMaster'));
els.btns.modSub.addEventListener('click', () => processWorkbook('modifySub'));
els.btns.stat.addEventListener('click', () => processWorkbook('statistic'));
els.btns.download.addEventListener('click', downloadResultExcelJS);

/* -------------------------
   processWorkbook 入口（ExcelJS）
   - mode: 'compare' | 'modifyMaster' | 'modifySub' | 'statistic'
   ------------------------- */
function processWorkbook(mode) {
    if (!workbook) return showMsg('请先加载文件', 'error');
    showMsg('正在处理...', 'loading');
    if (els && els.table) els.table.innerHTML = '';
    if (els && els.btns && els.btns.download) els.btns.download.style.display = 'none';

    try {
        if (mode === 'compare') runCompareExcelJS();
        else if (mode === 'modifyMaster') runModifyExcelJS(1);
        else if (mode === 'modifySub') runModifyExcelJS(2);
        else if (mode === 'statistic') runStatisticExcelJS();
    } catch (e) {
        console.error('处理失败:', e);
        showMsg(`运行出错: ${e.message}`, 'error');
    }
}


/* -------------------------
   文件加载/拖放/下载（ExcelJS 版）
   ------------------------- */
async function handleFileSelectExcelJS(e) {
    const file = (e.target && e.target.files && e.target.files[0]) || (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]);
    if (!file || !file.name || !file.name.endsWith('.xlsx')) return showMsg('请使用 .xlsx 文件', 'error');

    showMsg('正在处理，请稍候...', 'loading');

    try {
        const arrayBuffer = await file.arrayBuffer();
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(arrayBuffer);
        workbook = wb;
        fileName = file.name;

        if (els && els.fileInfo) {
            els.fileInfo.textContent = `当前文件: ${file.name}`;
            els.fileInfo.style.color = 'green';
        }
        init(); // 初始化医生匹配列表
        runCompareExcelJS(); // 默认加载后进行对比  

        // 启用按钮（若页面有）
        try { Object.values(els.btns).forEach(b => b.disabled = false); } catch (e) {}
        if (els && els.btns && els.btns.download) els.btns.download.style.display = 'none';
    } catch (err) {
        console.error('文件处理失败:', err);
        showMsg(`文件处理失败: ${err.message}`, 'error');
        workbook = null;
        fileName = '';
    }
}

async function syncCellMerges(doctor,flag = 0) {//对比并同步单元格合并情况：总表 第10行（B-Z列）→ 分表 第20行（A-Y列）
    const sheet_m=doctor.cell_m.worksheet;
    const sheet_s=doctor.cell_s.worksheet;
  try {
    // 1. 定义核心配置（行号、列范围映射）
    const config = {
      baseRow: doctor.cell_m.row,          // 总表 基准行：第10行
      targetRow: doctor.cell_s.row,        // 分表 目标行：第20行
      baseCols: { start: 3, end: 16 }, // 总表 列范围：B(3)~Z(16)
      targetCols: { start: 2, end: 15 } // 分表 列范围：A(2)~Y(15)
    };

    // 2. 提取总表 列）的合并单元格信息（基准合并规则）
    const baseMerges = [];
    sheet_m.mergedCells.forEach(merge => {
      // 筛选：仅保留目标行单行合并（合并仅在同一行内）
      if (
        merge.start.row === config.baseRow && 
        merge.end.row === config.baseRow && 
        merge.start.column >= config.baseCols.start && 
        merge.end.column <= config.baseCols.end
      ) {
        // 记录总表 的合并列区间（绝对列号）
        baseMerges.push({
          startCol: merge.start.column,
          endCol: merge.end.column
        });
      }
    });

    // 3. 提取分表 第20行（A-Y列）的现有合并单元格信息
    const targetMerges = [];
    sheet_s.mergedCells.forEach(merge => {
      if (
        merge.start.row === config.targetRow && 
        merge.end.row === config.targetRow && 
        merge.start.column >= config.targetCols.start && 
        merge.end.column <= config.targetCols.end
      ) {
        targetMerges.push({
          startCol: merge.start.column,
          endCol: merge.end.column
        });
      }
    });

    // 4. 对比合并规则是否一致（一致则直接返回，无需修改）
    const isMergesEqual = JSON.stringify(baseMerges.map(m => `${m.startCol}-${m.endCol}`)) ===
        JSON.stringify(targetMerges.map(m => `${m.startCol}-${m.endCol}`));
    if (isMergesEqual) {
        console.log(`ℹ️ 合并单元格总表：${baseRow}行，分表${targetRow}}行：校验一致，无需修改`);
        return true;
    }
    else console.log(`⚠️ 合并单元格总表：${baseRow}行，分表${targetRow}行：校验不一致，需同步修改`);
/*
    // 5. 不一致：先清除分表 目标行（A-Y列）的所有现有合并
    targetMerges.forEach(merge => {
      sheet_s.unMergeCells(
        config.targetRow, merge.startCol,
        config.targetRow, merge.endCol
      );
    });
    console.log(`ℹ️  已清除分表 第20行（A-Y列）原有合并`);

    // 6. 按总表 的合并规则，在分表 目标行创建对应合并（列映射：总表 列 -1 = 分表 列）
    baseMerges.forEach(baseMerge => {
      // 总表 列 → 分表 列：B(2)→A(1)、Z(26)→Y(25)，即 baseCol - 1
      const targetStartCol = baseMerge.startCol - 1;
      const targetEndCol = baseMerge.endCol - 1;

      // 在分表 第20行，创建对应合并
      sheet_s.mergeCells(
        config.targetRow, targetStartCol,
        config.targetRow, targetEndCol
      );
    });

    console.log(`✅ 合并同步完成：分表 第20行（A-Y列）已与总表 第10行（B-Z列）保持一致`);
    return true;*/
  } catch (err) {
    console.error(`❌ 同步合并失败：`, err.message);
    return false;
  }
}

function handleDropExcelJS(e) {
    e.preventDefault();
    if (els && els.dropArea) els.dropArea.classList.remove('dragover');
    const file = (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]);
    if (!file) return showMsg('未检测到文件', 'error');
    const fakeEvent = { target: { files: [file] } };
    handleFileSelectExcelJS(fakeEvent);
}

async function downloadResultExcelJS() {
    if (!workbook) return showMsg('没有可保存的工作簿', 'error');
    if (!fileName) return showMsg('文件名未定义', 'error');

    try {
        const buf = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const ts = new Date().toISOString().slice(0,10).replace(/-/g,'');
        const newName = fileName.replace('.xlsx', `_processed_${ts}.xlsx`);

        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = newName;
        a.click();

        showMsg(`文件已保存为: ${newName}`, 'success');
    } catch (e) {
        console.error('保存时出错', e);
        showMsg('保存失败（请在控制台查看错误）', 'error');
    }
}

