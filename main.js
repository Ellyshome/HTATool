

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
        const sheets = workbook.worksheets;
        if (!sheets || sheets.length === 0) return showMsg('工作簿没有任何工作表', 'error');
        const masterSheet = sheets[0];

        if (mode === 'compare') runCompareExcelJS(sheets, masterSheet);
        else if (mode === 'modifyMaster') runModifyExcelJS(sheets, masterSheet, 1);
        else if (mode === 'modifySub') runModifyExcelJS(sheets, masterSheet, 2);
        else if (mode === 'statistic') runStatisticExcelJS(masterSheet);
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
        showMsg('文件加载成功，请选择功能', 'success');

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

