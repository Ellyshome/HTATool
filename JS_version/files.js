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

/* ------文件管理-------------------
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
        else if (mode === 'modifyMaster') runModifyExcelJS(0);
        else if (mode === 'modifySub') runModifyExcelJS(1);
        else if (mode === 'statistic') runStatisticExcelJS();
    } catch (e) {
        console.error('处理失败:', e);
        showMsg(`运行出错: ${e.message}`, 'error');
    }
}

async function handleFileSelectExcelJS(e) {//文件加载/拖放/下载
    const file = (e.target && e.target.files && e.target.files[0]) || (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]);
    if (!file || !file.name || !file.name.endsWith('.xlsx')) return showMsg('请使用 .xlsx 文件', 'error');

    showMsg(`正在加载文件: ${file.name}`, 'loading');

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
        nd = init(); // 初始化医生匹配列表,获取未匹配医生列表
        const sectionCount = {};
        for (const doctor of matched) {
        const section = doctor.section;
        sectionCount[section] = (sectionCount[section] || 0) + 1;
        }
        // 拼接科室统计文本（适配任意科室）
        const sectionText = Object.entries(sectionCount)
        .map(([section, count]) => `表< ${section} >${count}人 `);
        showMsg(`文件加载完成，共匹配成功 ${matched.size} 位医生，其中${sectionText}注意核对！` ,'success');
        let html = '<thead><tr><th>姓名</th><th>科室</th><th>异常</th></thead><tbody>';
        nd.forEach(item => {
            html += `<tr><td>${item.name}</td><td>${item.section}_${item.row}</td><td>${item.reason}</td></tr>`;
        });

        html += '</tbody>';
    if (els && els.table) els.table.innerHTML = html;
        //runCompareExcelJS(); // 默认加载后进行对比  

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

function handleDropExcelJS(e) {//拖放文件处理
    e.preventDefault();
    if (els && els.dropArea) els.dropArea.classList.remove('dragover');
    const file = (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]);
    if (!file) return showMsg('未检测到文件', 'error');
    const fakeEvent = { target: { files: [file] } };
    handleFileSelectExcelJS(fakeEvent);
}

async function downloadResultExcelJS() {//下载处理结果。
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