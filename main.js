let workbook = null;
let fileName = '';
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
els.fileInput.addEventListener('change', handleFileSelect);
els.dropArea.addEventListener('dragover', (e) => { e.preventDefault(); els.dropArea.classList.add('dragover'); });
els.dropArea.addEventListener('drop', handleDrop);

els.btns.compare.addEventListener('click', () => processWorkbook('compare'));
els.btns.modMaster.addEventListener('click', () => processWorkbook('modifyMaster'));
els.btns.modSub.addEventListener('click', () => processWorkbook('modifySub'));
els.btns.stat.addEventListener('click', () => processWorkbook('statistic'));
els.btns.download.addEventListener('click', downloadResult);



