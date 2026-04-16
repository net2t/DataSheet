const STORAGE_KEY_DATA = 'sheetData_v2';
const STORAGE_KEY_HISTORY = 'sheetHistory_v1';

const STATUS_OPTIONS = [
    { value: 'STAGE 1', label: 'STAGE 1', icon: 'fa-flag' },
    { value: 'STAGE 2', label: 'STAGE 2', icon: 'fa-flag-checkered' },
    { value: 'STAGE 3', label: 'STAGE 3', icon: 'fa-trophy' },
    { value: 'STAGE 4', label: 'STAGE 4', icon: 'fa-certificate' },
    { value: 'ASSIGNED', label: 'ASSIGNED', icon: 'fa-user-check' },
    { value: 'CLOSED', label: 'CLOSED', icon: 'fa-circle-xmark' },
    { value: 'COPYRIGHT', label: 'COPYRIGHT', icon: 'fa-copyright' },
    { value: 'OPPOSITION', label: 'OPPOSITION', icon: 'fa-scale-balanced' },
    { value: 'REACTIVATION', label: 'REACTIVATION', icon: 'fa-rotate-right' },
    { value: 'IMPORTED DATA', label: 'IMPORTED DATA', icon: 'fa-file-import' }
];

const SUBSTATUS_OPTIONS = [
    { value: 'ACKNOWLEDGMENT RECEIVED', label: 'ACKNOWLEDGMENT RECEIVED', icon: 'fa-clipboard-check' },
    { value: 'FILED', label: 'FILED', icon: 'fa-file-signature' },
    { value: 'EXAMAMINATION DONE', label: 'EXAMAMINATION DONE', icon: 'fa-magnifying-glass' },
    { value: 'ACCEPTANCE DONE', label: 'ACCEPTANCE DONE', icon: 'fa-square-check' },
    { value: 'HEARING', label: 'HEARING', icon: 'fa-gavel' },
    { value: 'DEMAND NOTE RECEIVED', label: 'DEMAND NOTE RECEIVED', icon: 'fa-inbox' },
    { value: 'DEMAND NOTE SUBMITTED', label: 'DEMAND NOTE SUBMITTED', icon: 'fa-paper-plane' },
    { value: 'PUBLISHED', label: 'PUBLISHED', icon: 'fa-bullhorn' },
    { value: 'CERTIFICATE RECEIVED', label: 'CERTIFICATE RECEIVED', icon: 'fa-certificate' },
    { value: 'CER DISPATCH', label: 'CER DISPATCH', icon: 'fa-truck-fast' },
    { value: 'CER PRT', label: 'CER PRT', icon: 'fa-print' },
    { value: 'FAISAL (ISB)', label: 'FAISAL (ISB)', icon: 'fa-user' },
    { value: 'FAISAL (LHR)', label: 'FAISAL (LHR)', icon: 'fa-user' },
    { value: 'RASHID (LHR)', label: 'RASHID (LHR)', icon: 'fa-user' },
    { value: 'RASHID (KHI)', label: 'RASHID (KHI)', icon: 'fa-user' },
    { value: 'SULMAN (KRI)', label: 'SULMAN (KRI)', icon: 'fa-user' },
    { value: 'SULMAN (LHR)', label: 'SULMAN (LHR)', icon: 'fa-user' },
    { value: 'UZMA (KRI)', label: 'UZMA (KRI)', icon: 'fa-user' },
    { value: 'SULMAN (ISB)', label: 'SULMAN (ISB)', icon: 'fa-user' },
    { value: 'SULMAN ()', label: 'SULMAN ()', icon: 'fa-user' },
    { value: 'ABD / STOP', label: 'ABD / STOP', icon: 'fa-ban' },
    { value: 'OPPOSITION RECEIVED', label: 'OPPOSITION RECEIVED', icon: 'fa-inbox' },
    { value: 'OPPOSITION FILLED', label: 'OPPOSITION FILLED', icon: 'fa-file-arrow-up' },
    { value: 'RECEIVED', label: 'RECEIVED', icon: 'fa-inbox' },
    { value: 'WITHDRAWN', label: 'WITHDRAWN', icon: 'fa-arrow-rotate-left' },
    { value: 'ABANDONED', label: 'ABANDONED', icon: 'fa-circle-xmark' },
    { value: 'OLD RECORD', label: 'OLD RECORD', icon: 'fa-box-archive' }
];

const SUBSTATUS_BY_STATUS = {
    'STAGE 1': ['ACKNOWLEDGMENT RECEIVED', 'FILED'],
    'STAGE 2': ['EXAMAMINATION DONE', 'ACCEPTANCE DONE', 'HEARING'],
    'STAGE 3': ['DEMAND NOTE RECEIVED', 'DEMAND NOTE SUBMITTED'],
    'STAGE 4': ['PUBLISHED', 'CERTIFICATE RECEIVED', 'CER DISPATCH', 'CER PRT'],
    'ASSIGNED': ['FAISAL (ISB)', 'FAISAL (LHR)', 'RASHID (LHR)', 'RASHID (KHI)', 'SULMAN (KRI)', 'SULMAN (LHR)', 'UZMA (KRI)', 'SULMAN (ISB)', 'SULMAN ()'],
    'CLOSED': ['ABD / STOP'],
    'COPYRIGHT': ['FILED', 'OPPOSITION RECEIVED', 'OPPOSITION FILLED'],
    'OPPOSITION': ['RECEIVED', 'WITHDRAWN', 'FILED', 'ABANDONED'],
    'REACTIVATION': ['RECEIVED', 'FILED'],
    'IMPORTED DATA': ['OLD RECORD']
};

// Sample data based on the provided Google Sheet
let sheetData = [
    {
        dateTime: '2021-01-20T10:00:00',
        caseNo: 5,
        appName: 'TARGO',
        tmNo: 533589,
        classNo: 4,
        status: 'STAGE 3',
        subStatus: 'DEMAND NOTE SUBMITTED',
        notes: '',
        attachment: null
    },
    {
        dateTime: '2022-10-05T10:00:00',
        caseNo: 36,
        appName: 'ALKAIF PIZZA',
        tmNo: 583974,
        classNo: 43,
        status: 'IMPORTED DATA',
        subStatus: 'OLD RECORD',
        notes: '',
        attachment: null
    },
    {
        dateTime: '2023-12-31T10:00:00',
        caseNo: 47,
        appName: 'BEAU-TECH',
        tmNo: 565615,
        classNo: 3,
        status: 'STAGE 3',
        subStatus: 'DEMAND NOTE SUBMITTED',
        notes: 'SUBMITTED ON 03-Oct-2023',
        attachment: null
    }
];

let filteredData = [...sheetData];
let sortColumn = -1;
let sortDirection = 1;
let historyLog = [];
let viewMode = 'list';
let isCompact = false;

function getRowIndexByCaseNo(caseNo) {
    return sheetData.findIndex((r) => r.caseNo === caseNo);
}

function resolveDataIndex(filteredIndex) {
    const row = filteredData[filteredIndex];
    if (!row) return -1;
    return getRowIndexByCaseNo(row.caseNo);
}

// Initialize the application
document.addEventListener('DOMContentLoaded', function() {
    loadFromLocalStorage();
    loadHistoryFromLocalStorage();
    recalculateDerived();
    filteredData = [...sheetData];
    renderTable();
    renderCards();
    renderHistory();
    updateStatistics();
    setupEventListeners();
});

window.setActiveTab = setActiveTab;
window.clearHistory = clearHistory;
window.addNewRow = addNewRow;
window.exportData = exportData;
window.sortTable = sortTable;
window.makeEditable = makeEditable;
window.updateField = updateField;
window.deleteRow = deleteRow;
window.duplicateRow = duplicateRow;
window.attachFile = attachFile;
window.viewAttachment = viewAttachment;
window.setViewMode = setViewMode;
window.toggleCompact = toggleCompact;
window.handleRowClick = handleRowClick;

function formatDateTime(value) {
    if (!value) return '';
    const dt = new Date(value);
    if (Number.isNaN(dt.getTime())) return String(value);
    return dt.toLocaleString('en-GB', {
        day: '2-digit',
        month: 'short',
        year: '2-digit',
        hour: '2-digit',
        minute: '2-digit'
    }).replace(',', '');
}

function nowIsoLocal() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function statusMeta(value) {
    return STATUS_OPTIONS.find((s) => s.value === value) || { value, label: value, icon: 'fa-tag' };
}

function subStatusMeta(value) {
    return SUBSTATUS_OPTIONS.find((s) => s.value === value) || { value, label: value, icon: 'fa-tag' };
}

function getSubStatusOptionsForStatus(status) {
    const values = SUBSTATUS_BY_STATUS[status] || [];
    return values
        .map((v) => subStatusMeta(v))
        .filter((m) => m && m.value);
}

function normalizeRowStatus(row) {
    const allowed = getSubStatusOptionsForStatus(row.status).map((s) => s.value);
    if (allowed.length === 0) return;
    if (!allowed.includes(row.subStatus)) {
        row.subStatus = allowed[0];
    }
}

function logAction(entry) {
    historyLog.unshift({
        time: nowIsoLocal(),
        ...entry
    });
    saveHistoryToLocalStorage();
    renderHistory();
}

function setupEventListeners() {
    document.getElementById('searchInput').addEventListener('input', filterData);
    document.getElementById('statusFilter').addEventListener('change', filterData);
    document.getElementById('classFilter').addEventListener('change', filterData);
}

function setActiveTab(tab) {
    const sheetTab = document.getElementById('sheetTab');
    const historyTab = document.getElementById('historyTab');
    const tabSheet = document.getElementById('tabSheet');
    const tabHistory = document.getElementById('tabHistory');
    const sheetFilters = document.getElementById('sheetFilters');
    const sheetStats = document.getElementById('sheetStats');

    const isSheet = tab === 'sheet';
    sheetTab.classList.toggle('hidden', !isSheet);
    historyTab.classList.toggle('hidden', isSheet);
    sheetFilters.classList.toggle('hidden', !isSheet);
    sheetStats.classList.toggle('hidden', !isSheet);
    tabSheet.dataset.active = isSheet ? 'true' : 'false';
    tabHistory.dataset.active = isSheet ? 'false' : 'true';
}

function recalculateDerived() {
    const tmCounts = new Map();
    for (const row of sheetData) {
        normalizeRowStatus(row);
        const key = row.tmNo ? String(row.tmNo) : '';
        if (!key) continue;
        tmCounts.set(key, (tmCounts.get(key) || 0) + 1);
    }

    for (const row of sheetData) {
        const key = row.tmNo ? String(row.tmNo) : '';
        row.isDuplicate = key && (tmCounts.get(key) || 0) > 1;
        row.hasTmMatch = row.isDuplicate;
        const match = key
            ? sheetData.find((r) => String(r.tmNo) === key && r.caseNo !== row.caseNo)
            : null;
        row.lookupCaseNo = match ? match.caseNo : null;
    }
}

function setViewMode(mode) {
    viewMode = mode;
    const sheetTab = document.getElementById('sheetTab');
    const cardTab = document.getElementById('cardTab');
    const btnList = document.getElementById('btnListView');
    const btnCard = document.getElementById('btnCardView');
    if (!sheetTab || !cardTab || !btnList || !btnCard) return;

    const isList = mode === 'list';
    sheetTab.classList.toggle('hidden', !isList);
    cardTab.classList.toggle('hidden', isList);
    btnList.classList.toggle('nb-btn-secondary', isList);
    btnList.classList.toggle('nb-btn-ghost', !isList);
    btnCard.classList.toggle('nb-btn-secondary', !isList);
    btnCard.classList.toggle('nb-btn-ghost', isList);

    if (!isList) {
        renderCards();
    }
}

function toggleCompact() {
    isCompact = !isCompact;
    const root = document.body;
    if (!root) return;
    root.classList.toggle('compact', isCompact);
}

function handleRowClick(ev, caseNo) {
    const target = ev.target;
    if (!target) return;
    const interactive = target.closest('button, a, input, select, textarea, label');
    if (interactive) return;
    openRecord(caseNo);
}

function buildClassOptions(selected) {
    let html = '<option value="">—</option>';
    for (let i = 1; i <= 45; i += 1) {
        const isSel = Number(selected) === i;
        html += `<option value="${i}" ${isSel ? 'selected' : ''}>${String(i).padStart(2, '0')}</option>`;
    }
    return html;
}

function renderTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';

    filteredData.forEach((row, index) => {
        const status = statusMeta(row.status);
        const subStatus = subStatusMeta(row.subStatus);
        const subStatusOptions = getSubStatusOptionsForStatus(row.status);
        const tmMatch = row.hasTmMatch;
        const tmMatchIcon = tmMatch ? 'fa-magnifying-glass' : '';
        const tmMatchClass = tmMatch ? 'blink' : '';
        const lookupOk = row.lookupCaseNo !== null && row.lookupCaseNo !== undefined;
        const lookupPillClass = lookupOk ? 'pill-warn' : '';
        const lookupIcon = lookupOk ? 'fa-magnifying-glass' : 'fa-minus';
        const attachmentIcon = row.attachment ? 'fa-paperclip' : 'fa-upload';

        const tr = document.createElement('tr');
        tr.className = 'hover-row';
        tr.setAttribute('onclick', `handleRowClick(event, ${row.caseNo})`);
        tr.innerHTML = `
            <td class="px-4 py-3 text-sm nb-td">${formatDateTime(row.dateTime)}</td>
            <td class="px-4 py-3 text-sm font-black nb-td">${row.caseNo}</td>
            <td class="px-4 py-3 text-sm nb-td editable" onclick="makeEditable(this, ${index}, 'appName', 'text')">${row.appName || 'Click to edit'}</td>
            <td class="px-4 py-3 text-sm nb-td editable" onclick="makeEditable(this, ${index}, 'tmNo', 'number')">
                <div class="flex items-center gap-2">
                    <span>${row.tmNo ?? ''}</span>
                    ${tmMatch ? `<i class="fas ${tmMatchIcon} ${tmMatchClass}" style="color: var(--nb-accent)"></i>` : ''}
                </div>
            </td>
            <td class="px-4 py-3 text-sm nb-td">
                <select class="nb-select" onchange="updateField(${index}, 'classNo', this.value === '' ? null : Number(this.value))">
                    ${buildClassOptions(row.classNo)}
                </select>
            </td>

            <td class="px-4 py-3 text-sm nb-td">
                <div class="pill">
                    <i class="fas ${status.icon}"></i>
                    <select class="nb-select" onchange="updateField(${index}, 'status', this.value)">
                        ${STATUS_OPTIONS.map((s) => `<option value="${s.value}" ${row.status === s.value ? 'selected' : ''}>${s.label}</option>`).join('')}
                    </select>
                </div>
            </td>

            <td class="px-4 py-3 text-sm nb-td">
                <div class="pill">
                    <i class="fas ${subStatus.icon}"></i>
                    <select class="nb-select" onchange="updateField(${index}, 'subStatus', this.value)">
                        ${subStatusOptions.map((s) => `<option value="${s.value}" ${row.subStatus === s.value ? 'selected' : ''}>${s.label}</option>`).join('')}
                    </select>
                </div>
            </td>

            <td class="px-4 py-3 text-sm nb-td text-center">
                ${row.isDuplicate ? `<i class="fas fa-check" style="color: var(--nb-success); font-size: 18px"></i>` : ''}
            </td>

            <td class="px-4 py-3 text-sm nb-td text-center">
                <span class="pill ${lookupPillClass}">
                    <i class="fas ${lookupIcon}"></i>
                    ${lookupOk ? `CASE ${row.lookupCaseNo}` : '—'}
                </span>
            </td>

            <td class="px-4 py-3 text-sm nb-td text-center">
                <div class="flex items-center justify-center gap-2">
                    <input id="file_${row.caseNo}" type="file" class="hidden" onchange="attachFile(${index}, this.files[0])" />
                    <button class="nb-btn nb-btn-ghost" onclick="document.getElementById('file_${row.caseNo}').click()">
                        <i class="fas ${attachmentIcon}"></i>
                    </button>
                    ${row.attachment ? `<button class="nb-btn nb-btn-ghost" onclick="viewAttachment(${index})"><i class="fas fa-eye"></i></button>` : ''}
                </div>
            </td>

            <td class="px-4 py-3 text-sm nb-td text-gray-900 editable" onclick="makeEditable(this, ${index}, 'notes', 'text')">${row.notes || 'Click to add...'}</td>

            <td class="px-4 py-3 text-sm nb-td text-center">
                <button onclick="deleteRow(${index})" class="nb-btn nb-btn-ghost" style="border-color: var(--nb-danger); color: var(--nb-danger)">
                    <i class="fas fa-trash"></i>
                </button>
                <button onclick="duplicateRow(${index})" class="nb-btn nb-btn-ghost ml-2">
                    <i class="fas fa-copy"></i>
                </button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

function renderCards() {
    const body = document.getElementById('cardBody');
    if (!body) return;
    body.innerHTML = '';

    filteredData.forEach((row) => {
        const status = statusMeta(row.status);
        const subStatus = subStatusMeta(row.subStatus);
        const card = document.createElement('div');
        card.className = 'nb-card p-4 cursor-pointer';
        card.setAttribute('onclick', `openRecord(${row.caseNo})`);

        card.innerHTML = `
            <div class="flex items-start justify-between gap-3">
                <div>
                    <div class="text-sm font-black opacity-70">CASE</div>
                    <div class="text-2xl font-black">${row.caseNo}</div>
                </div>
                <div class="pill">
                    <i class="fas ${status.icon}"></i>
                    <span>${row.status}</span>
                </div>
            </div>

            <div class="mt-3 font-black">${row.appName || ''}</div>

            <div class="grid grid-cols-2 gap-2 mt-3">
                <div class="nb-card p-3" style="box-shadow:none">
                    <div class="text-xs font-black opacity-70">TM NO</div>
                    <div class="font-black">${row.tmNo ?? ''}</div>
                </div>
                <div class="nb-card p-3" style="box-shadow:none">
                    <div class="text-xs font-black opacity-70">CLASS</div>
                    <div class="font-black">${row.classNo ?? ''}</div>
                </div>
            </div>

            <div class="pill mt-3">
                <i class="fas ${subStatus.icon}"></i>
                <span>${row.subStatus}</span>
            </div>

            <div class="mt-3 text-sm font-bold opacity-80">${formatDateTime(row.dateTime)}</div>
        `;

        body.appendChild(card);
    });
}

function makeEditable(cell, rowIndex, field, inputType) {
    const dataIndex = resolveDataIndex(rowIndex);
    if (dataIndex < 0) return;
    const currentValue = sheetData[dataIndex][field];
    const input = document.createElement('input');
    input.type = inputType || 'text';
    input.value = currentValue || '';
    input.className = 'edit-input';
    
    input.onblur = function() {
        updateField(rowIndex, field, input.type === 'number' ? (this.value === '' ? null : Number(this.value)) : this.value);
        renderTable();
    };
    
    input.onkeydown = function(e) {
        if (e.key === 'Enter') {
            this.blur();
        } else if (e.key === 'Escape') {
            recalculateDerived();
            renderTable();
        }
    };
    
    cell.innerHTML = '';
    cell.appendChild(input);
    input.focus();
}

function updateField(rowIndex, field, value) {
    const dataIndex = resolveDataIndex(rowIndex);
    if (dataIndex < 0) return;
    const prev = sheetData[dataIndex][field];
    sheetData[dataIndex][field] = value;

    const prevDateTime = sheetData[dataIndex].dateTime;
    sheetData[dataIndex].dateTime = nowIsoLocal();

    if (field === 'status') {
        const beforeSub = sheetData[dataIndex].subStatus;
        normalizeRowStatus(sheetData[dataIndex]);
        if (beforeSub !== sheetData[dataIndex].subStatus) {
            logAction({
                caseNo: sheetData[dataIndex].caseNo,
                action: 'AUTO',
                field: 'subStatus',
                from: beforeSub,
                to: sheetData[dataIndex].subStatus
            });
        }
    }

    recalculateDerived();
    filterData();
    updateStatistics();
    saveToLocalStorage();
    logAction({
        caseNo: sheetData[dataIndex].caseNo,
        action: 'UPDATE',
        field,
        from: prev,
        to: value
    });

    if (prevDateTime !== sheetData[dataIndex].dateTime) {
        logAction({
            caseNo: sheetData[dataIndex].caseNo,
            action: 'TOUCH',
            field: 'dateTime',
            from: prevDateTime,
            to: sheetData[dataIndex].dateTime
        });
    }
}

function filterData() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase();
    const statusFilter = document.getElementById('statusFilter').value;
    const classFilter = document.getElementById('classFilter').value;

    const safeToString = (value) => {
        if (value === null || value === undefined) return '';
        if (typeof value === 'object') {
            if (value.name) return String(value.name);
            try {
                return JSON.stringify(value);
            } catch {
                return '';
            }
        }
        return String(value);
    };

    filteredData = sheetData.filter(row => {
        const matchesSearch = !searchTerm || 
            Object.values(row).some(value => 
                safeToString(value).toLowerCase().includes(searchTerm)
            );
        
        const matchesStatus = !statusFilter || row.status === statusFilter;
        const matchesClass = !classFilter || String(row.classNo) === classFilter;

        return matchesSearch && matchesStatus && matchesClass;
    });

    renderTable();
    renderCards();
    updateStatistics();
}

function sortTable(columnIndex) {
    if (sortColumn === columnIndex) {
        sortDirection *= -1;
    } else {
        sortColumn = columnIndex;
        sortDirection = 1;
    }

    const fields = ['dateTime', 'caseNo', 'appName', 'tmNo', 'classNo', 'status', 'subStatus', 'isDuplicate', 'lookupCaseNo', 'attachment', 'notes'];
    const field = fields[columnIndex];

    filteredData.sort((a, b) => {
        let aVal = a[field];
        let bVal = b[field];

        if (typeof aVal === 'string') {
            aVal = aVal.toLowerCase();
            bVal = bVal.toLowerCase();
        }

        if (aVal < bVal) return -1 * sortDirection;
        if (aVal > bVal) return 1 * sortDirection;
        return 0;
    });

    renderTable();
}

function updateStatistics() {
    document.getElementById('totalCount').textContent = filteredData.length;
    document.getElementById('importedCount').textContent = filteredData.filter(row => row.status === 'IMPORTED DATA').length;
    document.getElementById('dupeCount').textContent = filteredData.filter(row => row.isDuplicate).length;
    document.getElementById('notesCount').textContent = filteredData.filter(row => row.notes && row.notes.trim() !== '').length;
}

function addNewRow() {
    const nextCase = Math.max(0, ...sheetData.map((r) => Number(r.caseNo) || 0)) + 1;
    
    const newRow = {
        dateTime: nowIsoLocal(),
        caseNo: nextCase,
        appName: 'NEW APPLICATION',
        tmNo: null,
        classNo: null,
        status: 'IMPORTED DATA',
        subStatus: 'OLD RECORD',
        notes: '',
        attachment: null
    };

    sheetData.push(newRow);
    recalculateDerived();
    filterData();
    updateStatistics();
    saveToLocalStorage();
    renderCards();
    logAction({
        caseNo: newRow.caseNo,
        action: 'ADD',
        field: '',
        from: '',
        to: ''
    });
}

function deleteRow(index) {
    if (confirm('Are you sure you want to delete this row?')) {
        const target = filteredData[index];
        const originalIndex = sheetData.findIndex(row => row.caseNo === target.caseNo);
        sheetData.splice(originalIndex, 1);
        recalculateDerived();
        filterData();
        updateStatistics();
        saveToLocalStorage();
        renderCards();
        logAction({
            caseNo: target.caseNo,
            action: 'DELETE',
            field: '',
            from: '',
            to: ''
        });
    }
}

function duplicateRow(index) {
    const originalRow = filteredData[index];
    const newRow = { ...originalRow };
    newRow.caseNo = Math.max(0, ...sheetData.map((r) => Number(r.caseNo) || 0)) + 1;
    newRow.dateTime = nowIsoLocal();
    newRow.appName = (originalRow.appName || '') + ' (COPY)';
    newRow.notes = (originalRow.notes || '');
    newRow.attachment = null;
    
    sheetData.push(newRow);
    recalculateDerived();
    filterData();
    updateStatistics();
    saveToLocalStorage();
    renderCards();
    logAction({
        caseNo: newRow.caseNo,
        action: 'DUPLICATE',
        field: '',
        from: originalRow.caseNo,
        to: newRow.caseNo
    });
}

function exportData() {
    let csv = 'DATE TIME,CASE NO,APP NAME,TM NO,CLASS,APPLICATION STATUS,APPLICATION SUB STATUS,DUP,LOOKUP,NOTES,ATTACHMENT\n';
    
    filteredData.forEach(row => {
        csv += `"${formatDateTime(row.dateTime)}","${row.caseNo}","${row.appName}","${row.tmNo ?? ''}","${row.classNo ?? ''}","${row.status}","${row.subStatus}","${row.isDuplicate ? 'TRUE' : 'FALSE'}","${row.lookupCaseNo ?? ''}","${row.notes}","${row.attachment ? row.attachment.name : ''}"\n`;
    });

    const blob = new Blob([csv], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `data_sheet_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    window.URL.revokeObjectURL(url);
}

function saveToLocalStorage() {
    localStorage.setItem(STORAGE_KEY_DATA, JSON.stringify(sheetData));
}

function loadFromLocalStorage() {
    const saved = localStorage.getItem(STORAGE_KEY_DATA);
    if (saved) {
        sheetData = JSON.parse(saved);
    }
}

function saveHistoryToLocalStorage() {
    localStorage.setItem(STORAGE_KEY_HISTORY, JSON.stringify(historyLog));
}

function loadHistoryFromLocalStorage() {
    const saved = localStorage.getItem(STORAGE_KEY_HISTORY);
    if (saved) {
        historyLog = JSON.parse(saved);
    }
}

function renderHistory() {
    const body = document.getElementById('historyBody');
    if (!body) return;
    body.innerHTML = '';

    historyLog.slice(0, 500).forEach((h) => {
        const tr = document.createElement('tr');
        tr.className = 'hover-row';
        tr.innerHTML = `
            <td class="px-4 py-3 text-sm nb-td">${formatDateTime(h.time)}</td>
            <td class="px-4 py-3 text-sm nb-td font-black">${h.caseNo ?? ''}</td>
            <td class="px-4 py-3 text-sm nb-td">${h.action ?? ''}</td>
            <td class="px-4 py-3 text-sm nb-td">${h.field ?? ''}</td>
            <td class="px-4 py-3 text-sm nb-td">${h.from ?? ''}</td>
            <td class="px-4 py-3 text-sm nb-td">${h.to ?? ''}</td>
        `;
        body.appendChild(tr);
    });
}

function clearHistory() {
    if (!confirm('Clear all history?')) return;
    historyLog = [];
    saveHistoryToLocalStorage();
    renderHistory();
}

function attachFile(rowIndex, file) {
    if (!file) return;
    const dataIndex = resolveDataIndex(rowIndex);
    if (dataIndex < 0) return;
    const prev = sheetData[dataIndex].attachment;
    sheetData[dataIndex].attachment = {
        name: file.name,
        size: file.size,
        type: file.type,
        lastModified: file.lastModified
    };

    const prevDateTime = sheetData[dataIndex].dateTime;
    sheetData[dataIndex].dateTime = nowIsoLocal();

    saveToLocalStorage();
    renderTable();
    renderCards();
    logAction({
        caseNo: sheetData[dataIndex].caseNo,
        action: 'ATTACH',
        field: 'attachment',
        from: prev ? prev.name : '',
        to: file.name
    });

    logAction({
        caseNo: sheetData[dataIndex].caseNo,
        action: 'TOUCH',
        field: 'dateTime',
        from: prevDateTime,
        to: sheetData[dataIndex].dateTime
    });
}

function viewAttachment(rowIndex) {
    const dataIndex = resolveDataIndex(rowIndex);
    if (dataIndex < 0) return;
    const att = sheetData[dataIndex].attachment;
    if (!att) return;
    alert(`Attached: ${att.name}`);
}

function openRecord(caseNo) {
    const row = sheetData.find((r) => r.caseNo === caseNo);
    if (!row) return;
    const modal = document.getElementById('recordModal');
    const details = document.getElementById('recordDetails');
    const log = document.getElementById('recordLog');
    const title = document.getElementById('recordTitle');

    if (!modal || !details || !log || !title) return;

    title.textContent = `CASE ${row.caseNo} — ${row.appName || ''}`;
    details.innerHTML = `
        <div class="grid grid-cols-1 md:grid-cols-2 gap-3">
            <div class="nb-card p-3" style="box-shadow:none">
                <div class="font-black">DATE TIME</div>
                <div class="font-bold opacity-80 mt-1">${formatDateTime(row.dateTime)}</div>
            </div>
            <div class="nb-card p-3" style="box-shadow:none">
                <div class="font-black">TM NO</div>
                <div class="font-bold opacity-80 mt-1">${row.tmNo ?? ''}</div>
            </div>
            <div class="nb-card p-3" style="box-shadow:none">
                <div class="font-black">CLASS</div>
                <div class="font-bold opacity-80 mt-1">${row.classNo ?? ''}</div>
            </div>
            <div class="nb-card p-3" style="box-shadow:none">
                <div class="font-black">STATUS</div>
                <div class="font-bold opacity-80 mt-1">${row.status} / ${row.subStatus}</div>
            </div>
            <div class="nb-card p-3 md:col-span-2" style="box-shadow:none">
                <div class="font-black">NOTES</div>
                <div class="font-bold opacity-80 mt-1 whitespace-pre-wrap">${row.notes || ''}</div>
            </div>
        </div>
    `;

    const items = historyLog.filter((h) => h.caseNo === caseNo).slice(0, 200);
    log.innerHTML = items.map((h) => `
        <div class="nb-card p-3" style="box-shadow:none">
            <div class="flex items-center justify-between gap-3">
                <div class="font-black">${h.action}</div>
                <div class="font-bold opacity-70">${formatDateTime(h.time)}</div>
            </div>
            <div class="font-bold opacity-80 mt-1">${h.field || ''}</div>
            <div class="text-sm opacity-80 mt-1">${(h.from ?? '') === '' && (h.to ?? '') === '' ? '' : `${h.from ?? ''} → ${h.to ?? ''}`}</div>
        </div>
    `).join('');

    modal.classList.remove('hidden');
}

function closeRecord() {
    const modal = document.getElementById('recordModal');
    if (!modal) return;
    modal.classList.add('hidden');
}

window.openRecord = openRecord;
window.closeRecord = closeRecord;
