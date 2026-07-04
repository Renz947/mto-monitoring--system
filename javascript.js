// ====================================================================
// 🔐 AUTHENTICATION & SESSION MANAGEMENT
// ====================================================================
const MTO_SESSION_KEY = 'mtoAuth_session_v1';

/** Reads the session from localStorage or sessionStorage */
function getSession() {
    try {
        const raw = localStorage.getItem(MTO_SESSION_KEY) || sessionStorage.getItem(MTO_SESSION_KEY);
        return raw ? JSON.parse(raw) : null;
    } catch(e) { return null; }
}

/** Logs out the current user and redirects to login page */
function logout() {
    if (!confirm('Sigurado ka bang nais mong mag-logout sa MTO System Monitoring?')) return;
    try {
        localStorage.removeItem(MTO_SESSION_KEY);
        sessionStorage.removeItem(MTO_SESSION_KEY);
    } catch(e) { /* ignore */ }
    window.location.replace('login.html');
}

/** Populates the sidebar user info bar with session data */
function initUserInfoBar() {
    const session = getSession();
    if (!session) return;

    const nameEl   = document.getElementById('userDisplayName');
    const roleEl   = document.getElementById('userRoleBadge');
    const avatarEl = document.getElementById('userAvatarBox');

    if (nameEl)   nameEl.textContent   = session.name  || session.username || 'User';
    if (roleEl)   roleEl.textContent   = session.role  || 'Operator';
    if (avatarEl) avatarEl.textContent = (session.name || session.username || 'U')[0].toUpperCase();
}

// ====================================================================
// 📊 GLOBAL DATA STATE UTILITIES
// ====================================================================
let globalData = []; 
let fileType = ""; 
let _skipChartRedraw = false; // perf: skip expensive chart redraw on row toggles
let genericHeaders = []; // Para sa dynamic generic mapping
let parsedEmployeesData = {};
let currentActiveEmployee = "";
let manualAttendanceLogs = []; // {id, name, date, time, status} — source of truth ng Live Log Feed
let manualLogsCounter = 0;
let _manualLogIdSeq = 0; // unique id generator para sa manual logs (di nagbabago kahit may delete)
let chartInstancePie = null;
let chartInstanceBar3D = null;
// Live Excel Modal State
let _liveExcelWorkbook = null;
let _liveExcelSelectedSheet = null;
let _liveExcelFile = null;
// 🔴 LIVE LOCAL FILE WATCH STATE — File System Access API (Chrome/Edge lang)
let _pendingLiveFileHandle = null; // di pa committed na handle, galing sa picker, naghihintay sa "I-load sa Dashboard"
let _liveFileHandle = null;        // ACTIVE na FileSystemFileHandle na pinapanood (null = walang live sync)
let _liveFileLastModified = null;  // huling nakitang lastModified timestamp, para malaman kung nag-iba
let _livePollIntervalId = null;    // setInterval id ng polling loop
let _liveFileFailCount = 0;        // bilang ng magkakasunod na error bago tuluyang itigil
const LIVE_POLL_MS = 4000;         // bawat ilang segundo titingnan kung na-save ulit ang file

// 🔗 LIVE GOOGLE SHEETS WATCH STATE — gumagana sa LAHAT ng browser/computer, di tulad ng local file link
let _liveSheetsUrl = null;          // ang CSV export URL na aktibong pinapanood
let _liveSheetsLastContent = null;  // huling nakuhang CSV text, para malaman kung may pagbabago
let _liveSheetsPollIntervalId = null;
let _liveSheetsFetching = false;          // guard: huwag mag-overlap ng concurrent fetches
const LIVE_SHEETS_POLL_MS = 3000;         // 3s pag visible ang tab (bilis 2x vs 6s dati)
const LIVE_SHEETS_POLL_HIDDEN_MS = 10000; // 10s lang pag naka-background ang tab (makatipid)
// 🛠️ PARA PALITAN ANG SHEET NA AWTOMATIKONG GAGAMITIN NG LAHAT NG VISITOR:
// Buksan ang gustong tab sa Google Sheets, kopyahin ang URL mula sa address bar,
// at i-paste dito. Iwanang '' kung walang dapat awtomatikong mag-live sa simula.
const DEFAULT_LIVE_SHEETS_URL = ''; // ← I-paste dito ang CSV URL mula sa sarili mong "Publish to web" dialog
const DAYS_OF_WEEK = ["SUNDAY", "MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY"];

// 🎯 AUTO-DETECT RECORD LOOKUP STATE — ginagamit para hindi paulit-ulit
// lumabas ang popup habang pareho pa rin ang laman ng search bar.
let _lastAutoDetectQuery = null;

// 🔴 DUPLICATE ID DETECTION STATE
// Ito ang nagsu-store ng signature ng huling na-detect na duplicate set
// para hindi paulit-ulit mag-pop up ang modal sa bawat background poll
// kapag hindi naman nagbago ang duplicates.
let _lastDuplicateSignature = null;
let _currentDuplicateGroups = [];  // buong group data para sa re-open ng modal

// ====================================================================
// 💾 PERSISTENCE LAYER — auto-save sa browser (localStorage) para hindi
// nawawala ang na-upload na data kapag nag-reload ng page. Ang "Tanggalin
// ang Data" button ang gagamitin para sadyang burahin ang naka-save na data.
// ====================================================================
const STORAGE_KEY = 'mtoSystemMonitoring_savedState_v1';
let _saveDebounceTimer = null;

function saveAppState() {
    try {
        const isAttendanceTabActive = !document.getElementById('attendanceView').classList.contains('hidden');
        const titleEl = document.getElementById('tableTitle');
        const state = {
            globalData,
            fileType,
            genericHeaders,
            currentFileTitle: titleEl ? titleEl.innerText : '',
            parsedEmployeesData,
            currentActiveEmployee,
            manualAttendanceLogs,
            manualLogsCounter,
            _manualLogIdSeq,
            activeView: isAttendanceTabActive ? 'attendance' : 'dashboard'
        };
        localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
        // Karaniwang dahilan: puno na ang storage ng browser (mahigpit ang limitasyon, ~5-10MB).
        console.warn('Hindi na-save ang data sa browser storage:', e);
    }
}

// Debounced version — gamit para sa madalas na pag-type (hal. number inputs)
// para hindi mag-save sa kada letrang i-type ng user.
function saveAppStateDebounced() {
    clearTimeout(_saveDebounceTimer);
    _saveDebounceTimer = setTimeout(saveAppState, 400);
}

function loadAppState() {
    let raw;
    try {
        raw = localStorage.getItem(STORAGE_KEY);
    } catch (e) {
        console.warn('Hindi ma-access ang browser storage:', e);
        return;
    }
    if (!raw) return;

    let state;
    try {
        state = JSON.parse(raw);
    } catch (e) {
        console.warn('Sira ang naka-save na data, lalaktawan:', e);
        return;
    }

    globalData = Array.isArray(state.globalData) ? state.globalData : [];
    fileType = state.fileType || "";
    genericHeaders = Array.isArray(state.genericHeaders) ? state.genericHeaders : [];
    parsedEmployeesData = state.parsedEmployeesData && typeof state.parsedEmployeesData === 'object' ? state.parsedEmployeesData : {};
    currentActiveEmployee = state.currentActiveEmployee || "";
    manualAttendanceLogs = Array.isArray(state.manualAttendanceLogs) ? state.manualAttendanceLogs : [];
    manualLogsCounter = Number(state.manualLogsCounter) || 0;
    _manualLogIdSeq = Number(state._manualLogIdSeq) || 0;

    // Ibalik ang pamagat ng file sa header ng table
    if (state.currentFileTitle) {
        const titleEl = document.getElementById('tableTitle');
        if (titleEl) titleEl.innerText = state.currentFileTitle;
    }

    // Ibalik ang Dashboard Monitoring data (kung meron)
    if (globalData.length > 0) {
        document.getElementById('searchBar').disabled = false;
        document.getElementById('searchBar').placeholder = "Mag-type rito para mag-filter ng data...";
        buildDashboard(globalData);
        // 🔴 Badge lang ang ipapakita kapag may dupes sa restored data — walang auto-popup
        const _rg = detectDuplicateIDs(globalData);
        if (_rg.length > 0) {
            _currentDuplicateGroups = _rg;
            _lastDuplicateSignature = _rg.map(g=>g.id+':'+g.rows.length).sort().join('|');
            _showDuplicateBadge(_rg.length);
        }
    }

    // Ibalik ang Timesheet / Attendance file data (kung meron)
    if (Object.keys(parsedEmployeesData).length > 0) {
        refreshEditorUI();
    }

    // Ibalik ang Manual Attendance Logger feed (kung meron)
    renderManualLogs();

    // Ibalik ang huling aktibong tab (Dashboard o Attendance)
    if (state.activeView === 'attendance') {
        document.getElementById('btnAttendance').click();
    }
}

// Tinatanggal lahat ng naka-save na CSV/Excel data (dashboard, timesheet, at manual logs).
// Ginagamit ng "🗑️ Tanggalin ang Data" button sa sidebar.
function clearAppState() {
    const meronLaman = globalData.length > 0 || Object.keys(parsedEmployeesData).length > 0 || manualAttendanceLogs.length > 0;
    const msg = meronLaman
        ? "Sigurado ka bang nais mong tanggalin ang LAHAT ng na-upload na CSV/Excel data, kasama ang manual attendance logs? Hindi na ito mababawi."
        : "Walang naka-save na data sa ngayon. Sigurado ka bang nais mong magpatuloy?";
    if (!confirm(msg)) return;

    try {
        localStorage.removeItem(STORAGE_KEY);
    } catch (e) {
        console.warn('Hindi na-clear ang browser storage:', e);
        alert('⚠️ Hindi na-clear ang data. Posibleng naka-disable ang storage ng browser mo (hal. private/incognito mode o mahigpit na privacy settings).');
        return;
    }
    alert('✅ Tagumpay! Nabura na ang lahat ng naka-save na CSV/Excel data. Mag-re-reload na ngayon ang page.');
    // Reload para sigurado ang malinis na pagsisimula ng buong UI (charts, table, forms, atbp.)
    location.reload();
}

// ====================================================================
// 🗂️ SAVE FILE HISTORY — payroll cutoff snapshot archive
// ----------------------------------------------------------------------
// "🔄 Automatic Save" (sa Attendance Monitoring) ay kumukuha ng snapshot
// ng KASALUKUYANG timesheet data (lahat ng empleyado) at itinatago ito
// dito bilang bagong entry — naka-tag sa cutoff ng payroll (default:
// Ika-7 hanggang Ika-22 ng buwan). Ang "💾 Save" ay plain save lamang
// (walang idinadagdag sa history). Lahat ng ito ay naka-view/buksan sa
// loob mismo ng website — walang kailangang i-download na file.
// ====================================================================
let saveFileHistory = [];
let _saveHistoryIdSeq = 0;
const SAVE_HISTORY_KEY = 'mtoSystemMonitoring_saveFileHistory_v1';

// 🛠️ PALITAN DITO kung magbabago ang cutoff schedule ng payroll niyo.
// Default: 7 hanggang 22 ng buwan. Ang kabilang kalahati ng buwan
// (23 hanggang 6 ng susunod na buwan) ay awtomatikong nako-compute na rin.
const PAYROLL_CUTOFF_START_DAY = 7;
const PAYROLL_CUTOFF_END_DAY = 22;

/** Nagbabalik ng { label, startDate, endDate } base sa kasalukuyang petsa,
 *  ayon sa cutoff schedule sa itaas. */
function getCurrentCutoffInfo(refDate) {
    const d = refDate ? new Date(refDate) : new Date();
    const day = d.getDate(), month = d.getMonth(), year = d.getFullYear();
    const MONTHS = ["Enero","Pebrero","Marso","Abril","Mayo","Hunyo","Hulyo","Agosto","Setyembre","Oktubre","Nobyembre","Disyembre"];
    let startDate, endDate;

    if (day >= PAYROLL_CUTOFF_START_DAY && day <= PAYROLL_CUTOFF_END_DAY) {
        startDate = new Date(year, month, PAYROLL_CUTOFF_START_DAY);
        endDate = new Date(year, month, PAYROLL_CUTOFF_END_DAY);
    } else if (day > PAYROLL_CUTOFF_END_DAY) {
        startDate = new Date(year, month, PAYROLL_CUTOFF_END_DAY + 1);
        endDate = new Date(year, month + 1, PAYROLL_CUTOFF_START_DAY - 1);
    } else {
        startDate = new Date(year, month - 1, PAYROLL_CUTOFF_END_DAY + 1);
        endDate = new Date(year, month, PAYROLL_CUTOFF_START_DAY - 1);
    }

    const sameMonth = startDate.getMonth() === endDate.getMonth() && startDate.getFullYear() === endDate.getFullYear();
    const label = sameMonth
        ? `${MONTHS[startDate.getMonth()]} ${startDate.getDate()}–${endDate.getDate()}, ${endDate.getFullYear()}`
        : `${MONTHS[startDate.getMonth()]} ${startDate.getDate()} – ${MONTHS[endDate.getMonth()]} ${endDate.getDate()}, ${endDate.getFullYear()}`;

    return { label, startDate, endDate };
}

/** Kabuuang bilang/halaga sa LAHAT ng empleyado (hindi lang ang kasalukuyang pinili). */
function calculateGrandTotals() {
    const empNames = Object.keys(parsedEmployeesData);
    let totalDaysPresent = 0, totalBasic = 0, totalAllowance = 0, totalOTHours = 0, totalOTPay = 0;
    empNames.forEach(name => {
        (parsedEmployeesData[name] || []).forEach(rec => {
            if (rec.remarks === "PRESENT" || rec.remarks === "DOUBLE PAY") totalDaysPresent++;
            totalBasic += Number(rec.dailyRate) || 0;
            totalAllowance += Number(rec.allowance) || 0;
            totalOTHours += Number(rec.otHours) || 0;
            if ((Number(rec.otHours) || 0) > 0) {
                const r = String(rec.shift).toUpperCase().includes("AM") ? 93 : 101;
                totalOTPay += r * (Number(rec.otHours) || 0);
            }
        });
    });
    return {
        totalEmployees: empNames.length,
        totalDaysPresent,
        totalBasic,
        totalAllowance,
        totalOTHours,
        totalOTPay,
        totalGross: totalBasic + totalAllowance + totalOTPay
    };
}

function persistSaveFileHistory() {
    try {
        localStorage.setItem(SAVE_HISTORY_KEY, JSON.stringify({ saveFileHistory, _saveHistoryIdSeq }));
    } catch (e) {
        console.warn('Hindi na-save ang Save File History:', e);
        showToast('⚠️ Hindi na-save sa history — posibleng puno na ang storage ng browser.', true);
    }
}

function loadSaveFileHistory() {
    try {
        const raw = localStorage.getItem(SAVE_HISTORY_KEY);
        if (!raw) return;
        const parsed = JSON.parse(raw);
        saveFileHistory = Array.isArray(parsed.saveFileHistory) ? parsed.saveFileHistory : [];
        _saveHistoryIdSeq = Number(parsed._saveHistoryIdSeq) || 0;
    } catch (e) {
        console.warn('Hindi ma-access ang naka-save na Save File History:', e);
    }
}

/** Karaniwang lohika ng paggawa ng bagong Save File History entry mula sa
 *  kasalukuyang attendance data. Ginagamit ng "💾 Save" (auto-label, walang
 *  prompt) at ng "🔄 Automatic Save" (may prompt para sa custom label). */
function addSnapshotToHistory(label) {
    const cutoff = getCurrentCutoffInfo();
    const totals = calculateGrandTotals();
    _saveHistoryIdSeq++;
    saveFileHistory.unshift({
        id: _saveHistoryIdSeq,
        label: label,
        savedAt: new Date().toISOString(),
        cutoffLabel: cutoff.label,
        employees: JSON.parse(JSON.stringify(parsedEmployeesData)),
        totals
    });
    persistSaveFileHistory();
    updateSaveHistoryBadge();
    return cutoff;
}

/** 🔄 AUTOMATIC SAVE — kumukuha ng snapshot ng buong attendance data ngayon
 *  at itinatago ito bilang bagong entry sa Save File History, naka-tag sa
 *  kasalukuyang cutoff ng payroll. May prompt para sa custom na label. */
function autoSaveToHistory() {
    const empNames = Object.keys(parsedEmployeesData);
    if (empNames.length === 0) {
        alert("⚠️ Wala pang datos sa Attendance Monitoring na pwedeng i-save. Mag-lagay muna ng empleyado o mag-upload ng timesheet file.");
        return;
    }

    const cutoff = getCurrentCutoffInfo();
    const defaultLabel = `Cut-off ${cutoff.label}`;
    const userLabel = prompt("📌 Pangalan ng Save (cutoff ng payroll):", defaultLabel);
    if (userLabel === null) return; // kinansela ng user

    const finalLabel = (userLabel.trim() || defaultLabel);
    addSnapshotToHistory(finalLabel);
    saveAppState();
    showToast(`✅ Na-save sa Save File History: "${finalLabel}"`);
}

/** 💾 SAVE — mabilis na manual save ng kasalukuyang datos. Awtomatiko na
 *  itong nadadagdag bilang bagong entry sa Save File History (auto-labeled
 *  gamit ang oras ng pag-save, walang abalang prompt), basta may laman ang
 *  Attendance Monitoring — kaya lagi itong makikita sa "🗂️ Save File History". */
function manualSaveOnly() {
    saveAppState();
    const empNames = Object.keys(parsedEmployeesData);
    if (empNames.length > 0) {
        const stamp = new Date().toLocaleString('en-PH', { dateStyle: 'medium', timeStyle: 'short' });
        const cutoff = addSnapshotToHistory(`Save — ${stamp}`);
        showToast(`✅ Na-save at naidagdag sa Save File History! (Cut-off ${cutoff.label})`);
    } else {
        showToast("✅ Na-save ang kasalukuyang datos!");
    }
}

function updateSaveHistoryBadge() {
    const badge = document.getElementById('saveHistoryCountBadge');
    if (!badge) return;
    if (saveFileHistory.length > 0) {
        badge.textContent = saveFileHistory.length;
        badge.classList.remove('hidden');
    } else {
        badge.classList.add('hidden');
    }
}

function openSaveHistoryModal() {
    document.getElementById('saveHistoryModal').classList.remove('hidden');
    showSaveHistoryListView();
}

function closeSaveHistoryModal() {
    document.getElementById('saveHistoryModal').classList.add('hidden');
}

function showSaveHistoryListView() {
    document.getElementById('saveHistoryListView').classList.remove('hidden');
    document.getElementById('saveHistoryDetailView').classList.add('hidden');
    renderSaveHistoryList();
}

function renderSaveHistoryList() {
    const container = document.getElementById('saveHistoryListBody');
    if (!container) return;

    if (saveFileHistory.length === 0) {
        container.innerHTML = `<div class="py-16 text-center text-slate-400 dark:text-slate-500 text-sm">📭 Wala pang naka-save na file history.<br>Pindutin ang <strong>"🔄 Automatic Save"</strong> sa Attendance Monitoring para gumawa ng unang save.</div>`;
        return;
    }

    container.innerHTML = saveFileHistory.map(entry => {
        const savedLabel = new Date(entry.savedAt).toLocaleString('en-PH', { dateStyle: 'medium', timeStyle: 'short' });
        return `
        <div class="bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-800 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div class="min-w-0">
                <p class="font-black text-sm text-slate-800 dark:text-slate-100 truncate">📁 ${escapeHtml(entry.label)}</p>
                <p class="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5">🕒 Na-save: ${savedLabel} &nbsp;•&nbsp; 🗓️ Cut-off: ${escapeHtml(entry.cutoffLabel)}</p>
                <p class="text-[11px] text-slate-500 dark:text-slate-400 mt-0.5">👥 ${entry.totals.totalEmployees} empleyado &nbsp;•&nbsp; 📅 ${entry.totals.totalDaysPresent} total days present</p>
            </div>
            <div class="flex items-center gap-3 flex-shrink-0">
                <div class="text-right">
                    <p class="text-[10px] text-slate-400 uppercase font-bold">Total Gross</p>
                    <p class="text-sm font-black text-emerald-600 dark:text-emerald-400">₱${entry.totals.totalGross.toLocaleString('en-US', {minimumFractionDigits:2})}</p>
                </div>
                <div class="flex gap-1.5">
                    <button onclick="viewSaveHistoryEntry(${entry.id})" title="Buksan / Tingnan" class="w-8 h-8 flex items-center justify-center rounded-lg bg-orange-100 dark:bg-orange-900/30 text-orange-600 dark:text-orange-400 hover:bg-orange-600 hover:text-white cursor-pointer transition-colors">👁️</button>
                    <button onclick="deleteSaveHistoryEntry(${entry.id})" title="Burahin" class="w-8 h-8 flex items-center justify-center rounded-lg bg-red-100 dark:bg-red-900/30 text-red-500 hover:bg-red-600 hover:text-white cursor-pointer transition-colors">🗑️</button>
                </div>
            </div>
        </div>`;
    }).join('');
}

/** Buksan/tingnan ang detalye ng isang naunang save (view lang — sa loob mismo ng website). */
function viewSaveHistoryEntry(id) {
    const entry = saveFileHistory.find(e => e.id === id);
    if (!entry) return;

    document.getElementById('saveHistoryListView').classList.add('hidden');
    document.getElementById('saveHistoryDetailView').classList.remove('hidden');

    const savedLabel = new Date(entry.savedAt).toLocaleString('en-PH', { dateStyle: 'medium', timeStyle: 'short' });
    document.getElementById('saveHistoryDetailTitle').textContent = entry.label;
    document.getElementById('saveHistoryDetailSubtitle').textContent = `🕒 Na-save: ${savedLabel}   •   🗓️ Cut-off: ${entry.cutoffLabel}`;
    document.getElementById('saveHistoryDetailEmployees').textContent = entry.totals.totalEmployees;
    document.getElementById('saveHistoryDetailDays').textContent = entry.totals.totalDaysPresent;
    document.getElementById('saveHistoryDetailOT').textContent = entry.totals.totalOTHours + " hrs";
    document.getElementById('saveHistoryDetailGross').textContent = "₱" + entry.totals.totalGross.toLocaleString('en-US', {minimumFractionDigits:2});

    const tbody = document.getElementById('saveHistoryDetailTableBody');
    const empNames = Object.keys(entry.employees);
    tbody.innerHTML = empNames.length === 0 ? `<tr><td colspan="5" class="p-6 text-center text-slate-400">Walang empleyado sa save na ito.</td></tr>` : empNames.map(name => {
        const records = entry.employees[name] || [];
        let daysPresent = 0, basic = 0, allowance = 0, otHours = 0, otPay = 0;
        records.forEach(rec => {
            if (rec.remarks === "PRESENT" || rec.remarks === "DOUBLE PAY") daysPresent++;
            basic += Number(rec.dailyRate) || 0;
            allowance += Number(rec.allowance) || 0;
            otHours += Number(rec.otHours) || 0;
            if ((Number(rec.otHours) || 0) > 0) {
                const r = String(rec.shift).toUpperCase().includes("AM") ? 93 : 101;
                otPay += r * (Number(rec.otHours) || 0);
            }
        });
        const gross = basic + allowance + otPay;
        return `<tr>
            <td class="p-2 font-bold text-slate-800 dark:text-slate-200">${escapeHtml(name)}</td>
            <td class="p-2 text-center">${records.length}</td>
            <td class="p-2 text-center">${daysPresent}</td>
            <td class="p-2 text-center">${otHours}</td>
            <td class="p-2 text-right font-bold text-emerald-600 dark:text-emerald-400">₱${gross.toLocaleString('en-US', {minimumFractionDigits:2})}</td>
        </tr>`;
    }).join('');

    document.getElementById('saveHistoryRestoreBtn').onclick = () => restoreSaveHistoryEntry(entry.id);
}

/** 📂 Ibinabalik ang datos ng isang naunang save PABALIK sa Attendance Monitoring
 *  editor (papalitan ang kasalukuyang laman — may kumpirmasyon muna). */
function restoreSaveHistoryEntry(id) {
    const entry = saveFileHistory.find(e => e.id === id);
    if (!entry) return;
    if (!confirm(`⚠️ Papalitan nito ang KASALUKUYANG datos sa Attendance Monitoring ng datos mula sa save na "${entry.label}".\n\nNais mo bang magpatuloy?`)) return;

    parsedEmployeesData = JSON.parse(JSON.stringify(entry.employees));
    currentActiveEmployee = Object.keys(parsedEmployeesData)[0] || "";
    refreshEditorUI();
    saveAppState();
    closeSaveHistoryModal();
    showToast(`📂 Na-load ang save: "${entry.label}"`);
}

function deleteSaveHistoryEntry(id) {
    const entry = saveFileHistory.find(e => e.id === id);
    if (!entry) return;
    if (!confirm(`Sigurado ka bang nais mong burahin ang save history na "${entry.label}"? Hindi na ito mababawi.`)) return;

    saveFileHistory = saveFileHistory.filter(e => e.id !== id);
    persistSaveFileHistory();
    renderSaveHistoryList();
    updateSaveHistoryBadge();
    showToast("🗑️ Nabura ang save history entry.");
}

/** Simpleng paraan para makagawa ng text na ligtas na i-inject bilang innerHTML. */
function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str === null || str === undefined ? '' : String(str);
    return div.innerHTML;
}

/** Maliit na di-nakakaabalang notification (di tulad ng blocking na alert()). */
function showToast(message, isError) {
    let container = document.getElementById('toastContainer');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toastContainer';
        container.className = 'fixed bottom-4 right-4 z-[999] flex flex-col gap-2 items-end pointer-events-none';
        document.body.appendChild(container);
    }
    const toast = document.createElement('div');
    toast.className = `pointer-events-auto max-w-xs px-4 py-3 rounded-xl shadow-2xl text-xs font-bold text-white ${isError ? 'bg-red-600' : 'bg-emerald-600'}`;
    toast.style.transition = 'opacity 0.4s ease, transform 0.4s ease';
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateX(8px)';
        setTimeout(() => toast.remove(), 400);
    }, 2800);
}

// ====================================================================
// 🌓 THEME CONTROL LAYER
// ====================================================================
function toggleTheme() {
    const html = document.documentElement;
    const btnIcon = document.getElementById('themeBtnIcon');
    const btnText = document.getElementById('themeBtnText');
    
    if (html.classList.contains('dark')) {
        html.classList.remove('dark');
        html.classList.add('light');
        btnIcon.innerText = "🌙";
        btnText.innerText = "LIGHT MODE";
        localStorage.setItem('theme', 'light');
        updateHighchartsTheme(false);
    } else {
        html.classList.remove('light');
        html.classList.add('dark');
        btnIcon.innerText = "☀️";
        btnText.innerText = "DARK MODE";
        localStorage.setItem('theme', 'dark');
        updateHighchartsTheme(true);
    }
}

function updateHighchartsTheme(isDark) {
    const color = isDark ? '#cbd5e1' : '#334155';
    if (window.Highcharts) {
        Highcharts.setOptions({
            legend: { itemStyle: { color: color } },
            xAxis: { labels: { style: { color: color } } },
            yAxis: { labels: { style: { color: color } } }
        });
    }
    if(globalData.length > 0 && fileType !== "attendance") {
        document.getElementById('searchBar').dispatchEvent(new Event('input'));
    }
}

// ====================================================================
// 🗺️ TAB NAVIGATION CONTROL INTERFACE
// ====================================================================
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('btnDashboard').addEventListener('click', function() {
        this.className = "w-full flex items-center space-x-2 py-2.5 px-4 rounded-xl bg-orange-600 text-white font-bold transition-all shadow-md shadow-orange-600/20 cursor-pointer";
        document.getElementById('btnAttendance').className = "w-full flex items-center space-x-2 py-2.5 px-4 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-bold transition-all hover:bg-orange-600 hover:text-white cursor-pointer";
        document.getElementById('dashboardView').classList.remove('hidden');
        document.getElementById('attendanceView').classList.add('hidden');
        saveAppState();
    });

    document.getElementById('btnAttendance').addEventListener('click', function() {
        this.className = "w-full flex items-center space-x-2 py-2.5 px-4 rounded-xl bg-orange-600 text-white font-bold transition-all shadow-md shadow-orange-600/20 cursor-pointer";
        document.getElementById('btnDashboard').className = "w-full flex items-center space-x-2 py-2.5 px-4 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-bold transition-all hover:bg-orange-600 hover:text-white cursor-pointer";
        document.getElementById('attendanceView').classList.remove('hidden');
        document.getElementById('dashboardView').classList.add('hidden');
        saveAppState();
    });

    const savedTheme = localStorage.getItem('theme') || 'light';
    const html = document.documentElement;
    if(savedTheme === 'dark') {
        html.classList.remove('light');
        html.classList.add('dark');
        document.getElementById('themeBtnIcon').innerText = "☀️";
        document.getElementById('themeBtnText').innerText = "LIGHT MODE";
        updateHighchartsTheme(true);
    } else {
        document.getElementById('themeBtnIcon').innerText = "🌙";
        document.getElementById('themeBtnText').innerText = "DARK MODE";
        updateHighchartsTheme(false);
    }
    
    const exportBtn = document.getElementById('btnExportDashboard');
    if (exportBtn) {
        exportBtn.addEventListener('click', handleDashboardExport);
    }

    const clearDataBtn = document.getElementById('clearDataBtn');
    if (clearDataBtn) {
        clearDataBtn.addEventListener('click', clearAppState);
    }


    // 📂 Live Excel Modal — minsan-lang-mag-upload (snapshot) na file input handler
    const liveExcelInput = document.getElementById('liveExcelFileInput');
    if (liveExcelInput) {
        liveExcelInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            _pendingLiveFileHandle = null; // hindi ito live-link — plain snapshot lang
            _liveExcelFile = file;
            const reader = new FileReader();
            reader.onload = function(evt) {
                const data = new Uint8Array(evt.target.result);
                _liveExcelWorkbook = XLSX.read(data, { type: 'array', cellDates: true });
                document.getElementById('liveExcelFileName').innerHTML =
                    '<span class="text-emerald-600 dark:text-emerald-400 font-black">✅ ' + file.name + '</span> <span class="text-slate-400 font-normal">(isang beses lang)</span>';
                renderSheetButtons(_liveExcelWorkbook.SheetNames);
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // 👤 Ipakita ang user info sa sidebar (name, role, avatar)
    initUserInfoBar();

    // 🗂️ I-load ang Save File History (mga naunang "🔄 Automatic Save")
    loadSaveFileHistory();
    updateSaveHistoryBadge();

    // 🔁 Subukan munang i-resume ang Google Sheets live link (kung naka-configure);
    // kung wala o di ma-access, babalik sa dating na-cache na data — para hindi
    // mawala ang laman ng website pag nag-reload.
    initLiveDataSource();
});

// ====================================================================
// 🛠️ DYNAMIC PARSING DATA UTILITIES
// ====================================================================
function parseCSVLine(text) {
    let insideQuote = false, entries = [''], index = 0;
    for (let i = 0; i < text.length; i++) {
        let char = text[i];
        if (char === '"') { insideQuote = !insideQuote; } 
        else if (char === ',' && !insideQuote) { entries[++index] = ''; } 
        else { entries[index] += char; }
    }
    return entries;
}

function cleanNumber(val) {
    if (!val) return 0;
    let clean = val.replace(/"/g, '').replace(/,/g, '').trim();
    return parseInt(clean) || 0;
}

// ====================================================================
// 📅 LATEST DATE CARD UPDATER + DATE PICKER SEARCH
// ====================================================================

// Global: all unique YYYY-MM-DD dates parsed from current data
let _allParsedDates = {};  // map: "YYYY-MM-DD" -> count of records

// Helper: parse any date string into a YYYY-MM-DD string, or null
function parseDateToISO(str) {
    if (!str || str === '-' || str === '' || str === 'N/A') return null;
    let s = String(str).trim();
    if (!s) return null;
    // Already YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    // ISO with time component: 2026-06-15T00:00:00.000Z
    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) return s.split('T')[0];
    // Excel serial number (pure 4-5 digit number, e.g. 46180 = 2026-06-15)
    // Dates in modern Excel are typically in the 40000-80000 range
    if (/^\d{4,5}$/.test(s)) {
        let serial = parseInt(s);
        // Convert Excel serial to Unix timestamp: (serial - 25569) days since Jan 1 1970
        let d = new Date((serial - 25569) * 86400 * 1000);
        if (!isNaN(d.getTime())) {
            let y = d.getUTCFullYear();
            if (y >= 1990 && y <= 2100) {
                return y + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' +
                    String(d.getUTCDate()).padStart(2, '0');
            }
        }
    }
    // Try native Date parse (handles ISO, RFC2822, many locale formats)
    let d = new Date(s);
    if (!isNaN(d.getTime()) && d.getFullYear() >= 1990 && d.getFullYear() <= 2100) {
        return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' +
            String(d.getDate()).padStart(2, '0');
    }
    // Try MM/DD/YYYY, M/D/YYYY, or DD/MM/YYYY (Filipino/European format)
    let slashParts = s.split('/');
    if (slashParts.length === 3) {
        let raw0 = slashParts[0].trim(), raw1 = slashParts[1].trim(), raw2 = slashParts[2].trim();
        let y = raw2.length === 2 ? '20' + raw2 : raw2;
        // MM/DD/YYYY attempt (default)
        let d1 = new Date(`${y}-${raw0.padStart(2,'0')}-${raw1.padStart(2,'0')}`);
        if (!isNaN(d1.getTime()) && d1.getFullYear() >= 1990 && d1.getFullYear() <= 2100) {
            return d1.getFullYear() + '-' + String(d1.getMonth() + 1).padStart(2,'0') + '-' +
                String(d1.getDate()).padStart(2,'0');
        }
        // DD/MM/YYYY fallback (when first part > 12, it must be the day)
        if (parseInt(raw0) > 12) {
            let d2 = new Date(`${y}-${raw1.padStart(2,'0')}-${raw0.padStart(2,'0')}`);
            if (!isNaN(d2.getTime()) && d2.getFullYear() >= 1990 && d2.getFullYear() <= 2100) {
                return d2.getFullYear() + '-' + String(d2.getMonth() + 1).padStart(2,'0') + '-' +
                    String(d2.getDate()).padStart(2,'0');
            }
        }
    }
    return null;
}

// Get ISO date string from a row depending on fileType
function getISODateFromRow(row) {
    // Direct field mappings for known file types
    if (fileType === 'raw' || fileType === 'crossdock') {
        let d = parseDateToISO(row.date);
        if (d) return d;
    }
    if (fileType === 'generic') {
        for (let h of genericHeaders) {
            if (h.toLowerCase().includes('date') || h.toLowerCase().includes('time')) {
                let d = parseDateToISO(row[h]);
                if (d) return d;
            }
        }
    }
    // Universal fallback: scan all string fields whose KEY NAME suggests a date.
    // This ensures ANY file type (progress, device, custom) works if it has a date column.
    const skipKeys = new Set(['_idx', 'qty', 'allocated', 'forecast', 'shipped',
                               'discrepancy', 'percentage', 'specRows', 'specCols',
                               'dailyRate', 'otHours', 'allowance']);
    const dateHints = ['date', 'time', 'dt', 'day', 'when', 'created', 'updated',
                       'ship', 'receiv', 'deliver', 'dispatch', 'sent', 'arrival'];
    for (let key of Object.keys(row)) {
        if (skipKeys.has(key)) continue;
        let keyLow = key.toLowerCase();
        if (!dateHints.some(hint => keyLow.includes(hint))) continue;
        let val = String(row[key] || '').trim();
        if (!val || val === '-' || val.length < 5) continue;
        let d = parseDateToISO(val);
        if (d && d >= '1990-01-01' && d <= '2100-12-31') return d;
    }
    return null;
}

function updateLatestDateCard(data) {
    // Build a fresh date map from ALL globalData (not just filtered)
    _allParsedDates = {};
    globalData.forEach(row => {
        let iso = getISODateFromRow(row);
        if (iso) _allParsedDates[iso] = (_allParsedDates[iso] || 0) + 1;
    });

    const allDates = Object.keys(_allParsedDates).sort();
    const picker = document.getElementById('datePicker');

    if (allDates.length === 0) {
        // No date column in this file type
        document.getElementById('latestDateLabel').innerText = '—';
        document.getElementById('latestDateCount').innerText = '—';
        picker.disabled = true;
        picker.value = '';
        document.getElementById('clearDateBtn').classList.add('hidden');
        return;
    }

    // Enable date picker — max is always TODAY so user can always pick the current date
    const todayISO = new Date().toISOString().split('T')[0];
    picker.disabled = false;
    picker.min = allDates[0];                 // earliest date in the file
    picker.max = todayISO;                    // always allow picking today

    // Default: always show today in the date picker
    if (!picker.value) {
        picker.value = todayISO;
    }

    // Display the currently selected date info
    const displayISO = picker.value || todayISO;
    const displayCount = _allParsedDates[displayISO] || 0;

    // Format: M/D/YYYY
    const parts = displayISO.split('-');
    const displayLabel = parseInt(parts[1]) + '/' + parseInt(parts[2]) + '/' + parts[0];
    document.getElementById('latestDateLabel').innerText = displayLabel;
    document.getElementById('latestDateCount').innerText = displayCount > 0
        ? displayCount.toLocaleString() + ' record/box' + (displayCount !== 1 ? 'es' : '')
        : '0 — Walang record sa date na ito';
}

// Called when user picks a date from the date picker
function searchByDate(isoDate) {
    if (!isoDate || !_allParsedDates) return;

    const count = _allParsedDates[isoDate] || 0;
    const parts = isoDate.split('-');
    const displayLabel = parseInt(parts[1]) + '/' + parseInt(parts[2]) + '/' + parts[0];

    document.getElementById('latestDateLabel').innerText = displayLabel;
    document.getElementById('latestDateCount').innerText = count > 0
        ? count.toLocaleString() + ' record/box' + (count !== 1 ? 'es' : '')
        : '0 — Walang record sa date na ito';

    // Show clear button
    document.getElementById('clearDateBtn').classList.remove('hidden');

    // Filter the table to show only rows matching this date
    if (count > 0) {
        const filtered = globalData.filter(row => getISODateFromRow(row) === isoDate);
        _skipChartRedraw = true;
        buildDashboard(filtered);
        _skipChartRedraw = false;
    } else {
        // No records — show empty table
        _skipChartRedraw = true;
        buildDashboard([]);
        _skipChartRedraw = false;
    }
}

// Clear the date filter and restore full data
function clearDateSearch() {
    const picker = document.getElementById('datePicker');
    picker.value = '';
    document.getElementById('clearDateBtn').classList.add('hidden');
    // Restore full table — updateLatestDateCard inside buildDashboard
    // will re-apply the correct default (today if data exists, else latest in file)
    buildDashboard(globalData);
}

// ====================================================================
// 📊 SUMMARY CARDS UPDATER (Pending, Done, Local, HighValue)
// Returns the computed totals so the pie chart uses the EXACT same numbers.
// ====================================================================
function updateSummaryCards(data) {
    let totalPending = 0, totalDone = 0, totalLocal = 0, totalHighValue = 0;

    if (fileType === "raw") {
        data.forEach(row => {
            let statusLow = row.status.toLowerCase();
            let isDone = statusLow.includes('done') || (statusLow.includes('ship') && !statusLow.includes('pending'));
            if (isDone) totalDone++; else totalPending++;
            if (row.type === "HighValue") totalHighValue++; else totalLocal++;
        });
    } else if (fileType === "progress") {
        data.forEach(row => {
            let isDone = row.percentage >= 100 || row.discrepancy <= 0;
            if (isDone) totalDone++; else totalPending++;
        });
    } else if (fileType === "crossdock") {
        data.forEach(row => {
            let isDone = row.statusReceived.toLowerCase().includes('done');
            if (isDone) totalDone++; else totalPending++;
        });
    } else if (fileType === "device") {
        data.forEach(row => {
            let statusLow = (row.status || '').toLowerCase();
            if (statusLow === 'done') totalDone++; else totalPending++;
        });
    } else if (fileType === "generic") {
        totalDone = 0; totalPending = data.length;
    } else if (fileType === "cancelled") {
        // done = rows with picker assigned; pending = rows without picker
        data.forEach(row => {
            if (row.picker && row.picker !== '') totalDone++; else totalPending++;
        });
    }

    document.getElementById('summaryPending').innerText = totalPending.toLocaleString();
    document.getElementById('summaryDone').innerText = totalDone.toLocaleString();
    document.getElementById('summaryLocal').innerText = (fileType === "raw") ? totalLocal.toLocaleString() : '—';
    document.getElementById('summaryHighValue').innerText = (fileType === "raw") ? totalHighValue.toLocaleString() : '—';

    // Return values so pie chart uses EXACTLY the same numbers as the cards
    return { totalPending, totalDone, totalLocal, totalHighValue };
}


// ====================================================================
// 🔢 SERIES NUMBER SUMMARY PANEL UPDATER
// Nagpapakita ng breakdown ng bawat unique Series Number at bilang ng
// MTB boxes nito. Lalabas lang ito kapag fileType === "raw".
// Tinatawag sa loob ng buildDashboard() pagkatapos maproseso ang data.
// ====================================================================
function updateSeriesPanel(batchMap, activeISO) {
    const panel    = document.getElementById('seriesNumberPanel');
    const countEl  = document.getElementById('seriesUniqueCount');
    const listEl   = document.getElementById('seriesNumberList');
    if (!panel || !countEl || !listEl) return;

    // Itago ang panel kapag hindi raw data
    if (fileType !== 'raw') {
        panel.classList.add('hidden');
        return;
    }
    panel.classList.remove('hidden');

    // Ipakita ang petsa ng Series Summary (today o napiling date)
    const subtitleEl = document.getElementById('seriesDateSubtitle');
    if (subtitleEl && activeISO) {
        const _parts    = activeISO.split('-');
        const _todayISO = new Date().toISOString().split('T')[0];
        const _isToday  = activeISO === _todayISO;
        const _label    = parseInt(_parts[1]) + '/' + parseInt(_parts[2]) + '/' + _parts[0];
        subtitleEl.textContent = _isToday
            ? 'Today (' + _label + ') — Latest'
            : 'Selected Date: ' + _label;
    }

    const entries     = Object.entries(batchMap);
    const uniqueCount = entries.length;
    countEl.innerText = uniqueCount.toLocaleString();

    // Calculate and display grand total MTB boxes
    const totalBoxes = entries.reduce((sum, [, c]) => sum + c, 0);
    const totalEl    = document.getElementById('seriesTotalBoxes');
    if (totalEl) totalEl.innerText = totalBoxes.toLocaleString();

    if (uniqueCount === 0) {
        listEl.innerHTML = '<p class="text-xs text-slate-400 dark:text-slate-500 italic">No Series Number found for this date.</p>';
        return;
    }

    // Sort: highest count first
    entries.sort((a, b) => b[1] - a[1]);

    // Rotating color palette for each series (light mode)
    const palettes = [
        { bg: 'background:#fff7ed;border-color:#fed7aa;color:#c2410c;' },
        { bg: 'background:#f0f9ff;border-color:#bae6fd;color:#0369a1;' },
        { bg: 'background:#f0fdf4;border-color:#bbf7d0;color:#15803d;' },
        { bg: 'background:#faf5ff;border-color:#e9d5ff;color:#7e22ce;' },
        { bg: 'background:#fff1f2;border-color:#fecdd3;color:#be123c;' },
        { bg: 'background:#fffbeb;border-color:#fde68a;color:#b45309;' },
        { bg: 'background:#f0fdfa;border-color:#99f6e4;color:#0f766e;' },
        { bg: 'background:#eef2ff;border-color:#c7d2fe;color:#4338ca;' },
    ];

    // Dark mode palette — same color families, deep dark bg with vivid border+text
    const darkPalettes = [
        { bg: 'background:#431407;border-color:#ea580c;color:#fdba74;' },   // orange
        { bg: 'background:#0c2040;border-color:#0284c7;color:#7dd3fc;' },   // blue
        { bg: 'background:#052e16;border-color:#16a34a;color:#86efac;' },   // green
        { bg: 'background:#2e1065;border-color:#9333ea;color:#d8b4fe;' },   // purple
        { bg: 'background:#4c0519;border-color:#e11d48;color:#fda4af;' },   // rose
        { bg: 'background:#451a03;border-color:#d97706;color:#fcd34d;' },   // amber
        { bg: 'background:#042f2e;border-color:#0d9488;color:#5eead4;' },   // teal
        { bg: 'background:#1e1b4b;border-color:#4f46e5;color:#a5b4fc;' },   // indigo
    ];

    const isDark = document.documentElement.classList.contains('dark');

    let html = '';
    entries.forEach(([seriesName, count], idx) => {
        const p       = isDark ? darkPalettes[idx % darkPalettes.length] : palettes[idx % palettes.length];
        const bgStyle = p.bg;
        const shortName = seriesName.length > 34 ? seriesName.slice(0, 32) + '…' : seriesName;

        // Vertical card: COUNT on top, NAME below
        html += `<div title="${seriesName}" style="${bgStyle}border:1px solid;border-radius:14px;padding:8px 14px;display:flex;flex-direction:column;align-items:center;gap:3px;font-size:12px;cursor:default;flex-shrink:0;min-width:110px;text-align:center;">
            <span style="font-family:monospace;font-weight:900;font-size:20px;line-height:1.1;">${count.toLocaleString()}<span style="font-size:10px;font-weight:700;margin-left:2px;opacity:0.75;">box</span></span>
            <span style="font-weight:600;font-size:11px;max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;opacity:0.85;">${shortName}</span>
        </div>`;
    });

    listEl.innerHTML = html;
}

function toggleType(index) {
    let row = globalData[index];
    row.type = (row.type === "Local") ? "HighValue" : "Local";
    _skipChartRedraw = true;
    let query = document.getElementById('searchBar').value;
    if (query) { document.getElementById('searchBar').dispatchEvent(new Event('input')); } else { buildDashboard(globalData); }
    _skipChartRedraw = false;
    refreshRecordDetailModalIfOpen(); // ✅ i-sync ang popup kung open ito habang na-toggle mula rito
    saveAppState();
}

function buildDashboard(data) {
    const tbody = document.getElementById('tableBody');
    const thead = document.getElementById('tableHeader');
    
    // I-reset ang series panel — ipakita lang kapag raw file ang aktibo
    const _spanel = document.getElementById('seriesNumberPanel');
    if (_spanel && fileType !== 'raw') _spanel.classList.add('hidden');

    if (data.length === 0) { 
        tbody.innerHTML = `<tr><td class="py-8 text-center text-slate-400">Walang natagpuang tugma.</td></tr>`; 
        return; 
    }
    
    let doneCount = 0, pendingCount = 0;
    let whMap = {}, batchMap = {};
    let sumForecast = 0, sumAllocated = 0, sumShipped = 0, sumDiscrepancy = 0;
    let htmlBuffer = [];
    
    const textClass = "text-slate-800 dark:text-slate-200";
    const fontMuted = "text-slate-500 dark:text-slate-400";
    
    if (fileType === "progress") {
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400"><th class="py-3.5 px-4">WW</th><th class="py-3.5 px-4">Month</th><th class="py-3.5 px-4">Warehouse</th><th class="py-3.5 px-4">MTO Batch Name</th><th class="py-3.5 px-4 text-right">Forecast</th><th class="py-3.5 px-4 text-right">Allocated</th><th class="py-3.5 px-4 text-right">Total Shipped</th><th class="py-3.5 px-4 text-right">Discrepancy</th><th class="py-3.5 px-4 text-center">Status</th><th class="py-3.5 px-4 text-center">Percentage</th></tr>`;
        
        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            sumForecast += row.forecast; sumAllocated += row.allocated; sumShipped += row.shipped; sumDiscrepancy += row.discrepancy;
            if (row.wh) whMap[row.wh] = (whMap[row.wh] || 0) + row.shipped;
            if (row.batch) batchMap[row.batch] = (batchMap[row.batch] || 0) + row.shipped;
            
            let isDone = row.percentage >= 100 || row.discrepancy <= 0;
            if (isDone) doneCount++; else pendingCount++;
            
            let origIndex = row._idx;
            let statusBadge = isDone ? 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition-all">Done</button>` : 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 hover:bg-amber-200 transition-all">Pending</button>`;
            
            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40"><td class="py-2.5 px-4 font-mono text-xs ${fontMuted}">${row.ww}</td><td class="py-2.5 px-4 font-bold ${textClass}">${row.month}</td><td class="py-2.5 px-4 font-black text-orange-600 dark:text-orange-400">${row.wh}</td><td class="py-2.5 px-4 font-medium ${textClass} truncate max-w-[150px]">${row.batch}</td><td class="py-2.5 px-4 text-right font-mono ${fontMuted}">${row.forecast.toLocaleString()}</td><td class="py-2.5 px-4 text-right font-mono text-blue-600 dark:text-blue-400">${row.allocated.toLocaleString()}</td><td class="py-2.5 px-4 text-right font-mono ${textClass} font-bold">${row.shipped.toLocaleString()}</td><td class="py-2.5 px-4 text-right font-mono text-red-500">${row.discrepancy.toLocaleString()}</td><td class="py-2.5 px-4 text-center">${statusBadge}</td><td class="py-2.5 px-4 text-center font-mono font-bold text-emerald-600 dark:text-emerald-400">${row.percentage}%</td></tr>`);
        }
        updateMetricsDisplay(sumForecast, "Total Forecast Volume", sumAllocated, "Total Allocated Qty", sumShipped, "Total Shipped (Done)", sumDiscrepancy, "Total Discrepancy Volume");
        generateHighchartsGraphs(doneCount, pendingCount, whMap, batchMap);

    } else if (fileType === "raw") {
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400"><th class="py-3.5 px-4">Box / MTB ID</th><th class="py-3.5 px-4 text-center">Type</th><th class="py-3.5 px-4">Series / Serial Number</th><th class="py-3.5 px-4">Destination Warehouse</th><th class="py-3.5 px-4 text-right">Qty</th><th class="py-3.5 px-4 text-center">Status</th><th class="py-3.5 px-4 text-center">Shipped Date</th><th class="py-3.5 px-4">Remarks</th></tr>`;
        
        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            let statusLow = row.status.toLowerCase();
            let isDone = statusLow.includes('done') || (statusLow.includes('ship') && !statusLow.includes('pending'));
            if (isDone) doneCount++; else pendingCount++;
            if (row.wh) whMap[row.wh] = (whMap[row.wh] || 0) + 1;
            if (row.series) batchMap[row.series] = (batchMap[row.series] || 0) + 1;
            
            let origIndex = row._idx;
            let statusBadge = isDone ? 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition-all">Done</button>` : 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 hover:bg-amber-200 transition-all">Pending</button>`;
            
            let typeColor = row.type === "HighValue" ? "bg-purple-100 text-purple-800 border-purple-300" : "bg-sky-100 text-sky-800 border-sky-300";
            let typeBadge = `<button onclick="toggleType(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-md text-[11px] font-black tracking-wide border transition-all ${typeColor}">${row.type}</button>`;
            
            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40"><td class="py-2.5 px-4 font-mono font-bold ${textClass}">${row.mtb}</td><td class="py-2.5 px-4 text-center">${typeBadge}</td><td class="py-2.5 px-4 ${fontMuted} text-xs font-mono">${row.series || ''}</td><td class="py-2.5 px-4 text-orange-600 dark:text-orange-400 font-bold">${row.wh || ''}</td><td class="py-2.5 px-4 text-right font-mono ${fontMuted}">${row.qty ? row.qty.toLocaleString() : '1'}</td><td class="py-2.5 px-4 text-center">${statusBadge}</td><td class="py-2.5 px-4 text-center ${fontMuted} font-mono text-xs">${row.date || ''}</td><td class="py-2.5 px-4 ${fontMuted} text-xs italic truncate max-w-[120px]" title="${row.remarks || ''}">${row.remarks || ''}</td></tr>`);
        }
        updateMetricsDisplay(data.length, "Total Loaded Boxes", doneCount, "Boxes Shipped (Done)", pendingCount, "Pending for Ship", "N/A", "Discrepancy (Raw)");
        generateHighchartsGraphs(doneCount, pendingCount, whMap, batchMap);
    } else if (fileType === "crossdock") {
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400"><th class="py-3.5 px-4">Date</th><th class="py-3.5 px-4">MTB ID</th><th class="py-3.5 px-4">SKU Name</th><th class="py-3.5 px-4 text-right">Req Qty</th><th class="py-3.5 px-4 text-right">Act Qty</th><th class="py-3.5 px-4">Pallet</th><th class="py-3.5 px-4 text-center">Checked By</th><th class="py-3.5 px-4 text-center">Status</th></tr>`;
        
        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            sumAllocated += row.qty; sumShipped += row.actualQty;
            let isReceived = row.statusReceived.toLowerCase().includes('done');
            if (isReceived) doneCount++; else pendingCount++;
            if (row.pallet) whMap[row.pallet] = (whMap[row.pallet] || 0) + row.actualQty;
            if (row.skuName) batchMap[row.skuName] = (batchMap[row.skuName] || 0) + row.actualQty;
            
            let origIndex = row._idx;
            let statusBadge = isReceived ? 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition-all">Done</button>` : 
                `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 hover:bg-amber-200 transition-all">Pending</button>`;
            
            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40"><td class="py-2.5 px-4 font-mono text-xs ${fontMuted}">${row.date}</td><td class="py-2.5 px-4 font-bold font-mono ${textClass}">${row.mtb}</td><td class="py-2.5 px-4 ${textClass} font-medium truncate max-w-[200px]" title="${row.skuName}">${row.skuName}</td><td class="py-2.5 px-4 text-right font-mono ${fontMuted}">${row.qty.toLocaleString()}</td><td class="py-2.5 px-4 text-right font-mono text-blue-600 dark:text-blue-400 font-bold">${row.actualQty.toLocaleString()}</td><td class="py-2.5 px-4 text-xs ${fontMuted} truncate max-w-[150px]">${row.pallet}</td><td class="py-2.5 px-2 text-center"><input type="text" value="${row.checkedBy}" oninput="updateCheckerName(${origIndex}, this.value)" class="w-full text-center font-bold bg-slate-100 dark:bg-slate-900 border border-slate-300 dark:border-slate-700 text-slate-800 dark:text-slate-100 rounded px-1.5 py-1 text-xs focus:bg-white focus:border-indigo-500 focus:outline-none" /></td><td class="py-2.5 px-4 text-center">${statusBadge}</td></tr>`);
        }
        updateMetricsDisplay(data.length, "Total Line Items", sumAllocated, "Total Ordered Qty", sumShipped, "Total Actual Box Qty", (sumAllocated - sumShipped), "Variance / Gap Qty");
        generateHighchartsGraphs(doneCount, pendingCount, whMap, batchMap);
        
    } else if (fileType === "device") {
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400">
            <th class="py-3.5 px-4">Device ID</th>
            <th class="py-3.5 px-4">Device Type</th>
            <th class="py-3.5 px-4 text-center">Specification (Rows)</th>
            <th class="py-3.5 px-4 text-center">Specification (Columns)</th>
            <th class="py-3.5 px-4 text-center">Status</th>
            <th class="py-3.5 px-4 text-center">Device Progress</th>
            <th class="py-3.5 px-4 text-center">Template ID</th>
        </tr>`;

        let progressMap = {}, typeMap = {};
        const isDarkMode = document.documentElement.classList.contains('dark');
        const inputStyle = isDarkMode
            ? 'background:#1e293b;color:#e2e8f0;border:1px solid #475569;'
            : 'background:#f8fafc;color:#1e293b;border:1px solid #cbd5e1;';

        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            let statusLow = row.status.toLowerCase();
            let isDone = statusLow === 'done';
            if (isDone) doneCount++; else pendingCount++;

            typeMap[row.deviceType] = (typeMap[row.deviceType] || 0) + 1;
            progressMap[row.progress] = (progressMap[row.progress] || 0) + 1;

            let origIndex = row._idx;

            // --- STATUS: clickable button cycling Normal → Done → Pending → Normal ---
            let statusColor, statusLabel;
            if (statusLow === 'normal') {
                statusColor = 'background:#e0f2fe;color:#0369a1;border:1px solid #7dd3fc;';
                statusLabel = 'Normal';
            } else if (statusLow === 'done') {
                statusColor = 'background:#d1fae5;color:#065f46;border:1px solid #6ee7b7;';
                statusLabel = 'Done';
            } else {
                statusColor = 'background:#fef3c7;color:#92400e;border:1px solid #fcd34d;';
                statusLabel = row.status || 'Pending';
            }
            let statusBadge = `<button onclick="cycleDeviceStatus(${origIndex})" style="${statusColor}padding:2px 10px;border-radius:9999px;font-size:11px;font-weight:700;cursor:pointer;">${statusLabel}</button>`;

            // --- SPEC ROWS: editable number input ---
            let specRowInput = `<input type="number" min="0" value="${row.specRows !== undefined ? row.specRows : 0}" onchange="updateDeviceField(${origIndex},'specRows',this.value)" style="${inputStyle}width:70px;border-radius:6px;padding:2px 6px;font-size:12px;font-family:monospace;text-align:center;" />`;

            // --- SPEC COLS: editable number input ---
            let specColInput = `<input type="number" min="0" value="${row.specCols !== undefined ? row.specCols : 0}" onchange="updateDeviceField(${origIndex},'specCols',this.value)" style="${inputStyle}width:70px;border-radius:6px;padding:2px 6px;font-size:12px;font-family:monospace;text-align:center;" />`;

            // --- DEVICE PROGRESS: editable text input (type to change) ---
            let progressInput = `<input type="text" value="${row.progress || ''}" onchange="updateDeviceField(${origIndex},'progress',this.value)" style="${inputStyle}width:140px;border-radius:6px;padding:2px 8px;font-size:12px;font-weight:600;" placeholder="Type progress..." />`;

            // --- TEMPLATE ID: show exactly as loaded from Excel/CSV ---
            let tplVal = (row.templateId && row.templateId !== '-' && row.templateId.trim() !== '') ? row.templateId : '—';
            let tplDisplay = tplVal !== '—'
                ? `<span style="background:#ede9fe;color:#5b21b6;border:1px solid #c4b5fd;border-radius:6px;padding:2px 8px;font-size:11px;font-family:monospace;font-weight:700;">${tplVal}</span>`
                : `<span style="color:#94a3b8;font-size:12px;">—</span>`;

            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40">
                <td class="py-2.5 px-4 font-mono font-bold ${textClass}">${row.deviceId}</td>
                <td class="py-2.5 px-4 font-medium ${textClass}">${row.deviceType}</td>
                <td class="py-2 px-4 text-center">${specRowInput}</td>
                <td class="py-2 px-4 text-center">${specColInput}</td>
                <td class="py-2.5 px-4 text-center">${statusBadge}</td>
                <td class="py-2 px-4 text-center">${progressInput}</td>
                <td class="py-2.5 px-4 text-center">${tplDisplay}</td>
            </tr>`);
        }

        updateMetricsDisplay(
            data.length,   "Total Devices",
            doneCount,     "Completed Devices",
            pendingCount,  "Devices In Progress",
            Object.keys(typeMap).length, "Device Types"
        );
        generateHighchartsGraphs(doneCount, pendingCount, typeMap, progressMap);

    } else if (fileType === "generic") {
        let headerRowHtml = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400">`;
        genericHeaders.forEach(h => {
            headerRowHtml += `<th class="py-3.5 px-4 capitalize">${h}</th>`;
        });
        headerRowHtml += `</tr>`;
        thead.innerHTML = headerRowHtml;

        data.forEach(row => {
            let rowHtml = `<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40">`;
            genericHeaders.forEach(h => {
                rowHtml += `<td class="py-2.5 px-4 text-xs ${textClass}">${row[h] || '-'}</td>`;
            });
            rowHtml += `</tr>`;
            htmlBuffer.push(rowHtml);
        });
        
        updateMetricsDisplay(data.length, "Total Records", genericHeaders.length, "Total Columns", 0, "Custom File Active", 0, "No Discrepancy Calc");
        clearHighchartsGraphs();


    } else if (fileType === "cancelled") {
        // ✅ MTO Monitoring / Cancelled — all 11 columns with correct headers
        thead.innerHTML = `<tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400 text-xs uppercase tracking-wide">
            <th class="py-3.5 px-3">WHs</th>
            <th class="py-3.5 px-3">MT Order</th>
            <th class="py-3.5 px-3">Picking ID</th>
            <th class="py-3.5 px-3">SKU</th>
            <th class="py-3.5 px-3">SKU Name</th>
            <th class="py-3.5 px-3 text-right">Qty</th>
            <th class="py-3.5 px-3 text-right">Picked Qty</th>
            <th class="py-3.5 px-3 text-right">Checked Qty</th>
            <th class="py-3.5 px-3 text-right">Lack Item</th>
            <th class="py-3.5 px-3">Reason</th>
            <th class="py-3.5 px-3">Picker</th>
        </tr>`;

        let whsMap2 = {}, reasonMap2 = {};
        let sumQty2 = 0, sumLack2 = 0;

        for (let i = 0; i < data.length; i++) {
            let row = data[i];
            sumQty2  += row.qty || 0;
            sumLack2 += row.lackItem || 0;

            if (row.whs)    whsMap2[row.whs]       = (whsMap2[row.whs]       || 0) + 1;
            if (row.reason) reasonMap2[row.reason]  = (reasonMap2[row.reason] || 0) + 1;

            let hasPicker  = row.picker  && row.picker  !== '';
            if (hasPicker) doneCount++; else pendingCount++;

            // Reason badge color mapping
            const reasonColors = {
                'NO OTHER LOCATION'  : 'bg-orange-100 text-orange-800 border-orange-300',
                'EXPIRY ITEMS'       : 'bg-red-100    text-red-800    border-red-300',
                'INVALID PICKING ID' : 'bg-purple-100 text-purple-800 border-purple-300',
                'LIQUOR'             : 'bg-sky-100    text-sky-800    border-sky-300',
                'INCOMPLETE CHECKING': 'bg-yellow-100 text-yellow-800 border-yellow-300',
                'SMALL ITEMS'        : 'bg-teal-100   text-teal-800   border-teal-300',
                'DAMAGED ITEMS'      : 'bg-rose-100   text-rose-800   border-rose-300',
                'LACKING ITEM/LOST'  : 'bg-pink-100   text-pink-800   border-pink-300'
            };
            let reasonKey   = (row.reason || '').toUpperCase().trim();
            let reasonClass = reasonColors[reasonKey] || 'bg-slate-100 text-slate-700 border-slate-300';
            let reasonBadge = row.reason
                ? `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold border ${reasonClass} whitespace-nowrap">${row.reason}</span>`
                : `<span class="${fontMuted} text-xs">—</span>`;

            let pickerBadge = hasPicker
                ? `<span class="px-2 py-0.5 rounded-md text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">${row.picker}</span>`
                : `<span class="${fontMuted} text-xs italic">—</span>`;

            let lackDisplay = row.lackItem > 0
                ? `<span class="font-mono font-bold text-red-500">${row.lackItem.toLocaleString()}</span>`
                : `<span class="${fontMuted} text-xs">—</span>`;

            htmlBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40">
                <td class="py-2 px-3 font-black text-orange-600 dark:text-orange-400 text-xs">${row.whs || '—'}</td>
                <td class="py-2 px-3 font-mono font-bold ${textClass} text-xs">${row.mtOrder}</td>
                <td class="py-2 px-3 font-mono ${fontMuted} text-xs">${row.pickingId || '—'}</td>
                <td class="py-2 px-3 ${fontMuted} text-xs font-mono truncate max-w-[100px]" title="${row.sku}">${row.sku || '—'}</td>
                <td class="py-2 px-3 ${textClass} text-xs truncate max-w-[180px]" title="${row.skuName}">${row.skuName || '—'}</td>
                <td class="py-2 px-3 text-right font-mono ${fontMuted} text-xs">${row.qty > 0 ? row.qty.toLocaleString() : '—'}</td>
                <td class="py-2 px-3 text-right font-mono text-blue-600 dark:text-blue-400 text-xs">${row.pickedQty > 0 ? row.pickedQty.toLocaleString() : '—'}</td>
                <td class="py-2 px-3 text-right font-mono text-indigo-600 dark:text-indigo-400 text-xs">${row.checkedQty > 0 ? row.checkedQty.toLocaleString() : '—'}</td>
                <td class="py-2 px-3 text-right">${lackDisplay}</td>
                <td class="py-2 px-3">${reasonBadge}</td>
                <td class="py-2 px-3 text-center">${pickerBadge}</td>
            </tr>`);
        }

        let uniqueReasons = Object.keys(reasonMap2).length;
        updateMetricsDisplay(data.length, "Total Cancelled Lines", sumQty2, "Total Qty", sumLack2, "Total Lack Items", uniqueReasons, "Cancellation Reasons");
        generateHighchartsGraphs(doneCount, pendingCount, whsMap2, reasonMap2);
    }
    
    tbody.innerHTML = htmlBuffer.join('');
    document.getElementById('rowCount').innerText = `May kabuuang ${data.length} na talaan ang aktibo sa system.`;

    // ✅ Update the 4 summary cards AND capture values for pie chart sync
    const cardTotals = updateSummaryCards(data);

    // ✅ Update the latest date card (this sets datePicker.value correctly)
    updateLatestDateCard(data);

    // 🔢 Series Number Summary — update AFTER updateLatestDateCard so datePicker is already set
    if (fileType === 'raw') {
        const _dp        = document.getElementById('datePicker');
        const _todayISO  = new Date().toISOString().split('T')[0];
        const _activeISO = (_dp && _dp.value) ? _dp.value : _todayISO;
        const _seriesMap = {};
        globalData.forEach(row => {
            if (getISODateFromRow(row) === _activeISO && row.series)
                _seriesMap[row.series] = (_seriesMap[row.series] || 0) + 1;
        });
        updateSeriesPanel(_seriesMap, _activeISO);
    }

    // ✅ Re-draw pie chart so it ALWAYS matches the cards exactly
    if (fileType !== "generic") {
        refreshPieChartFromCards(cardTotals);
    }
}

function updateCheckerName(index, bagongPangalan) { globalData[index].checkedBy = bagongPangalan.toUpperCase(); refreshRecordDetailModalIfOpen(); saveAppStateDebounced(); }

function toggleStatus(index) {
    let row = globalData[index];
    if (fileType === "progress") {
        if (row.percentage < 100 || row.discrepancy > 0) {
            row.shipped = row.allocated; row.discrepancy = 0; row.percentage = 100;
        } else {
            row.shipped = 0; row.discrepancy = row.allocated; row.percentage = 0;
        }
    } else if (fileType === "raw") {
        let statusLow = row.status.toLowerCase();
        if (statusLow.includes('done') || (statusLow.includes('ship') && !statusLow.includes('pending'))) {
            row.status = "Pending";
        } else {
            row.status = "Done"; let ngayon = new Date(); row.date = ngayon.toISOString().split('T')[0]; 
        }
    } else if (fileType === "crossdock") {
        row.statusReceived = row.statusReceived.toLowerCase() === 'done' ? "Pending" : "Done";
    }
    // Skip chart redraw on status toggle — only rebuild table + cards
    _skipChartRedraw = true;
    let query = document.getElementById('searchBar').value;
    if (query) { document.getElementById('searchBar').dispatchEvent(new Event('input')); } else { buildDashboard(globalData); }
    _skipChartRedraw = false;
    refreshRecordDetailModalIfOpen(); // ✅ i-sync ang popup kung open ito habang na-toggle mula rito
    saveAppState();
}
// ====================================================================
// 🖊️ DEVICE TABLE INTERACTIVE CONTROLS
// ====================================================================
// Cycle device Status: Normal → Done → Pending → Normal
function cycleDeviceStatus(index) {
    let row = globalData[index];
    let current = (row.status || 'Normal').toLowerCase();
    if (current === 'normal') {
        row.status = 'Done';
    } else if (current === 'done') {
        row.status = 'Pending';
    } else {
        row.status = 'Normal';
    }
    _skipChartRedraw = true;
    let query = document.getElementById('searchBar').value;
    if (query) { document.getElementById('searchBar').dispatchEvent(new Event('input')); } else { buildDashboard(globalData); }
    _skipChartRedraw = false;
    refreshRecordDetailModalIfOpen(); // ✅ i-sync ang popup kung open ito habang na-toggle mula rito
    saveAppState();
}

// Update any device field (specRows, specCols, progress) without full re-render
function updateDeviceField(index, field, value) {
    globalData[index][field] = value;
    refreshRecordDetailModalIfOpen();
    saveAppStateDebounced();
}

function updateMetricsDisplay(m1v, m1l, m2v, m2l, m3v, m3l, m4v, m4l) {
    document.getElementById('metric1Value').innerText = m1v.toLocaleString(); document.getElementById('metric1Label').innerText = m1l;
    document.getElementById('metric2Value').innerText = m2v.toLocaleString(); document.getElementById('metric2Label').innerText = m2l;
    document.getElementById('metric3Value').innerText = m3v.toLocaleString(); document.getElementById('metric3Label').innerText = m3l;
    document.getElementById('metric4Value').innerText = m4v.toLocaleString(); document.getElementById('metric4Label').innerText = m4l;
}

function clearHighchartsGraphs() {
    ['chart1Container', 'chart2Container', 'chart3Container'].forEach(id => {
        const el = document.getElementById(id);
        if(el) el.innerHTML = `<div class="h-full flex items-center justify-center text-xs text-slate-400 italic">Walang visual graph para sa custom list file.</div>`;
    });
}


// ====================================================================
// 🥧 PIE CHART SYNC — always mirrors the 4 summary cards exactly
// ====================================================================
function refreshPieChartFromCards(cardTotals) {
    if (_skipChartRedraw) return; // perf: skip on status toggle
    const isDark = document.documentElement.classList.contains('dark');
    const lblColor = isDark ? '#94a3b8' : '#475569';

    let pieData = [];

    if (fileType === "raw") {
        // For Raw files: show Done/Pending AND Local/HighValue in one pie
        pieData = [
            { name: '✅ Done',        y: cardTotals.totalDone,       color: '#10b981' },
            { name: '⏳ Pending',     y: cardTotals.totalPending,    color: '#f59e0b' },
            { name: '📦 Local',       y: cardTotals.totalLocal,      color: '#38bdf8' },
            { name: '💎 High Value',  y: cardTotals.totalHighValue,  color: '#a855f7' }
        ].filter(d => d.y > 0);
    } else {
        // For all other file types: show Done vs Pending only
        pieData = [
            { name: '✅ Done / Received', y: cardTotals.totalDone,    color: '#10b981' },
            { name: '⏳ Pending',         y: cardTotals.totalPending, color: '#f59e0b' }
        ].filter(d => d.y > 0);
    }

    Highcharts.chart('chart1Container', {
        chart: { type: 'pie', options3d: { enabled: true, alpha: 45, beta: 0 }, backgroundColor: 'transparent' },
        title: { text: '📊 Status Overview', style: { fontSize: '12px', fontWeight: 'bold', color: lblColor } },
        plotOptions: {
            pie: {
                depth: 35,
                allowPointSelect: true,
                cursor: 'pointer',
                dataLabels: {
                    enabled: true,
                    format: '<b>{point.name}</b><br>{point.y} ({point.percentage:.0f}%)',
                    style: { fontSize: '9px', color: lblColor, fontWeight: 'bold' }
                },
                showInLegend: true
            }
        },
        legend: { itemStyle: { fontSize: '10px', color: lblColor } },
        credits: { enabled: false },
        series: [{ name: 'Count', data: pieData }]
    });
}

// ====================================================================
// 📊 HIGHCHARTS VISUALIZATION ENGINE
// ====================================================================
function generateHighchartsGraphs(doneCount, pendingCount, whMap, batchMap) {
    if (_skipChartRedraw) return; // perf: skip on status toggle
    const isDark = document.documentElement.classList.contains('dark');
    const lblColor = isDark ? '#94a3b8' : '#475569';

    // chart1 is now handled by refreshPieChartFromCards() to keep it in sync with the cards

    let whData = Object.keys(whMap).map(key => ({ name: key, y: whMap[key] })).slice(0, 10);
    Highcharts.chart('chart2Container', {
        chart: { type: 'pie', options3d: { enabled: true, alpha: 30, beta: 0 }, backgroundColor: 'transparent' },
        title: {
            text: fileType === "crossdock" ? '📊 Pallet Load Mix' : fileType === "cancelled" ? '🏬 WHs Distribution' : '🚚 Warehouse Mix',
            style: { fontSize: '13px', fontWeight: 'bold', color: lblColor }
        },
        plotOptions: {
            pie: {
                innerSize: '45%',
                depth: 25,
                allowPointSelect: true,
                cursor: 'pointer',
                dataLabels: { enabled: false },
                showInLegend: true
            }
        },
        legend: {
            enabled: true,
            itemStyle: { fontSize: '9px', color: lblColor, fontWeight: '600' },
            layout: 'vertical',
            align: 'right',
            verticalAlign: 'middle',
            maxHeight: 160
        },
        credits: { enabled: false },
        series: [{ name: 'Share Volume', data: whData }]
    });

    let sortedBatches = Object.keys(batchMap).sort((a,b) => batchMap[b] - batchMap[a]).slice(0, 7);
    let batchValues = sortedBatches.map(b => batchMap[b]);
    Highcharts.chart('chart3Container', {
        chart: { type: 'column', options3d: { enabled: true, alpha: 15, beta: 15, depth: 50, viewDistance: 25 }, backgroundColor: 'transparent' },
        title: { text: null }, 
        xAxis: { categories: sortedBatches, labels: { style: { fontSize: '9px', color: lblColor } } },
        yAxis: { title: { text: null }, labels: { style: { color: lblColor } } },
        plotOptions: { column: { depth: 25, color: '#6366f1' } }, 
        legend: { enabled: false }, 
        credits: { enabled: false }, 
        series: [{ name: 'Volume', data: batchValues }]
    });
}
// ====================================================================
// 🔄 CORE DATA PROCESSOR — reusable parser for main uploader + Live Excel modal
// ====================================================================
function processFileContents(contents, fileName, isFreshLoad = true) {
    // ✅ Bagong (user-initiated) na pag-load = pwede ulit mag-popup ang auto-detect
    // kahit pareho ang laman ng search bar. Sa mga background/live-sync na poll
    // (isFreshLoad = false), hindi ito rini-reset para hindi paulit-ulit bumalik
    // ang popup kada ilang segundo habang naka-live sync.
    if (isFreshLoad) _lastAutoDetectQuery = null;
    lines = contents.split(/\r?\n/).filter(l => l.trim() !== "");
    if(lines.length === 0) {
        // ✅ FIX: Kapag live Google Sheets mode at may existing data na, huwag i-clear.
        // Posible na ang Google ay nag-return ng blangko o HTML redirect sa poll cycle na ito.
        if (_liveSheetsUrl && globalData.length > 0) return;
        // Ang file ay ganap na walang laman — ipakita blangkong dashboard
        globalData = [];
        document.getElementById('tableTitle').innerText = "📁 " + fileName;
        document.getElementById('searchBar').disabled = false;
        document.getElementById('searchBar').placeholder = "Mag-type rito para mag-filter ng data...";
        document.getElementById('btnDashboard').click(); buildDashboard([]);
        saveAppState();
        return;
    }

    let isAttendanceFile = false;
    for (let i = 0; i < Math.min(lines.length, 5); i++) {
        let lowLine = lines[i].toLowerCase();
        if (lowLine.includes('attendance') || lowLine.includes('number of ot') || lowLine.includes('daily rate')) {
            isAttendanceFile = true; break;
        }
    }

    if (isAttendanceFile) {
        fileType = "attendance";
        parseAttendanceCSV(contents); 
        document.getElementById('btnAttendance').click(); 
        saveAppState();
        return;
    }
    
    let startRow = 0;
    let foundMatch = false;
    for (let i = 0; i < lines.length; i++) {
        let lowerLine = lines[i].toLowerCase();
        if (lowerLine.includes('mto batch') || lowerLine.includes('box id') || lowerLine.includes('mtb') || lowerLine.includes('sku name') || lowerLine.includes('type') || lowerLine.includes('series') || lowerLine.includes('destination') || lowerLine.includes('high value') || lowerLine.includes('device id') || lowerLine.includes('device type') || lowerLine.includes('device progress')) {
            startRow = i; 
            foundMatch = true;
            break;
        }
    }
    
    let headers = parseCSVLine(lines[startRow]).map(h => h.trim());
    let lowerHeaders = headers.map(h => h.toLowerCase());
    let loadedRows = [];
    
    if (foundMatch && lowerHeaders.some(h => h.includes('device id') || h.includes('device type') || h.includes('device progress'))) {
        fileType = "device"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let deviceIdIdx    = lowerHeaders.findIndex(h => h.includes('device id'));
        let deviceTypeIdx  = lowerHeaders.findIndex(h => h.includes('device type'));
        let specRowsIdx    = lowerHeaders.findIndex(h => h.includes('specification') && h.includes('row'));
        let specColsIdx    = lowerHeaders.findIndex(h => h.includes('specification') && h.includes('col'));
        let statusIdx      = lowerHeaders.findIndex(h => h.includes('status'));
        let progressIdx    = lowerHeaders.findIndex(h => h.includes('device progress') || h.includes('progress'));
        let templateIdIdx  = lowerHeaders.findIndex(h => h.includes('template id') || h.includes('template'));

        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim());
            if (!columns[deviceIdIdx] || columns[deviceIdIdx].trim() === "") continue;
            loadedRows.push({
                deviceId:   columns[deviceIdIdx]   ? columns[deviceIdIdx].trim()   : '-',
                deviceType: columns[deviceTypeIdx] ? columns[deviceTypeIdx].trim() : '-',
                specRows:   columns[specRowsIdx]   ? columns[specRowsIdx].trim()   : '0',
                specCols:   columns[specColsIdx]   ? columns[specColsIdx].trim()   : '0',
                status:     columns[statusIdx]     ? columns[statusIdx].trim()     : 'Normal',
                progress:   columns[progressIdx]   ? columns[progressIdx].trim()   : '-',
                templateId: columns[templateIdIdx] ? columns[templateIdIdx].trim() : '-'
            });
        }
    } else if (foundMatch && lowerHeaders.some(h => h.includes('sku name')) && lowerHeaders.some(h => h.includes('checked by'))) {
        fileType = "crossdock"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let dateIdx = lowerHeaders.findIndex(h => h.includes('date')), skuNameIdx = lowerHeaders.findIndex(h => h.includes('sku name'));
        let qtyIdx = lowerHeaders.findIndex(h => h === 'qty' || h.includes('req qty')), mtbIdx = lowerHeaders.findIndex(h => h.includes('mtb'));
        let actQtyIdx = lowerHeaders.findIndex(h => h.includes('actual qty') || h.includes('actual box')), palletIdx = lowerHeaders.findIndex(h => h.includes('pallet')), checkerIdx = lowerHeaders.findIndex(h => h.includes('checked by')), statusRecIdx = lowerHeaders.findIndex(h => h.includes('status if received'));
        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim()); if (!columns[mtbIdx]) continue;
            loadedRows.push({ date: columns[dateIdx] ? columns[dateIdx].trim() : '', skuName: columns[skuNameIdx] ? columns[skuNameIdx].trim() : '', qty: cleanNumber(columns[qtyIdx]), mtb: columns[mtbIdx].trim().replace(/"/g, ''), actualQty: cleanNumber(columns[actQtyIdx]), pallet: columns[palletIdx] ? columns[palletIdx].trim() : '', checkedBy: columns[checkerIdx] && columns[checkerIdx].trim() !== "" ? columns[checkerIdx].trim().toUpperCase() : '', statusReceived: columns[statusRecIdx] ? columns[statusRecIdx].trim() : 'Pending' });
        }
    } else if (foundMatch && lowerHeaders.some(h => h.includes('forecast') || h.includes('mto batch'))) {
        fileType = "progress"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let whIdx = lowerHeaders.findIndex(h => h.includes('wh') || h.includes('warehouse')), wwIdx = lowerHeaders.findIndex(h => h.includes('work week') || h.includes('ww')), monthIdx = lowerHeaders.findIndex(h => h.includes('month')), batchIdx = lowerHeaders.findIndex(h => h.includes('mto batch') || h.includes('batch')), forecastIdx = lowerHeaders.findIndex(h => h.includes('forecast')), allocatedIdx = lowerHeaders.findIndex(h => h.includes('allocated')), shippedIdx = lowerHeaders.findIndex(h => h.includes('total shipped') || h.includes('shipped')), discIdx = lowerHeaders.findIndex(h => h.includes('discrepancy')), pctIdx = lowerHeaders.findIndex(h => h.includes('percentage') || h.includes('rate'));
        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim()); if (!columns[batchIdx] || columns[batchIdx].toLowerCase().includes('overall')) continue;
            loadedRows.push({ wh: columns[whIdx] ? columns[whIdx].trim().toUpperCase() : '', ww: columns[wwIdx] ? columns[wwIdx].trim() : '-', month: columns[monthIdx] ? columns[monthIdx].trim() : '-', batch: columns[batchIdx].trim().replace(/"/g, ''), forecast: cleanNumber(columns[forecastIdx]), allocated: cleanNumber(columns[allocatedIdx]), shipped: cleanNumber(columns[shippedIdx]), discrepancy: cleanNumber(columns[discIdx]), percentage: parseInt(columns[pctIdx] ? columns[pctIdx].replace('%', '') : '0') || 0 });
        }
    } else if (foundMatch && lowerHeaders.some(h => h.includes('lack item') || h.includes('lack') || h.includes('picking id'))) {
        // ✅ NEW: MTO Monitoring / Cancelled file type
        fileType = "cancelled"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let whsIdxC    = lowerHeaders.findIndex(h => h === 'whs' || h.includes('wh'));
        let mtOrderIdx = lowerHeaders.findIndex(h => h === 'mt order' || h.includes('mt order'));
        let pickIdIdx  = lowerHeaders.findIndex(h => h.includes('picking id') || h.includes('picking'));
        let skuIdxC    = lowerHeaders.findIndex(h => h === 'sku');
        let skuNmIdx   = lowerHeaders.findIndex(h => h === 'sku name' || (h.includes('sku') && h.includes('name')));
        let qty2Idx    = lowerHeaders.findIndex(h => h === 'qty');
        let pickedIdx  = lowerHeaders.findIndex(h => h.includes('picked qty') || h.includes('picked'));
        let checkedIdx = lowerHeaders.findIndex(h => h.includes('checked qty') || (h.includes('checked') && !h.includes('checked by')));
        let lackIdx    = lowerHeaders.findIndex(h => h.includes('lack item') || h.includes('lack'));
        let reasonIdx  = lowerHeaders.findIndex(h => h.includes('reason'));
        let pickerIdx  = lowerHeaders.findIndex(h => h.includes('picker'));

        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim());
            if (!columns[mtOrderIdx] || columns[mtOrderIdx].trim() === '') continue;
            loadedRows.push({
                whs:        columns[whsIdxC]    ? columns[whsIdxC].trim().toUpperCase()          : '',
                mtOrder:    columns[mtOrderIdx] ? columns[mtOrderIdx].trim()                     : '',
                pickingId:  columns[pickIdIdx]  ? columns[pickIdIdx].trim()                      : '',
                sku:        columns[skuIdxC]    ? columns[skuIdxC].trim().replace(/"/g,'')        : '',
                skuName:    columns[skuNmIdx]   ? columns[skuNmIdx].trim().replace(/"/g,'')       : '',
                qty:        cleanNumber(columns[qty2Idx]),
                pickedQty:  cleanNumber(columns[pickedIdx]),
                checkedQty: cleanNumber(columns[checkedIdx]),
                lackItem:   cleanNumber(columns[lackIdx]),
                reason:     columns[reasonIdx]  ? columns[reasonIdx].trim()                      : '',
                picker:     columns[pickerIdx]  ? columns[pickerIdx].trim().toUpperCase()         : ''
            });
        }
    } else if (foundMatch) {
        fileType = "raw"; document.getElementById('tableTitle').innerText = "📁 " + fileName;
        let typeIdx = lowerHeaders.findIndex(h => h === 'type' || h.includes('class') || h.includes('high value'));
        // FIX: (1) Detect headers na actual MTB ID ang value (hal. "MTB954023")
        //      (2) Ibukod ang "Box ID Status (Shipped)" — naglalaman ng "box id" pero status column ito
        let mtbIdx = lowerHeaders.findIndex(h =>
            h === 'mtb' ||
            /^mtb\d+$/i.test(h) ||                          // header na parang "MTB954023"
            h.includes('mtb id') ||
            (h.includes('box id') && !h.includes('status')) || // "box id" pero HINDI "box id status"
            h === 'box'
        );
        let seriesIdx = lowerHeaders.findIndex(h => h.includes('series') || h.includes('pallet') || h.includes('serial') || h.includes('number'));
        let whIdx = lowerHeaders.findIndex(h => h.includes('wh') || h.includes('destination') || h.includes('warehouse') || h.includes('dest'));
        let statusIdx = lowerHeaders.findIndex(h => h.includes('status') || h.includes('shipped)'));
        let dateIdx = lowerHeaders.findIndex(h => h.includes('date') || h.includes('shipped date') || h.includes('time'));
        let qtyIdx = lowerHeaders.findIndex(h => h === 'qty' || h.includes('quantity') || h.includes('pcs'));
        let remarksIdx = lowerHeaders.findIndex(h => h.includes('remark') || h.includes('note'));
        
        if (mtbIdx === -1) mtbIdx = lowerHeaders.findIndex(h => h.includes('mtb')) !== -1 ? lowerHeaders.findIndex(h => h.includes('mtb')) : 1;
        if (seriesIdx === -1) seriesIdx = lowerHeaders.findIndex(h => h.includes('series')) !== -1 ? lowerHeaders.findIndex(h => h.includes('series')) : 2;
        if (whIdx === -1) whIdx = lowerHeaders.findIndex(h => h.includes('wh') || h.includes('dest')) !== -1 ? lowerHeaders.findIndex(h => h.includes('wh') || h.includes('dest')) : 3;
        if (statusIdx === -1) statusIdx = lowerHeaders.findIndex(h => h.includes('status')) !== -1 ? lowerHeaders.findIndex(h => h.includes('status')) : 4;
        
        for (let i = startRow + 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim()); if (!columns[mtbIdx] || columns[mtbIdx].trim() === "" || columns[mtbIdx].toLowerCase().includes('type')) continue;
            
            let rawTypeValue = typeIdx !== -1 && columns[typeIdx] ? columns[typeIdx].trim().toUpperCase() : '';
            let rawType = "Local";
            if (rawTypeValue.includes("HIGH VALUE")) {
                rawType = "HighValue";
            }
            
            loadedRows.push({ type: rawType, mtb: columns[mtbIdx].trim().replace(/"/g, ''), series: columns[seriesIdx] ? columns[seriesIdx].trim().replace(/"/g, '') : '', wh: columns[whIdx] ? columns[whIdx].trim().toUpperCase() : '', qty: qtyIdx !== -1 && columns[qtyIdx] ? cleanNumber(columns[qtyIdx]) : 1, status: columns[statusIdx] ? columns[statusIdx].trim() : 'Pending', date: dateIdx !== -1 && columns[dateIdx] ? columns[dateIdx].trim() : '', remarks: remarksIdx !== -1 && columns[remarksIdx] ? columns[remarksIdx].trim() : '' });
        }
    } else {
        fileType = "generic";
        document.getElementById('tableTitle').innerText = "📁 " + fileName;
        headers = parseCSVLine(lines[0]).map(h => h.trim().replace(/"/g, ''));
        genericHeaders = headers;
        
        for (let i = 1; i < lines.length; i++) {
            let columns = parseCSVLine(lines[i].trim());
            if (columns.length === 0 || (columns.length === 1 && columns[0] === "")) continue;
            let rowObj = {};
            headers.forEach((h, index) => {
                rowObj[h] = columns[index] ? columns[index].trim().replace(/"/g, '') : '';
            });
            loadedRows.push(rowObj);
        }
    }
    
    if (loadedRows.length > 0) {
        // Stamp stable index on every row — avoids expensive indexOf() calls
        loadedRows.forEach((r, i) => { r._idx = i; });
        globalData = loadedRows;
        // Reset date picker on new file upload
        const dp = document.getElementById('datePicker');
        dp.value = '';
        dp.disabled = true;
        document.getElementById('clearDateBtn').classList.add('hidden');
        document.getElementById('latestDateLabel').innerText = '—';
        document.getElementById('latestDateCount').innerText = '—';
        document.getElementById('searchBar').disabled = false;
        document.getElementById('searchBar').placeholder = "Mag-type rito para mag-filter ng data...";
        document.getElementById('btnDashboard').click(); buildDashboard(globalData);
        // 🔴 Automatic duplicate ID check — lalabas agad ang popup kung may duplicates
        checkAndShowDuplicates(globalData, isFreshLoad);
        saveAppState();
    } else {
        // ✅ FIX: Kapag live Google Sheets mode at may existing data na, huwag i-clear.
        // Baka ang Google Sheets ay nag-return ng walang valid rows dahil sa HTML redirect
        // o pansamantalang error sa pag-parse — panatilihin ang last known good data.
        if (_liveSheetsUrl && globalData.length > 0) return;
        // Walang data rows — ipakita ang blangkong dashboard (huwag mag-alert)
        globalData = [];
        const dp = document.getElementById('datePicker');
        dp.value = ''; dp.disabled = true;
        document.getElementById('clearDateBtn').classList.add('hidden');
        document.getElementById('latestDateLabel').innerText = '—';
        document.getElementById('latestDateCount').innerText = '—';
        document.getElementById('searchBar').disabled = false;
        document.getElementById('searchBar').placeholder = "Mag-type rito para mag-filter ng data...";
        document.getElementById('btnDashboard').click(); buildDashboard([]);
        _hideDuplicateBadge(); _lastDuplicateSignature = null; _currentDuplicateGroups = [];
        saveAppState();
    }

}

// ====================================================================
// 📊 LIVE EXCEL MODAL — multi-sheet Excel selector for Attendance & Dashboard
// ====================================================================

// Shared helper: i-convert lahat ng date cells sa isang sheet papuntang YYYY-MM-DD
// string bago i-export sa CSV. Ginagamit ng main uploader, snapshot upload, AT
// ng live-watch polling loop — iisa lang ang logic para walang magkakaibang resulta.
function convertSheetDatesToISO(sheet) {
    if (!sheet['!ref']) return;
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let r = range.s.r; r <= range.e.r; r++) {
        for (let c = range.s.c; c <= range.e.c; c++) {
            const addr = XLSX.utils.encode_cell({ r, c });
            const cell = sheet[addr];
            if (cell && cell.t === 'd' && cell.v instanceof Date) {
                const d = cell.v;
                const iso = d.getFullYear() + '-' +
                    String(d.getMonth() + 1).padStart(2, '0') + '-' +
                    String(d.getDate()).padStart(2, '0');
                cell.t = 's'; cell.v = iso; cell.w = iso;
            }
        }
    }
}

// Shared helper: gumawa ng mga pindutan para sa bawat sheet sa workbook, at
// awtomatikong piliin ang una. Ginagamit pareho ng live-link at snapshot path.
function renderSheetButtons(sheetNames) {
    const container = document.getElementById('liveExcelSheetButtons');
    container.innerHTML = '';
    sheetNames.forEach((name, idx) => {
        const btn = document.createElement('button');
        btn.dataset.sheet = name;
        btn.className = idx === 0
            ? 'px-3 py-1.5 rounded-lg text-xs font-bold bg-emerald-600 text-white border border-emerald-600 cursor-pointer'
            : 'px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 hover:bg-emerald-500 hover:text-white border border-slate-300 dark:border-slate-600 cursor-pointer transition-colors';
        btn.innerText = '📄 ' + name;
        btn.onclick = () => selectLiveExcelSheet(name);
        container.appendChild(btn);
    });
    document.getElementById('liveExcelSheetSelector').classList.remove('hidden');
    selectLiveExcelSheet(sheetNames[0]);
}

function openLiveExcelModal() {
    // Reset the modal to a clean state
    const fileInput = document.getElementById('liveExcelFileInput');
    if (fileInput) fileInput.value = '';
    const sheetsInput = document.getElementById('googleSheetsUrlInput');
    if (sheetsInput) sheetsInput.value = '';
    document.getElementById('liveExcelFileName').textContent = 'Wala pang napiling file';
    document.getElementById('liveExcelSheetSelector').classList.add('hidden');
    document.getElementById('liveExcelSheetButtons').innerHTML = '';
    document.getElementById('liveExcelPreview').classList.add('hidden');
    document.getElementById('liveExcelLoadBtn').disabled = true;
    _liveExcelWorkbook = null;
    _pendingLiveFileHandle = null; // huwag dalhin ang dating di-pa-committed na pick
    document.getElementById('liveExcelModal').classList.remove('hidden');
}

function closeLiveExcelModal() {
    document.getElementById('liveExcelModal').classList.add('hidden');
}

function selectLiveExcelSheet(sheetName) {
    _liveExcelSelectedSheet = sheetName;
    // Highlight selected button
    const btns = document.getElementById('liveExcelSheetButtons').querySelectorAll('button');
    btns.forEach(btn => {
        if (btn.dataset.sheet === sheetName) {
            btn.className = 'px-3 py-1.5 rounded-lg text-xs font-bold bg-emerald-600 text-white border border-emerald-600 cursor-pointer';
        } else {
            btn.className = 'px-3 py-1.5 rounded-lg text-xs font-bold bg-slate-200 dark:bg-slate-700 text-slate-700 dark:text-slate-300 hover:bg-emerald-500 hover:text-white border border-slate-300 dark:border-slate-600 cursor-pointer transition-colors';
        }
    });
    // Show preview stats
    const sheet = _liveExcelWorkbook.Sheets[sheetName];
    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : null;
    const rowCount = range ? (range.e.r - range.s.r) : 0;
    const colCount = range ? (range.e.c - range.s.c + 1) : 0;
    document.getElementById('liveExcelPreviewText').innerText = '📄 Sheet: "' + sheetName + '"';
    document.getElementById('liveExcelPreviewRows').innerText = rowCount.toLocaleString() + ' rows × ' + colCount + ' columns';
    document.getElementById('liveExcelPreview').classList.remove('hidden');
    document.getElementById('liveExcelLoadBtn').disabled = false;
}

function loadSelectedExcelSheet() {
    if (!_liveExcelWorkbook || !_liveExcelSelectedSheet || !_liveExcelFile) return;
    const sheet = _liveExcelWorkbook.Sheets[_liveExcelSelectedSheet];
    convertSheetDatesToISO(sheet);
    const contents = XLSX.utils.sheet_to_csv(sheet);
    const fileName = _liveExcelFile.name + ' › ' + _liveExcelSelectedSheet;
    closeLiveExcelModal();

    const badge = document.getElementById('liveExcelActiveBadge');

    // Tanging isang live source lang dapat ang aktibo sa isang pagkakataon
    clearGoogleSheetsPolling();
    _liveSheetsUrl = null;
    try { localStorage.removeItem('mto_live_sheets_url'); } catch(e) {}

    if (_pendingLiveFileHandle) {
        // ✅ May live-linked handle — i-commit bilang ACTIVE at simulan ang polling
        _liveFileHandle = _pendingLiveFileHandle;
        _liveFileLastModified = _liveExcelFile.lastModified;
        _liveFileFailCount = 0;
        clearLivePolling();
        _livePollIntervalId = setInterval(checkLiveFileForChanges, LIVE_POLL_MS);
        if (badge) { badge.textContent = '🔴 ' + _liveExcelSelectedSheet; badge.classList.remove('hidden'); }
        updateLiveSyncIndicator(true, false, 'local');
    } else {
        // Plain snapshot lang — itigil ang dating live sync (kung meron) dahil
        // static na ngayon ang ipinapakita sa dashboard
        clearLivePolling();
        _liveFileHandle = null;
        if (badge) { badge.textContent = _liveExcelSelectedSheet; badge.classList.remove('hidden'); }
        updateLiveSyncIndicator(false);
    }
    _pendingLiveFileHandle = null;

    processFileContents(contents, fileName);
}

// ====================================================================
// 🔴 LIVE LOCAL FILE WATCH — File System Access API (Chrome/Edge desktop lang)
// Hindi tulad ng regular <input type="file">, ang showOpenFilePicker() ay
// nagbibigay ng FileSystemFileHandle na pwedeng paulit-ulit basahin ulit nang
// hindi na kailangang mag-upload — kaya posible ang totoong "live" na pag-sync.
// ====================================================================
async function linkLiveLocalFile() {
    if (!('showOpenFilePicker' in window)) {
        alert('⚠️ Hindi suportado ng browser mo ang Live Local File feature.\n\nGumamit ng Google Chrome o Microsoft Edge (desktop) para gumana ito. Hindi ito available sa Safari o Firefox — pwede mo pa ring gamitin ang "Minsan Lang Mag-upload" option.');
        return;
    }
    try {
        const [handle] = await window.showOpenFilePicker({
            types: [{
                description: 'Excel Files',
                accept: {
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                    'application/vnd.ms-excel': ['.xls']
                }
            }],
            excludeAcceptAllOption: false,
            multiple: false
        });

        const file = await handle.getFile();
        const buf = await file.arrayBuffer();
        const data = new Uint8Array(buf);

        _pendingLiveFileHandle = handle; // hindi pa ACTIVE — magko-commit lang pagka-click ng "I-load sa Dashboard"
        _liveExcelFile = file;
        _liveExcelWorkbook = XLSX.read(data, { type: 'array', cellDates: true });

        document.getElementById('liveExcelFileName').innerHTML =
            '<span class="text-red-600 dark:text-red-400 font-black">🔴 ' + file.name + '</span> <span class="text-slate-400 font-normal">(live-linked)</span>';

        renderSheetButtons(_liveExcelWorkbook.SheetNames);
    } catch (err) {
        if (err.name === 'AbortError') return; // kinansela lang ng user ang picker — tahimik na huminto
        console.error('Live file link error:', err);
        let msg = '⚠️ Hindi na-link ang file.';
        if (err.name === 'SecurityError' || err.name === 'NotAllowedError') {
            msg += '\n\nKaramihan, ito ay dahil binuksan ang dashboard sa pamamagitan lang ng dobleng-click sa index.html. Kailangan itong i-serve via local server (hal. "Live Server" extension sa VS Code, o python -m http.server) o i-host online para gumana ang Live Local File link.';
        }
        alert(msg);
    }
}

function clearLivePolling() {
    if (_livePollIntervalId) {
        clearInterval(_livePollIntervalId);
        _livePollIntervalId = null;
    }
}

// Tinatawag bawat ilang segundo habang naka-live-link. Tinitingnan ang lastModified
// ng file — kung walang pagbabago, wala ring ginagawa (mabilis at di nakaka-abala).
async function checkLiveFileForChanges() {
    if (!_liveFileHandle) { clearLivePolling(); return; }
    try {
        const file = await _liveFileHandle.getFile();
        _liveFileFailCount = 0; // matagumpay na nabasa — i-reset ang error counter
        if (file.lastModified === _liveFileLastModified) return; // walang bagong pagbabago
        _liveFileLastModified = file.lastModified;

        const buf = await file.arrayBuffer();
        const data = new Uint8Array(buf);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });

        if (!workbook.Sheets[_liveExcelSelectedSheet]) {
            console.warn('Live-linked sheet hindi na natagpuan pagkatapos ng update:', _liveExcelSelectedSheet);
            return; // baka na-rename o natanggal — laktawan ang cycle na ito, panoorin pa rin
        }
        const sheet = workbook.Sheets[_liveExcelSelectedSheet];
        convertSheetDatesToISO(sheet);
        const contents = XLSX.utils.sheet_to_csv(sheet);
        const fileName = file.name + ' › ' + _liveExcelSelectedSheet;
        processFileContents(contents, fileName, false); // background poll — huwag i-reset ang auto-detect state
        flashLiveSyncIndicator();
    } catch (err) {
        console.warn('Live file check failed:', err);
        _liveFileFailCount++;
        if (_liveFileFailCount >= 3) {
            // Tatlong magkakasunod na error (hal. na-delete o na-move ang file) — itigil na
            clearLivePolling();
            _liveFileHandle = null;
            updateLiveSyncIndicator(false, true);
        }
    }
}

function stopLiveFileWatchManually() {
    clearLivePolling();
    _liveFileHandle = null;
    _liveFileLastModified = null;
    clearGoogleSheetsPolling();
    _liveSheetsUrl = null;
    try { localStorage.removeItem('mto_live_sheets_url'); } catch(e) {}
    updateLiveSyncIndicator(false);
}

function updateLiveSyncIndicator(isLive, isError, source) {
    const row = document.getElementById('liveSyncStatusRow');
    const text = document.getElementById('liveSyncStatusText');
    if (!row || !text) return;
    if (isError) {
        row.classList.remove('hidden');
        text.innerText = '⚠️ Nawalan ng koneksyon sa ' + (source === 'sheets' ? 'Google Sheets' : 'file');
        return;
    }
    if (isLive) {
        row.classList.remove('hidden');
        const label = source === 'sheets' ? ('Google Sheets • ' + (_liveExcelSelectedSheet || 'live')) : ('Live syncing • ' + (_liveExcelSelectedSheet || ''));
        text.innerText = label;
    } else {
        row.classList.add('hidden');
    }
}

function flashLiveSyncIndicator(source) {
    const text = document.getElementById('liveSyncStatusText');
    if (!text) return;
    const stamp = new Date().toLocaleTimeString('en-PH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    text.innerText = '✅ Na-update • ' + stamp + '  •  I-click ang 🔗 Sheets para mag-refresh agad';
    setTimeout(() => {
        if (_liveFileHandle) { text.innerText = 'Live syncing • ' + (_liveExcelSelectedSheet || ''); }
        else if (_liveSheetsUrl) { text.innerText = 'Google Sheets • ' + (_liveExcelSelectedSheet || 'live') + '  •  I-click ang badge para mag-refresh'; }
    }, 4000);
}

// ====================================================================
// 🔗 LIVE GOOGLE SHEETS WATCH — gumagana sa kahit anong browser/computer
// dahil sentral na URL ang kinukuha (fetch), hindi local file lang.
// Ito ang tamang paraan kapag ang datos ay nasa Google Sheets, hindi sa
// isang .xlsx file na nakatago lang sa harddisk ng isang computer.
// ====================================================================

// Kunin ang sheet ID at gid mula sa URL na kinopya mula sa address bar ng
// Google Sheets (e.g. .../spreadsheets/d/SHEET_ID/edit?gid=GID), at gawing
// CSV export URL na pwedeng i-fetch nang walang pag-log-in.
function parseGoogleSheetsUrl(url) {
    const u = url.trim();

    // ── CASE 1: Pub/CSV URL na — galing sa "Publish to web" dialog ──────────
    // Format: .../d/e/2PACX-.../pub?gid=GID&single=true&output=csv
    // I-use nang direkta — huwag nang i-reconstruct, masisira lang
    if (u.includes('output=csv') || (u.includes('/pub?') && u.includes('gid='))) {
        const gidMatch = u.match(/[?&]gid=([0-9]+)/);
        const gid = gidMatch ? gidMatch[1] : '0';
        // Siguraduhing may output=csv sa URL (baka wala pa)
        const csvUrl = u.includes('output=csv') ? u : (u + (u.includes('?') ? '&' : '?') + 'output=csv');
        return { sheetId: 'pub', gid, csvUrl };
    }

    // ── CASE 2: Regular edit URL ─────────────────────────────────────────────
    // Format: .../d/SHEET_ID/edit?gid=GID
    // Nota: kailangan pa rin na naka-publish ang sheet para gumana ito
    const idMatch = u.match(/\/d\/([a-zA-Z0-9-_]{20,})/);
    if (!idMatch) return null;
    const gidMatch = u.match(/[#&?]gid=([0-9]+)/);
    const gid = gidMatch ? gidMatch[1] : '0';
    return {
        sheetId: idMatch[1],
        gid,
        csvUrl: 'https://docs.google.com/spreadsheets/d/' + idMatch[1] + '/pub?output=csv&gid=' + gid + '&single=true'
    };
}

function linkGoogleSheetsUrl() {
    const input = document.getElementById('googleSheetsUrlInput');
    const url = input ? input.value.trim() : '';
    if (!url) { alert('⚠️ Maglagay muna ng Google Sheets URL — kopyahin mula sa address bar habang nasa tamang tab ka.'); return; }
    const parsed = parseGoogleSheetsUrl(url);
    if (!parsed) { alert('⚠️ Hindi nakilalang Google Sheets URL. Siguraduhing kinopya mula sa address bar ng Google Sheets (may "/spreadsheets/d/" dapat).'); return; }
    testAndLoadGoogleSheets(parsed, true);
}

async function testAndLoadGoogleSheets(parsed, isManualLink) {
    const btn = document.getElementById('googleSheetsLinkBtn');
    if (btn) { btn.disabled = true; btn.innerText = 'Sinusuri...'; }
    try {
        const resp = await fetch(parsed.csvUrl + '&_cb=' + Date.now(), { cache: 'no-store' });
        if (!resp.ok) throw new Error('HTTP_' + resp.status);
        const text = await resp.text();
        const sniff = text.trim().slice(0, 100).toLowerCase();
        if (sniff.startsWith('<!doctype html') || sniff.startsWith('<html')) {
            throw new Error('NOT_SHARED');
        }

        // Isa lang dapat ang aktibong live source — itigil ang local file live kung meron
        clearLivePolling();
        _liveFileHandle = null;

        _liveSheetsUrl = parsed.csvUrl;
        _liveSheetsLastContent = text;
        _liveExcelSelectedSheet = 'gid:' + parsed.gid;
        const fileName = 'Google Sheets (live, gid:' + parsed.gid + ')';
        processFileContents(text, fileName);
        clearGoogleSheetsPolling();
        _liveSheetsPollIntervalId = setInterval(checkGoogleSheetsForChanges, LIVE_SHEETS_POLL_MS);

        try { localStorage.setItem('mto_live_sheets_url', parsed.csvUrl); } catch(e) {}

        const badge = document.getElementById('liveExcelActiveBadge');
        if (badge) {
            badge.textContent = '🔗 Sheets';
            badge.classList.remove('hidden');
            badge.style.background = '';
            badge.style.color = '';
            badge.style.cursor = 'pointer';
            badge.title = '🔄 I-click para mag-refresh agad';
            badge.onclick = () => { if (!_liveSheetsFetching) checkGoogleSheetsForChanges(); };
        }
        updateLiveSyncIndicator(true, false, 'sheets');

        if (isManualLink) closeLiveExcelModal();
    } catch (err) {
        console.error('Google Sheets link error:', err);
        if (!isManualLink) return; // auto-start sa background lang — huwag manakot ng alert, tahimik na huminto
        let msg = '';
        const isCorsOrNetwork = (err instanceof TypeError) || err.name === 'TypeError' || err.name === 'NetworkError';
        if (isCorsOrNetwork || err.message === 'NOT_SHARED' || err.message === 'HTTP_403' || err.message === 'HTTP_401') {
            msg = '⚠️ Kailangan munang i-publish ang Google Sheet mo bago ito gumana.\n\nGawin ito isang beses:\n1. Buksan ang Google Sheets mo\n2. File → Share → Publish to web\n3. Sa "Link" tab, piliin ang tamang Sheet/tab (hal. "MT Order")\n4. Sa dropdown sa kanan, piliin "Comma-separated values (.csv)"\n5. I-click ang "Publish" → OK\n6. Bumalik dito at i-paste ulit ang URL → I-link ulit\n\nPagkatapos ng publish, kahit kanselahin mo ang dialog — okay na yun, hindi mo na kailangang kopyahin ang link mula doon.';
        } else if (err.message === 'HTTP_404') {
            msg = '⚠️ Hindi natagpuan ang sheet/tab na iyan. Siguraduhing tama ang URL at buo ang gid sa dulo nito.';
        } else {
            msg = '⚠️ Hindi ma-access ang Google Sheets na ito (' + (err.message || err) + ').\n\nSiguraduhing:\n1. Na-publish ang sheet (File → Share → Publish to web → CSV)\n2. Tama ang URL na kinopya mula sa address bar ng Google Sheets';
        }
        alert(msg);
    } finally {
        if (btn) { btn.disabled = false; btn.innerText = 'I-link'; }
    }
}

function clearGoogleSheetsPolling() {
    if (_liveSheetsPollIntervalId) {
        clearInterval(_liveSheetsPollIntervalId);
        _liveSheetsPollIntervalId = null;
    }
}

// ⚡ INSTANT REFRESH ON TAB RETURN — ito ang pangunahing pagpapabilis:
// Kapag umalis ka sa tab para mag-edit sa Google Sheets tapos bumalik ka,
// hindi na kailangan pang hintayin ang susunod na 3-second interval —
// agad mag-fetch nang mag-switch pabalik sa dashboard tab.
//
// Ginagawa rin nito ang adaptive polling:
//   • Naka-focus ang tab  → 3 segundong interval (LIVE_SHEETS_POLL_MS)
//   • Naka-background ang tab → 10 segundong interval (LIVE_SHEETS_POLL_HIDDEN_MS)
//   Nakatitipid ito ng bandwidth kapag hindi tinitingnan ang dashboard.
(function _initTabFocusRefresh() {
    let _lastVisibility = !document.hidden;

    function _restartPollingAdaptive() {
        if (!_liveSheetsUrl) return;
        clearGoogleSheetsPolling();
        const ms = document.hidden ? LIVE_SHEETS_POLL_HIDDEN_MS : LIVE_SHEETS_POLL_MS;
        _liveSheetsPollIntervalId = setInterval(checkGoogleSheetsForChanges, ms);
    }

    document.addEventListener('visibilitychange', function () {
        const nowVisible = !document.hidden;
        if (nowVisible && !_lastVisibility) {
            // Bumalik sa tab — agad mag-fetch TAPOS i-restart sa fast interval
            if (_liveSheetsUrl && !_liveSheetsFetching) checkGoogleSheetsForChanges();
            _restartPollingAdaptive();
        } else if (!nowVisible && _lastVisibility) {
            // Umalis sa tab — lumipat sa slow interval para hindi mag-aksaya ng bandwidth
            _restartPollingAdaptive();
        }
        _lastVisibility = nowVisible;
    });

    window.addEventListener('focus', function () {
        // Backup para sa mga browser na hindi laging nag-fire ang visibilitychange
        if (_liveSheetsUrl && !_liveSheetsFetching && !document.hidden) {
            checkGoogleSheetsForChanges();
        }
    });
})();

async function checkGoogleSheetsForChanges() {
    if (!_liveSheetsUrl) { clearGoogleSheetsPolling(); return; }
    if (_liveSheetsFetching) return; // ⚡ huwag mag-overlap ng concurrent fetches
    _liveSheetsFetching = true;

    // Ipakita ang "checking..." pulse sa sync indicator
    const syncText = document.getElementById('liveSyncStatusText');
    const prevText = syncText ? syncText.innerText : '';
    if (syncText && !prevText.includes('✅')) {
        syncText.innerText = '🔄 Sinisigurado ang pinakabago...';
    }

    try {
        const resp = await fetch(_liveSheetsUrl + '&_cb=' + Date.now(), { cache: 'no-store' });
        if (!resp.ok) throw new Error('HTTP_' + resp.status);
        const text = await resp.text();

        // ✅ FIX 1 (PANGUNAHING SANHI NG FLICKER): I-reject ang HTML redirect pages.
        const sniff = text.trim().slice(0, 300).toLowerCase();
        if (sniff.startsWith('<!doctype') || sniff.startsWith('<html') ||
            sniff.includes('<title>') || sniff.includes('google sign in') ||
            sniff.includes('<script')) {
            console.warn('Google Sheets poll: HTML response natanggap, nilalaktawan (redirect/login page).');
            return;
        }

        // ✅ FIX 2: Trim bago ikompara — Google minsan nagdadagdag ng trailing newlines/BOM
        const trimmed = text.trim();
        if (trimmed === (_liveSheetsLastContent || '').trim()) {
            // Walang pagbabago — i-restore lang ang indicator text
            if (syncText && syncText.innerText.includes('Sinisigurado')) {
                syncText.innerText = 'Google Sheets • ' + (_liveExcelSelectedSheet || 'live');
            }
            return;
        }
        _liveSheetsLastContent = text;
        processFileContents(text, 'Google Sheets (live)', false);
        flashLiveSyncIndicator('sheets');
    } catch (err) {
        console.warn('Google Sheets poll failed:', err);
        if (syncText && syncText.innerText.includes('Sinisigurado')) {
            syncText.innerText = prevText || 'Google Sheets • ' + (_liveExcelSelectedSheet || 'live');
        }
    } finally {
        _liveSheetsFetching = false; // palaging i-release ang guard
    }
}

// Awtomatikong sinisimulan sa bawat pag-load ng page — ginagamit ang URL na
// dati nang na-link sa device na ito (localStorage), o kung wala pa, ang
// DEFAULT_LIVE_SHEETS_URL na nakatakda sa itaas ng file na ito. Kapag
// bumalik ang dashboard sa naka-publish na site, automatic na itong susunod
// nang walang kahit sinong i-click pa.
async function initLiveDataSource() {
    let savedUrl = null;
    try { savedUrl = localStorage.getItem('mto_live_sheets_url'); } catch(e) {}

    // I-migrate ang lumang /export?format=csv URL papuntang bagong /pub?output=csv format
    // Kung hindi miae-clear ito, patuloy na magfa-fail ang auto-start sa background
    if (savedUrl && savedUrl.includes('/export?format=csv')) {
        try { localStorage.removeItem('mto_live_sheets_url'); } catch(e) {}
        savedUrl = null;
    }

    let csvUrlToUse = savedUrl;
    let gidToUse = null;
    if (!csvUrlToUse && DEFAULT_LIVE_SHEETS_URL) {
        const parsed = parseGoogleSheetsUrl(DEFAULT_LIVE_SHEETS_URL);
        if (parsed) { csvUrlToUse = parsed.csvUrl; gidToUse = parsed.gid; }
    }

    if (!csvUrlToUse) { loadAppState(); return; } // walang naka-configure — gamitin ang dating cached na data

    try {
        const resp = await fetch(csvUrlToUse + '&_cb=' + Date.now(), { cache: 'no-store' });
        if (!resp.ok) throw new Error('HTTP_' + resp.status);
        const text = await resp.text();
        const sniff = text.trim().slice(0, 100).toLowerCase();
        if (sniff.startsWith('<!doctype html') || sniff.startsWith('<html')) throw new Error('NOT_SHARED');

        _liveSheetsUrl = csvUrlToUse;
        _liveSheetsLastContent = text;
        const gidMatch = csvUrlToUse.match(/gid=([0-9]+)/);
        _liveExcelSelectedSheet = 'gid:' + (gidMatch ? gidMatch[1] : (gidToUse || '0'));
        processFileContents(text, 'Google Sheets (live)');
        clearGoogleSheetsPolling();
        _liveSheetsPollIntervalId = setInterval(checkGoogleSheetsForChanges, LIVE_SHEETS_POLL_MS);

        const badge = document.getElementById('liveExcelActiveBadge');
        if (badge) {
            badge.textContent = '🔗 Sheets';
            badge.classList.remove('hidden');
            badge.style.background = '';
            badge.style.color = '';
            badge.style.cursor = 'pointer';
            badge.title = '🔄 I-click para mag-refresh agad';
            badge.onclick = () => { if (!_liveSheetsFetching) checkGoogleSheetsForChanges(); };
        }
        updateLiveSyncIndicator(true, false, 'sheets');
    } catch (err) {
        console.warn('Auto Google Sheets live load failed, falling back sa cached data:', err);
        // ✅ FIX: Ipakita ang error para malaman ng user na hindi nakakonekta
        const badge = document.getElementById('liveExcelActiveBadge');
        if (badge) {
            badge.textContent = '⚠️ Hindi nakakonekta';
            badge.classList.remove('hidden');
            badge.style.background = '#ef4444';
            badge.style.color = '#fff';
            badge.style.cursor = 'pointer';
            badge.title = 'I-click para mag-link ng tamang Google Sheets URL';
            badge.onclick = () => openLiveExcelModal();
        }
        updateLiveSyncIndicator(false, true, 'sheets');
        loadAppState(); // huwag iwanang blangko ang dashboard kung di pa ma-access ang sheet
    }
}

// ====================================================================
// 📂 CORE FILE ROUTER (EXCEL & CSV ACCURATE CONVERSION)
// ====================================================================
document.getElementById('csvFileInput').addEventListener('change', function(e) {
    let file = e.target.files[0]; if (!file) return;
    let reader = new FileReader();
    let isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
    
    reader.onload = function(evt) {
        let contents = "";
        if (isExcel) {
            let data = new Uint8Array(evt.target.result);
            let workbook = XLSX.read(data, { type: 'array', cellDates: true });
            let sheet = workbook.Sheets[workbook.SheetNames[0]];
            convertSheetDatesToISO(sheet);
            contents = XLSX.utils.sheet_to_csv(sheet);
        } else { contents = evt.target.result.replace(/^\uFEFF/, ''); }
        processFileContents(contents, file.name);
    };
    if (isExcel) { reader.readAsArrayBuffer(file); } else { reader.readAsText(file); }
});


// ====================================================================
// 🔴 DUPLICATE ID AUTO-DETECTION — automatic na nag-scan ng buong dataset
// pagkatapos mag-load ng file (xlsx, csv, o Google Sheets live). Kapag may
// nakitang dalawa o higit pang row na magkaparehong MTB ID (o katumbas na
// primary ID field ng kasalukuyang fileType), awtomatikong lumalabas ang
// popup at nagpapakita kung aling IDs ang duplicated at ano ang laman nila.
// Sa background polling (live sync), lalabas lang ang modal kapag NAGBAGO
// ang set ng duplicates — hindi ito spam-pop-up tuwing nag-ri-refresh.
// ====================================================================

/** Ibinabalik ang primary ID field name ng kasalukuyang fileType para sa duplicate check. */
function _getDupIDField() {
    if (fileType === 'raw' || fileType === 'crossdock') return 'mtb';
    if (fileType === 'device') return 'deviceId';
    if (fileType === 'cancelled') return 'mtOrder';
    if (fileType === 'progress') return 'batch';
    if (fileType === 'generic') {
        const h = (genericHeaders || []).find(h => /id|mtb|batch|order|number|serial|series|code/i.test(h));
        return h || (genericHeaders && genericHeaders[0]) || null;
    }
    return null;
}

/** Nagbu-build ng label para sa ID field base sa fileType — para sa pamagat ng modal. */
function _getDupIDLabel() {
    if (fileType === 'raw' || fileType === 'crossdock') return 'MTB ID';
    if (fileType === 'device') return 'Device ID';
    if (fileType === 'cancelled') return 'MT Order';
    if (fileType === 'progress') return 'Batch';
    return 'ID';
}

/**
 * Ini-scan ang buong data para sa mga duplicate primary ID values.
 * @returns {Array<{id:string, rows:Array}>} — isa per unique duplicated ID, sorted by count desc.
 */
function detectDuplicateIDs(data) {
    if (!data || data.length === 0) return [];
    const field = _getDupIDField();
    if (!field) return [];

    const map = {};
    data.forEach(row => {
        const val = (row[field] || '').trim();
        if (!val) return;
        if (!map[val]) map[val] = { id: val, rows: [] };
        map[val].rows.push(row);
    });
    return Object.values(map).filter(g => g.rows.length > 1)
                             .sort((a, b) => b.rows.length - a.rows.length);
}

/** Gumagawa ng compact row card para sa loob ng duplicate group. */
function _renderDupRowCard(row) {
    let cells = '';
    if (fileType === 'raw') {
        cells = `
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Series</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200 break-words">${row.series || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">WH</span><p class="text-xs font-bold text-orange-600 dark:text-orange-400">${row.wh || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Status</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.status || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Type</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.type || '—'}</p></div>
            ${row.date ? `<div class="col-span-2"><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Date</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.date}</p></div>` : ''}`;
    } else if (fileType === 'crossdock') {
        cells = `
            <div class="col-span-2"><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">SKU Name</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200 break-words">${row.skuName || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Req Qty</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.qty || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Status</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.statusReceived || '—'}</p></div>`;
    } else if (fileType === 'device') {
        cells = `
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Type</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.deviceType || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Status</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.status || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Template</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.templateId || '—'}</p></div>`;
    } else if (fileType === 'cancelled') {
        cells = `
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Picking ID</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.pickingId || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">WHS</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.whs || '—'}</p></div>
            <div class="col-span-2"><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">SKU Name</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200 break-words">${row.skuName || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Reason</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.reason || '—'}</p></div>`;
    } else if (fileType === 'progress') {
        cells = `
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">WH</span><p class="text-xs font-bold text-orange-600 dark:text-orange-400">${row.wh || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Month</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${row.month || '—'}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Forecast</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${(row.forecast||0).toLocaleString()}</p></div>
            <div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">Shipped</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200">${(row.shipped||0).toLocaleString()}</p></div>`;
    } else {
        // generic — ipakita lahat ng visible columns
        const field = _getDupIDField();
        cells = (genericHeaders || []).filter(h => h !== field).slice(0,6).map(h =>
            `<div><span class="text-slate-400 dark:text-slate-500 text-[10px] uppercase font-bold">${h}</span><p class="text-xs font-semibold text-slate-700 dark:text-slate-200 break-words">${row[h] || '—'}</p></div>`
        ).join('');
    }
    return `<div class="bg-white dark:bg-slate-950 rounded-lg p-3 border border-red-100 dark:border-red-900/30 grid grid-cols-2 gap-2">${cells}</div>`;
}

/** Gumagawa ng HTML card para sa isang duplicate group (iisang ID na lumabas ng 2+x). */
function _renderDupGroupCard(group) {
    const rowCards = group.rows.map((r, i) =>
        `<div class="space-y-1"><p class="text-[10px] font-black text-red-400 dark:text-red-500 uppercase tracking-wider mb-1">Entry ${i + 1}</p>${_renderDupRowCard(r)}</div>`
    ).join('');
    return `<div class="border border-red-200 dark:border-red-800/40 rounded-xl p-4 bg-red-50/40 dark:bg-red-950/10">
        <div class="flex items-center gap-3 mb-3 flex-wrap">
            <span class="font-mono font-black text-sm text-red-700 dark:text-red-400 break-all">${group.id}</span>
            <span class="px-2.5 py-0.5 rounded-full text-[10px] font-black bg-red-100 text-red-700 border border-red-300 dark:bg-red-900/30 dark:text-red-400 dark:border-red-700/50 whitespace-nowrap">${group.rows.length}× DUPLICATED</span>
        </div>
        <div class="space-y-2">${rowCards}</div>
    </div>`;
}

/** Ipinapakita ang duplicate alert popup na may listahan ng lahat ng duplicated IDs. */
function showDuplicateModal(groups) {
    _currentDuplicateGroups = groups;
    const modal  = document.getElementById('duplicateDetectModal');
    const body   = document.getElementById('dupDetectBody');
    const sub    = document.getElementById('dupDetectSubtitle');
    if (!modal || !body) return;
    const label  = _getDupIDLabel();
    const count  = groups.length;
    sub.textContent = `${count} duplicate ${label}${count > 1 ? 's' : ''} found — ${groups.reduce((a,g)=>a+g.rows.length,0)} total affected rows`;
    body.innerHTML  = groups.map(_renderDupGroupCard).join('');
    modal.classList.remove('hidden');
}

function closeDuplicateModal() {
    const modal = document.getElementById('duplicateDetectModal');
    if (modal) modal.classList.add('hidden');
}

function _showDuplicateBadge(count) {
    const badge = document.getElementById('dupDetectBadge');
    const txt   = document.getElementById('dupDetectBadgeText');
    if (!badge) return;
    txt.textContent = `${count} Duplicate ${count === 1 ? _getDupIDLabel() : _getDupIDLabel() + 's'}`;
    badge.classList.remove('hidden');
}

function _hideDuplicateBadge() {
    const badge = document.getElementById('dupDetectBadge');
    if (badge) badge.classList.add('hidden');
}

/**
 * Pangunahing entry-point — tinatawag pagkatapos mag-load ng data.
 * @param {Array}   data        — ang na-load na rows (globalData)
 * @param {boolean} isFreshLoad — true = user-initiated upload; false = background poll
 */
function checkAndShowDuplicates(data, isFreshLoad) {
    // Only works on the Dashboard view with a known fileType
    if (!fileType || fileType === 'attendance') { _hideDuplicateBadge(); return; }

    const groups = detectDuplicateIDs(data);

    if (groups.length === 0) {
        // Wala nang duplicates — itago ang badge at i-clear ang signature
        _hideDuplicateBadge();
        _lastDuplicateSignature = null;
        _currentDuplicateGroups = [];
        return;
    }

    // Build signature para ma-detect kung nagbago ang set ng duplicates
    const sig = groups.map(g => `${g.id}:${g.rows.length}`).sort().join('|');
    const changed = sig !== _lastDuplicateSignature;

    _lastDuplicateSignature = sig;
    _currentDuplicateGroups = groups;
    _showDuplicateBadge(groups.length);

    // Auto-show ang popup: palagi sa fresh load; sa background poll, kapag nagbago lang
    if (isFreshLoad || changed) {
        showDuplicateModal(groups);
    }
}

// ====================================================================
// 🎯 AUTO-DETECT RECORD LOOKUP — kapag eksaktong tumugma ang na-type sa
// Search bar (hal. "MTB954027") sa isang ID field ng kasalukuyang fileType,
// awtomatikong lalabas ang detail popup. Gumagana ito kahit saang file type
// (raw, crossdock, device, cancelled, progress, generic) at kahit saan
// nanggaling ang data — manual xlsx/csv upload, Live Excel sheet, o Google
// Sheets live link — dahil iisa lang kasi ang globalData/fileType na
// pinagmumulan ng lahat ng ito.
// ====================================================================

/** Tumutukoy kung aling row(s) ang eksaktong tumutugma sa akmang ID field ng kasalukuyang fileType. */
function findExactMatchRows(rawQuery) {
    const q = (rawQuery || '').trim().toLowerCase();
    if (!q || !globalData || globalData.length === 0) return [];

    if (fileType === "raw" || fileType === "crossdock") {
        return globalData.filter(r => r.mtb && r.mtb.trim().toLowerCase() === q);
    }
    if (fileType === "device") {
        return globalData.filter(r =>
            (r.deviceId && r.deviceId.trim().toLowerCase() === q) ||
            (r.templateId && r.templateId !== '-' && r.templateId.trim().toLowerCase() === q)
        );
    }
    if (fileType === "cancelled") {
        return globalData.filter(r =>
            (r.mtOrder && r.mtOrder.trim().toLowerCase() === q) ||
            (r.pickingId && r.pickingId.trim().toLowerCase() === q)
        );
    }
    if (fileType === "progress") {
        return globalData.filter(r => r.batch && r.batch.trim().toLowerCase() === q);
    }
    if (fileType === "generic") {
        const idLikeHeaders = genericHeaders.filter(h => /id|mtb|batch|order|number|serial|series|code/i.test(h));
        const headersToCheck = idLikeHeaders.length > 0 ? idLikeHeaders : genericHeaders;
        return globalData.filter(r => headersToCheck.some(h => r[h] && String(r[h]).trim().toLowerCase() === q));
    }
    return [];
}

/** Helper para sa isang label/value block sa loob ng detail card. */
function _rdCell(label, valueHtml, wide) {
    return `<div class="${wide ? 'col-span-2' : ''}">
        <p class="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-wider mb-0.5">${label}</p>
        <div class="text-xs font-semibold text-slate-700 dark:text-slate-200 break-words">${valueHtml}</div>
    </div>`;
}

/** Gumagawa ng isang detail card (HTML string) para sa isang row, base sa kasalukuyang fileType. */
function renderRecordDetailCard(row) {
    const origIndex = row._idx;
    const fontMuted = "text-slate-500 dark:text-slate-400";
    const cardShell = (headerHtml, gridHtml) => `<div class="border border-orange-200 dark:border-orange-800/40 rounded-xl p-4 bg-orange-50/40 dark:bg-orange-950/10">
        <div class="flex items-center justify-between mb-3 flex-wrap gap-2">${headerHtml}</div>
        <div class="grid grid-cols-2 gap-3">${gridHtml}</div>
    </div>`;

    if (fileType === "raw") {
        let statusLow = row.status.toLowerCase();
        let isDone = statusLow.includes('done') || (statusLow.includes('ship') && !statusLow.includes('pending'));
        let statusBadge = isDone
            ? `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition-all">Done</button>`
            : `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 hover:bg-amber-200 transition-all">Pending</button>`;
        let typeColor = row.type === "HighValue" ? "bg-purple-100 text-purple-800 border-purple-300" : "bg-sky-100 text-sky-800 border-sky-300";
        let typeBadge = `<button onclick="toggleType(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-md text-[11px] font-black tracking-wide border transition-all ${typeColor}">${row.type}</button>`;

        return cardShell(
            `<span class="font-mono font-black text-base text-slate-800 dark:text-white">📦 ${row.mtb}</span>${typeBadge}`,
            _rdCell('Series / Serial Number', row.series || '—') +
            _rdCell('Destination Warehouse', `<span class="text-orange-600 dark:text-orange-400 font-bold">${row.wh || '—'}</span>`) +
            _rdCell('Qty', row.qty ? row.qty.toLocaleString() : '1') +
            _rdCell('Status', statusBadge) +
            _rdCell('Shipped Date', row.date || '—') +
            _rdCell('Remarks', row.remarks || '—', true)
        );
    }

    if (fileType === "crossdock") {
        let isReceived = row.statusReceived.toLowerCase().includes('done');
        let statusBadge = isReceived
            ? `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition-all">Done</button>`
            : `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 hover:bg-amber-200 transition-all">Pending</button>`;

        return cardShell(
            `<span class="font-mono font-black text-base text-slate-800 dark:text-white">📦 ${row.mtb}</span>${statusBadge}`,
            _rdCell('Date', row.date || '—') +
            _rdCell('SKU Name', row.skuName || '—', true) +
            _rdCell('Req Qty', row.qty.toLocaleString()) +
            _rdCell('Act Qty', `<span class="text-blue-600 dark:text-blue-400 font-bold">${row.actualQty.toLocaleString()}</span>`) +
            _rdCell('Pallet', row.pallet || '—') +
            _rdCell('Checked By', row.checkedBy || '—')
        );
    }

    if (fileType === "device") {
        let statusLow = (row.status || '').toLowerCase();
        let statusColor, statusLabel;
        if (statusLow === 'normal') { statusColor = 'background:#e0f2fe;color:#0369a1;border:1px solid #7dd3fc;'; statusLabel = 'Normal'; }
        else if (statusLow === 'done') { statusColor = 'background:#d1fae5;color:#065f46;border:1px solid #6ee7b7;'; statusLabel = 'Done'; }
        else { statusColor = 'background:#fef3c7;color:#92400e;border:1px solid #fcd34d;'; statusLabel = row.status || 'Pending'; }
        let statusBadge = `<button onclick="cycleDeviceStatus(${origIndex})" style="${statusColor}padding:2px 10px;border-radius:9999px;font-size:11px;font-weight:700;cursor:pointer;">${statusLabel}</button>`;
        let tplVal = (row.templateId && row.templateId !== '-' && row.templateId.trim() !== '') ? row.templateId : '—';

        return cardShell(
            `<span class="font-mono font-black text-base text-slate-800 dark:text-white">🖥️ ${row.deviceId}</span>${statusBadge}`,
            _rdCell('Device Type', row.deviceType || '—') +
            _rdCell('Template ID', tplVal) +
            _rdCell('Specification (Rows)', row.specRows !== undefined ? row.specRows : '0') +
            _rdCell('Specification (Columns)', row.specCols !== undefined ? row.specCols : '0') +
            _rdCell('Device Progress', row.progress || '—', true)
        );
    }

    if (fileType === "cancelled") {
        const reasonColors = {
            'NO OTHER LOCATION'  : 'bg-orange-100 text-orange-800 border-orange-300',
            'EXPIRY ITEMS'       : 'bg-red-100    text-red-800    border-red-300',
            'INVALID PICKING ID' : 'bg-purple-100 text-purple-800 border-purple-300',
            'LIQUOR'             : 'bg-sky-100    text-sky-800    border-sky-300',
            'INCOMPLETE CHECKING': 'bg-yellow-100 text-yellow-800 border-yellow-300',
            'SMALL ITEMS'        : 'bg-teal-100   text-teal-800   border-teal-300',
            'DAMAGED ITEMS'      : 'bg-rose-100   text-rose-800   border-rose-300',
            'LACKING ITEM/LOST'  : 'bg-pink-100   text-pink-800   border-pink-300'
        };
        let reasonKey = (row.reason || '').toUpperCase().trim();
        let reasonClass = reasonColors[reasonKey] || 'bg-slate-100 text-slate-700 border-slate-300';
        let reasonBadge = row.reason
            ? `<span class="px-2 py-0.5 rounded-full text-[10px] font-bold border ${reasonClass} whitespace-nowrap">${row.reason}</span>`
            : `<span class="${fontMuted} text-xs">—</span>`;
        let pickerBadge = (row.picker && row.picker !== '')
            ? `<span class="px-2 py-0.5 rounded-md text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">${row.picker}</span>`
            : `<span class="${fontMuted} text-xs italic">—</span>`;

        return cardShell(
            `<span class="font-mono font-black text-base text-slate-800 dark:text-white">🧾 ${row.mtOrder}</span>${reasonBadge}`,
            _rdCell('WHs', row.whs || '—') +
            _rdCell('Picking ID', row.pickingId || '—') +
            _rdCell('SKU', row.sku || '—') +
            _rdCell('SKU Name', row.skuName || '—', true) +
            _rdCell('Qty', row.qty > 0 ? row.qty.toLocaleString() : '—') +
            _rdCell('Picked Qty', row.pickedQty > 0 ? row.pickedQty.toLocaleString() : '—') +
            _rdCell('Checked Qty', row.checkedQty > 0 ? row.checkedQty.toLocaleString() : '—') +
            _rdCell('Lack Item', row.lackItem > 0 ? `<span class="text-red-500 font-bold">${row.lackItem.toLocaleString()}</span>` : '—') +
            _rdCell('Picker', pickerBadge, true)
        );
    }

    if (fileType === "progress") {
        let isDone = row.percentage >= 100 || row.discrepancy <= 0;
        let statusBadge = isDone
            ? `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300 hover:bg-emerald-200 transition-all">Done</button>`
            : `<button onclick="toggleStatus(${origIndex})" class="cursor-pointer px-2.5 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800 border border-amber-300 hover:bg-amber-200 transition-all">Pending</button>`;

        return cardShell(
            `<span class="font-mono font-black text-base text-slate-800 dark:text-white">📋 ${row.batch}</span>${statusBadge}`,
            _rdCell('Warehouse', `<span class="text-orange-600 dark:text-orange-400 font-bold">${row.wh || '—'}</span>`) +
            _rdCell('WW / Month', `${row.ww || '-'} • ${row.month || '-'}`) +
            _rdCell('Forecast', row.forecast.toLocaleString()) +
            _rdCell('Allocated', row.allocated.toLocaleString()) +
            _rdCell('Total Shipped', row.shipped.toLocaleString()) +
            _rdCell('Discrepancy', `<span class="text-red-500">${row.discrepancy.toLocaleString()}</span>`) +
            _rdCell('Percentage', `<span class="text-emerald-600 dark:text-emerald-400 font-bold">${row.percentage}%</span>`, true)
        );
    }

    // Generic fallback — ipakita ang lahat ng column na na-detect mula sa file
    let cells = genericHeaders.map(h => _rdCell(h, row[h] || '—')).join('');
    return `<div class="border border-orange-200 dark:border-orange-800/40 rounded-xl p-4 bg-orange-50/40 dark:bg-orange-950/10"><div class="grid grid-cols-2 gap-3">${cells}</div></div>`;
}

/** Ipinapakita ang popup na may detalye ng lahat ng eksaktong nahanap na row. */
function showRecordDetailModal(rows, displayQuery) {
    const modal = document.getElementById('recordDetailModal');
    const body = document.getElementById('recordDetailBody');
    const subtitle = document.getElementById('recordDetailSubtitle');
    if (!modal || !body || rows.length === 0) return;

    subtitle.textContent = rows.length > 1
        ? `${rows.length} matches found for "${displayQuery}"`
        : `1 match found for "${displayQuery}"`;
    body.innerHTML = rows.map(renderRecordDetailCard).join('');
    modal.classList.remove('hidden');
}

function closeRecordDetailModal() {
    const modal = document.getElementById('recordDetailModal');
    if (modal) modal.classList.add('hidden');
}

/** Kapag naka-open ang popup at binago ang laman ng isang row (hal. na-toggle ang status
 *  sa loob mismo ng popup), i-refresh ang laman nito para hindi ito maging stale. */
function refreshRecordDetailModalIfOpen() {
    const modal = document.getElementById('recordDetailModal');
    if (!modal || modal.classList.contains('hidden') || !_lastAutoDetectQuery) return;
    const matches = findExactMatchRows(_lastAutoDetectQuery);
    if (matches.length > 0) {
        document.getElementById('recordDetailBody').innerHTML = matches.map(renderRecordDetailCard).join('');
    } else {
        closeRecordDetailModal();
    }
}

// Dynamic Client Filter Control — debounced to avoid lag on fast typing
let _searchDebounceTimer = null;
document.getElementById('searchBar').addEventListener('input', function(e) {
    clearTimeout(_searchDebounceTimer);
    _searchDebounceTimer = setTimeout(function() {
    let query = e.target.value.toLowerCase().trim(); if (!query) { _lastAutoDetectQuery = null; buildDashboard(globalData); return; }
    let filtered = globalData.filter(row => {
        if (fileType === "progress") return row.wh.toLowerCase().includes(query) || row.batch.toLowerCase().includes(query) || row.month.toLowerCase().includes(query);
        if (fileType === "raw") return row.mtb.toLowerCase().includes(query) || row.wh.toLowerCase().includes(query) || row.series.toLowerCase().includes(query) || row.type.toLowerCase().includes(query) || (row.remarks && row.remarks.toLowerCase().includes(query));
        if (fileType === "crossdock") return row.mtb.toLowerCase().includes(query) || row.skuName.toLowerCase().includes(query) || row.checkedBy.toLowerCase().includes(query);
        if (fileType === "device") return row.deviceId.toLowerCase().includes(query) || row.deviceType.toLowerCase().includes(query) || row.status.toLowerCase().includes(query) || row.progress.toLowerCase().includes(query) || row.templateId.toLowerCase().includes(query);
        if (fileType === "cancelled") return row.mtOrder.toLowerCase().includes(query) || row.whs.toLowerCase().includes(query) || row.skuName.toLowerCase().includes(query) || row.pickingId.toLowerCase().includes(query) || (row.reason && row.reason.toLowerCase().includes(query)) || (row.picker && row.picker.toLowerCase().includes(query)) || (row.sku && row.sku.toLowerCase().includes(query));
        if (fileType === "generic") {
            return genericHeaders.some(h => String(row[h]).toLowerCase().includes(query));
        }
    });
    buildDashboard(filtered);

    // 🎯 AUTO-DETECT: kapag eksaktong tumugma ang query sa isang Box/MTB ID (o katumbas
    // na ID field ng kasalukuyang fileType), awtomatikong ipapakita ang detail popup —
    // isang beses lang kada natatanging query, para hindi ito paulit-ulit lumabas.
    let exactMatches = findExactMatchRows(query);
    if (exactMatches.length > 0 && query !== _lastAutoDetectQuery) {
        _lastAutoDetectQuery = query;
        showRecordDetailModal(exactMatches, e.target.value.trim());
    }
    }, 180); // debounce: wait 180ms after user stops typing
});

// ====================================================================
// 📅 ATTENDANCE TIMESHEET SYSTEM CONTROLLER
// ====================================================================
function parseAttendanceCSV(text) {
    const lines = text.split(/\r?\n/).map(line => parseCSVLine(line.trim())).filter(l => l.length > 0 && l.some(cell => cell.trim() !== ""));
    if (lines.length < 3) return;
    const headerRow = lines[0];
    const columnHeaders = lines[2] ? lines[2].map(h => h.toLowerCase().trim()) : [];
    parsedEmployeesData = {};
    
    let columnsPerEmployee = 7; 
    for (let i = 1; i < headerRow.length; i++) {
        if (headerRow[i] && headerRow[i].toUpperCase().includes('ATTENDANCE')) { columnsPerEmployee = i; break; }
    }

    for (let i = 0; i < headerRow.length; i += columnsPerEmployee) { 
        let empName = headerRow[i] ? headerRow[i].replace('ATTENDANCE', '').trim().toUpperCase() : "";
        if (!empName) continue;
        parsedEmployeesData[empName] = [];
        
        let otOffset = 5, allowanceOffset = -1; 
        for (let offset = 0; offset < columnsPerEmployee; offset++) {
            let hName = columnHeaders[i + offset] || "";
            if (hName.includes('allowance')) allowanceOffset = offset;
            if (hName.includes('of ot') || hName.includes('ot')) otOffset = offset;
        }

        for (let j = 3; j < lines.length; j++) {
            const row = lines[j]; if (!row || row.length <= i || !row[i] || row[i].trim() === "" || row[i].toLowerCase().includes('date')) continue; 
            let rawDate = row[i].trim(); let formattedDate = convertToHTMLDate(rawDate);
            const day = row[i+1] ? row[i+1].trim().toUpperCase() : "MONDAY";
            const shift = row[i+2] ? row[i+2].trim() : "PM SHIFT";
            const remarks = row[i+3] ? row[i+3].trim().toUpperCase() : "PRESENT";
            let dailyRate = parseFloat(row[i+4]) || 0;
            let allowance = allowanceOffset !== -1 && row[i + allowanceOffset] ? parseFloat(row[i + allowanceOffset]) || 0 : 0;
            let otHours = row[i + otOffset] ? parseFloat(row[i + otOffset]) || 0 : 0;
            
            if (dailyRate === 0 && remarks === "PRESENT") dailyRate = 650;
            if (dailyRate === 0 && remarks === "DOUBLE PAY") dailyRate = 1250;
            
            parsedEmployeesData[empName].push({ date: formattedDate, day, shift, remarks, dailyRate, allowance, otHours });
        }
    }
    refreshEditorUI();
}

function convertToHTMLDate(dateStr) {
    if(dateStr.includes('/')) {
        const parts = dateStr.split('/');
        if(parts.length === 3) return `${parts[2]}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`;
    }
    return dateStr;
}

function convertToCSVDate(dateStr) {
    if(dateStr.includes('-')) {
        const parts = dateStr.split('-');
        if(parts.length === 3) return `${parts[1]}/${parts[2]}/${parts[0]}`;
    }
    return dateStr;
}

function addNewEmployee() {
    const nameInput = document.getElementById('newEmployeeName'); const name = nameInput.value.trim().toUpperCase();
    if (name === "" || parsedEmployeesData[name]) return;
    parsedEmployeesData[name] = []; const ngayon = new Date();
    parsedEmployeesData[name].push({ date: ngayon.toISOString().split('T')[0], day: DAYS_OF_WEEK[ngayon.getDay()], shift: "PM SHIFT", remarks: "PRESENT", dailyRate: 650, allowance: 0, otHours: 0 });
    nameInput.value = ""; refreshEditorUI(); switchEmployeeTab(name);
    saveAppState();
}

function deleteEmployeeName(empName, event) {
    if (event) event.stopPropagation();
    if (confirm(`Sigurado ka bang nais mong burahin si ${empName}?`)) {
        delete parsedEmployeesData[empName]; const empNames = Object.keys(parsedEmployeesData);
        if (currentActiveEmployee === empName) currentActiveEmployee = empNames.length > 0 ? empNames[0] : "";
        if (empNames.length === 0) {
            // Keep rowActionsContainer and editorActions visible so user can still add employees
            document.getElementById('employeeTabs').innerHTML = "";
            document.getElementById('graphsSection').classList.add('hidden');
            document.getElementById('csvDataTable').classList.add('hidden');
            document.getElementById('computationSummary').classList.add('hidden');
        } else { refreshEditorUI(); }
        saveAppState();
    }
}
function refreshEditorUI() {
    const empNames = Object.keys(parsedEmployeesData); if(empNames.length === 0) return;
    // rowActionsContainer and editorActions are always visible (shown by default in HTML)
    document.getElementById('graphsSection').classList.remove('hidden');
    document.getElementById('graphsSection').classList.add('grid');
    renderEmployeeTabs();
}

function renderEmployeeTabs() {
    const tabsContainer = document.getElementById('employeeTabs'); tabsContainer.innerHTML = "";
    Object.keys(parsedEmployeesData).forEach((name, index) => {
        if (!currentActiveEmployee && index === 0) currentActiveEmployee = name;
        const isActive = currentActiveEmployee === name;
        const wrapper = document.createElement('div');
        wrapper.className = `inline-flex items-center rounded-lg overflow-hidden border ${isActive ? 'bg-orange-600 text-white border-orange-700 font-bold' : 'bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 border-slate-300 dark:border-slate-700 hover:bg-slate-200'}`;
        
        const tabBtn = document.createElement('button'); tabBtn.className = 'px-3 py-1.5 text-xs font-bold cursor-pointer'; tabBtn.innerText = name;
        tabBtn.onclick = () => switchEmployeeTab(name);
        
        const deleteBtn = document.createElement('button'); deleteBtn.className = `px-2 py-1.5 text-xs font-bold cursor-pointer ${isActive ? 'text-orange-200 hover:text-white' : 'text-slate-400 hover:text-red-500'}`;
        deleteBtn.innerHTML = '✕'; deleteBtn.onclick = (e) => deleteEmployeeName(name, e);
        
        wrapper.appendChild(tabBtn); wrapper.appendChild(deleteBtn); tabsContainer.appendChild(wrapper);
    });
    if (currentActiveEmployee) displayEmployeeData(currentActiveEmployee);
}

function switchEmployeeTab(empName) { currentActiveEmployee = empName; renderEmployeeTabs(); displayEmployeeData(empName); }

// 🔴 TINANGGAL ANG PETSA, ARAW, SIPET, AT GINAWANG "DELETE" ANG AKSYON DITO:
function displayEmployeeData(empName) {
    const table = document.getElementById('csvDataTable'); 
    
    let thead = table.querySelector('thead');
    if (!thead) {
        thead = document.createElement('thead');
        table.insertBefore(thead, table.firstChild);
    }
    thead.innerHTML = `
        <tr class="bg-slate-100 dark:bg-slate-900 text-slate-600 dark:text-slate-400 text-xs font-bold uppercase tracking-wider">
            <th class="p-3">Date</th>
            <th class="p-3">Day</th>
            <th class="p-3">Shift</th>
            <th class="p-3">Status / Remarks</th>
            <th class="p-3">Daily Rate (₱)</th>
            <th class="p-3">Allowance (₱)</th>
            <th class="p-3">OT (Hours)</th>
            <th class="p-3 text-center">DELETE</th>
        </tr>
    `;

    const tbody = document.getElementById('csvDataBody'); 
    tbody.innerHTML = ""; 
    
    const isDark = document.documentElement.classList.contains('dark');
    const inputBg = isDark ? "bg-slate-900 text-slate-100 border-slate-700" : "bg-white text-slate-800 border-slate-300";
    let attBuffer = [];
    
    (parsedEmployeesData[empName] || []).forEach((rec, index) => {
        let dayOptions = DAYS_OF_WEEK.map(d => `<option value="${d}" ${rec.day === d ? 'selected' : ''}>${d}</option>`).join('');
        
        attBuffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/40">
            <td class="p-2"><input type="date" class="w-full text-xs p-1.5 ${inputBg} border rounded-lg focus:outline-none focus:ring-1 focus:ring-orange-500" value="${rec.date}" onchange="updateCell('${empName}', ${index}, 'date', this.value)"></td>
            <td class="p-2"><select class="w-full text-xs p-1.5 ${inputBg} border rounded-lg focus:outline-none focus:ring-1 focus:ring-orange-500">${dayOptions}</select></td>
            <td class="p-2"><select class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-bold focus:outline-none focus:ring-1 focus:ring-orange-500" onchange="updateCell('${empName}', ${index}, 'shift', this.value)"><option value="AM SHIFT" ${rec.shift === 'AM SHIFT' ? 'selected' : ''}>AM SHIFT</option><option value="PM SHIFT" ${rec.shift === 'PM SHIFT' ? 'selected' : ''}>PM SHIFT</option></select></td>
            <td class="p-2"><select class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-bold focus:outline-none focus:ring-1 focus:ring-orange-500" onchange="updateRemarksEvent('${empName}', ${index}, this)"><option value="PRESENT" ${rec.remarks === 'PRESENT' ? 'selected' : ''}>PRESENT</option><option value="ABSENT" ${rec.remarks === 'ABSENT' ? 'selected' : ''}>ABSENT</option><option value="RESTDAY" ${rec.remarks === 'RESTDAY' ? 'selected' : ''}>RESTDAY</option><option value="DOUBLE PAY" ${rec.remarks === 'DOUBLE PAY' ? 'selected' : ''}>DOUBLE PAY</option><option value="CDO" ${rec.remarks === 'CDO' ? 'selected' : ''}>CDO</option></select></td>
            <td class="p-2"><input type="number" class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-mono text-slate-700 dark:text-slate-300" value="${rec.dailyRate}" oninput="updateCell('${empName}', ${index}, 'dailyRate', this.value)"></td>
            <td class="p-2"><input type="number" class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-mono font-bold text-emerald-600 dark:text-emerald-400" value="${rec.allowance || 0}" oninput="updateCell('${empName}', ${index}, 'allowance', this.value)"></td>
            <td class="p-2"><input type="number" class="w-full text-xs p-1.5 ${inputBg} border rounded-lg font-mono text-slate-700 dark:text-slate-300" value="${rec.otHours}" oninput="updateCell('${empName}', ${index}, 'otHours', this.value)"></td>
            <td class="p-2 text-center"><button onclick="deleteRow('${empName}', ${index})" class="bg-red-500 text-white font-bold p-1.5 rounded-lg hover:bg-red-600 text-xs transition-colors cursor-pointer shadow-sm">🗑️</button></td>
        </tr>`);
    });
    
    tbody.innerHTML = attBuffer.join('');
    table.classList.remove('hidden'); 
    document.getElementById('computationSummary').classList.remove('hidden');
    calculateTotals(empName); 
    updateVisualGraphs(parsedEmployeesData[empName] || []);
}

function deleteRow(empName, index) { if (confirm("Nais mo bang burahin ang hanay na ito?")) { parsedEmployeesData[empName].splice(index, 1); displayEmployeeData(empName); saveAppState(); } }

// ====================================================================
// ⚙️ SYSTEM CALCULATIONS & GRAPH REFRESHERS
// ====================================================================
function updateCell(empName, index, field, value) {
    if (field === 'dailyRate' || field === 'otHours' || field === 'allowance') { parsedEmployeesData[empName][index][field] = Number(value) || 0; } 
    else { parsedEmployeesData[empName][index][field] = value; }
    if (field === 'date') {
        let d = new Date(value);
        if(!isNaN(d.getTime())) { parsedEmployeesData[empName][index]['day'] = DAYS_OF_WEEK[d.getDay()]; displayEmployeeData(empName); saveAppState(); return; }
    }
    calculateTotals(empName); updateVisualGraphs(parsedEmployeesData[empName]);
    saveAppStateDebounced();
}

function updateRemarksEvent(empName, index, selectElement) {
    const val = selectElement.value; parsedEmployeesData[empName][index]['remarks'] = val;
    if (val === 'PRESENT') parsedEmployeesData[empName][index]['dailyRate'] = 650;
    else if (val === 'DOUBLE PAY') parsedEmployeesData[empName][index]['dailyRate'] = 1250;
    else if (['ABSENT', 'RESTDAY', 'CDO'].includes(val)) { parsedEmployeesData[empName][index]['dailyRate'] = 0; parsedEmployeesData[empName][index]['allowance'] = 0; }
    displayEmployeeData(empName);
    saveAppState();
}

function addNewRowToCurrentEmployee() {
    if(!currentActiveEmployee) return; const records = parsedEmployeesData[currentActiveEmployee];
    let newDate = "2026-06-01", newDay = "MONDAY";
    if(records.length > 0) {
        let d = new Date(records[records.length - 1].date); d.setDate(d.getDate() + 1);
        if(!isNaN(d.getTime())) { newDate = d.toISOString().split('T')[0]; newDay = DAYS_OF_WEEK[d.getDay()]; }
    }
    parsedEmployeesData[currentActiveEmployee].push({ date: newDate, day: newDay, shift: "PM SHIFT", remarks: "PRESENT", dailyRate: 650, allowance: 0, otHours: 0 });
    displayEmployeeData(currentActiveEmployee);
    saveAppState();
}

/** 📅 CALENDAR SET — binabasa ang "Mula" at "Hanggang" na petsa mula sa
 *  mini date-range picker sa tabi ng "+ Bagong Hanay", at kusang gumagawa
 *  ng isang bagong hanay PARA SA BAWAT ARAW sa loob ng saklaw (inclusive),
 *  direktang idinadagdag sa ibaba ng talahanayan ng kasalukuyang aktibong
 *  empleyado. Hindi na kailangang pindutin nang paulit-ulit ang "+ Bagong
 *  Hanay" — sapat na ang pagpili ng saklaw ng petsa. Nililaktawan ang mga
 *  araw na meron na (para walang duplicate), at pinananatiling pasunod-sunod
 *  (chronological) ang mga hanay pagkatapos idagdag. */
function generateRowsFromCalendarSet() {
    if (!currentActiveEmployee) {
        alert("⚠️ Pumili o gumawa muna ng empleyado bago mag-set ng Calendar.");
        return;
    }

    const fromInput = document.getElementById('calendarSetFrom');
    const toInput = document.getElementById('calendarSetTo');
    const fromVal = fromInput.value;
    const toVal = toInput.value;

    if (!fromVal || !toVal) {
        alert("⚠️ Pumili muna ng 'Mula' at 'Hanggang' na petsa sa Calendar Set.");
        return;
    }

    let startDate = new Date(fromVal);
    let endDate = new Date(toVal);
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        alert("⚠️ Hindi valid ang petsang pinili.");
        return;
    }
    if (startDate > endDate) { const tmp = startDate; startDate = endDate; endDate = tmp; } // auto-ayos kung baligtad

    const MAX_DAYS = 366; // proteksyon laban sa aksidenteng napakahabang saklaw (hal. maling taon)
    const totalDaysInRange = Math.round((endDate - startDate) / 86400000) + 1;
    if (totalDaysInRange > MAX_DAYS) {
        alert(`⚠️ Masyadong mahaba ang saklaw (${totalDaysInRange} araw). Hanggang ${MAX_DAYS} araw lang ang pwede sa isang pag-set.`);
        return;
    }

    const records = parsedEmployeesData[currentActiveEmployee];
    const existingDates = new Set(records.map(r => r.date));

    let added = 0, skipped = 0;
    let cursor = new Date(startDate);
    while (cursor <= endDate) {
        const iso = cursor.toISOString().split('T')[0];
        if (existingDates.has(iso)) {
            skipped++;
        } else {
            records.push({ date: iso, day: DAYS_OF_WEEK[cursor.getDay()], shift: "PM SHIFT", remarks: "PRESENT", dailyRate: 650, allowance: 0, otHours: 0 });
            existingDates.add(iso);
            added++;
        }
        cursor.setDate(cursor.getDate() + 1);
    }

    // Panatilihing pasunod-sunod (chronological) ang mga hanay pagkatapos idagdag
    records.sort((a, b) => new Date(a.date) - new Date(b.date));

    displayEmployeeData(currentActiveEmployee);
    saveAppState();

    fromInput.value = "";
    toInput.value = "";

    let msg = added > 0 ? `✅ ${added} bagong hanay ang naidagdag mula sa Calendar Set.` : "ℹ️ Walang bagong naidagdag — meron na ang lahat ng petsa sa saklaw.";
    if (skipped > 0) msg += ` (${skipped} araw nilaktawan, meron na.)`;
    showToast(msg, added === 0);
}

function checkAttendance(status) {
    const nameInput = document.getElementById('employeeName'); const name = nameInput.value.trim().toUpperCase();
    if (name === "") { alert("Pakiusap, ilagay ang iyong pangalan."); return; }

    const defaultTime = new Date().toTimeString().substring(0, 5);
    _manualLogIdSeq++;
    manualAttendanceLogs.unshift({
        id: _manualLogIdSeq,
        name: name,
        date: new Date().toLocaleDateString('en-PH'),
        time: defaultTime,
        status: status
    });
    manualLogsCounter++;

    renderManualLogs();
    nameInput.value = "";
    saveAppState();
}

// Muling ginuhit ang buong Live Log Feed mula sa manualAttendanceLogs array.
// Ito ang gumagawang posible na ma-restore ang feed kapag nag-reload ang page.
function renderManualLogs() {
    const tbody = document.getElementById('attendanceLog');
    if (!tbody) return;

    if (manualAttendanceLogs.length === 0) {
        tbody.innerHTML = `<tr id="noRecordRow"><td colspan="5" class="py-12 text-center text-slate-400 dark:text-slate-500">Walang manual record sa ngayon.</td></tr>`;
        document.getElementById('totalLogs').innerText = manualLogsCounter;
        updateActiveDashboardCardDirectly();
        return;
    }

    const isDark = document.documentElement.classList.contains('dark');
    const selectBg = isDark ? "bg-slate-900 border-slate-700 text-slate-100" : "bg-white border-slate-300 text-slate-800";
    let buffer = [];
    manualAttendanceLogs.forEach(log => {
        buffer.push(`<tr class="hover:bg-slate-100 dark:hover:bg-slate-800/30 transition-colors">
            <td class="p-2 font-bold text-slate-800 dark:text-slate-200">${log.name}</td>
            <td class="p-2 text-slate-600 dark:text-slate-400">${log.date}</td>
            <td class="p-2"><input type="time" class="${selectBg} border text-xs p-1 rounded" value="${log.time}" onchange="updateManualLogField(${log.id}, 'time', this.value)"></td>
            <td class="p-2"><select class="${selectBg} border text-xs p-1 rounded" onchange="updateManualLogField(${log.id}, 'status', this.value)"><option value="In" ${log.status === 'In' ? 'selected' : ''}>Time In</option><option value="Out" ${log.status === 'Out' ? 'selected' : ''}>Time Out</option></select></td>
            <td class="p-2 text-center"><button onclick="deleteManualLogRow(${log.id})" class="bg-red-500 text-white p-1 rounded text-[10px] cursor-pointer">🗑️</button></td>
        </tr>`);
    });

    tbody.innerHTML = buffer.join('');
    document.getElementById('totalLogs').innerText = manualLogsCounter;
    updateActiveDashboardCardDirectly();
}

// Para sa pag-edit ng time o status sa isang manual log row
function updateManualLogField(id, field, value) {
    const log = manualAttendanceLogs.find(l => l.id === id);
    if (log) { log[field] = value; }
    if (field === 'status') updateActiveDashboardCardDirectly();
    saveAppStateDebounced();
}

function deleteManualLogRow(id) {
    if (confirm("Nais mo bang burahin ang manual log na ito?")) {
        manualAttendanceLogs = manualAttendanceLogs.filter(l => l.id !== id);
        manualLogsCounter = Math.max(0, manualLogsCounter - 1);
        renderManualLogs();
        saveAppState();
    }
}

function updateActiveDashboardCardDirectly() {
    let activeInCount = 0; document.querySelectorAll('#attendanceLog select').forEach(sel => { if(sel.value === 'In') activeInCount++; });
    document.getElementById('activeUsers').innerText = activeInCount;
}

function calculateTotals(empName) {
    const records = parsedEmployeesData[empName] || []; 
    let daysPresent = 0, totalBasic = 0, totalAllowance = 0, totalOTHours = 0, estimatedOTPay = 0;
    
    records.forEach(rec => {
        if (rec.remarks === "PRESENT" || rec.remarks === "DOUBLE PAY") daysPresent++;
        totalBasic += Number(rec.dailyRate) || 0; totalAllowance += Number(rec.allowance) || 0; totalOTHours += Number(rec.otHours) || 0;
        if ((Number(rec.otHours) || 0) > 0) {
            let r = String(rec.shift).toUpperCase().includes("AM") ? 93 : 101;
            estimatedOTPay += r * (Number(rec.otHours) || 0);
        }
    });
    
    let grossSalary = totalBasic + totalAllowance + estimatedOTPay;
    document.getElementById('sumDaysPresent').innerText = daysPresent;
    document.getElementById('sumBasicSalary').innerText = "₱" + totalBasic.toLocaleString('en-US', {minimumFractionDigits: 2});
    document.getElementById('sumOTHours').innerText = totalOTHours + " hrs";
    document.getElementById('sumGrossSalary').innerText = "₱" + grossSalary.toLocaleString('en-US', {minimumFractionDigits: 2});
}

function updateVisualGraphs(records) {
    let remarksCount = { "PRESENT": 0, "ABSENT": 0, "RESTDAY": 0, "DOUBLE PAY": 0, "CDO": 0 };
    let labelsBar = [], dataRateBar = [], dataOTBar = [];
    records.forEach(rec => {
        if (remarksCount[rec.remarks] !== undefined) remarksCount[rec.remarks]++;
        labelsBar.push(rec.date.substring(5));
        dataRateBar.push(rec.dailyRate + (rec.allowance || 0));
        dataOTBar.push(rec.otHours);
    });

    const isDark = document.documentElement.classList.contains('dark');
    const lblColor = isDark ? '#94a3b8' : '#475569';

    // ✅ HIGHCHARTS 3D PIE — same style as dashboard pie chart
    const REMARK_COLORS = {
        "PRESENT":    '#10b981',  // green
        "ABSENT":     '#ef4444',  // red
        "RESTDAY":    '#f59e0b',  // amber
        "DOUBLE PAY": '#3b82f6',  // blue
        "CDO":        '#a855f7'   // purple
    };
    const REMARK_ICONS = {
        "PRESENT":    '✅',
        "ABSENT":     '❌',
        "RESTDAY":    '🌙',
        "DOUBLE PAY": '💰',
        "CDO":        '📋'
    };

    const pieData = Object.entries(remarksCount)
        .filter(([, count]) => count > 0)
        .map(([label, count]) => ({
            name: REMARK_ICONS[label] + ' ' + label,
            y: count,
            color: REMARK_COLORS[label]
        }));

    Highcharts.chart('percentageChart', {
        chart: {
            type: 'pie',
            options3d: { enabled: true, alpha: 45, beta: 0 },
            backgroundColor: 'transparent',
            margin: [30, 10, 60, 10],
            spacing: [4, 4, 4, 4]
        },
        title: {
            text: '📊 Attendance Breakdown',
            style: { fontSize: '11px', fontWeight: 'bold', color: lblColor },
            margin: 8
        },
        plotOptions: {
            pie: {
                depth: 35,
                allowPointSelect: true,
                cursor: 'pointer',
                center: ['50%', '48%'],
                size: '72%',
                dataLabels: {
                    enabled: true,
                    distance: 22,
                    formatter: function() {
                        return '<b>' + this.point.name + '</b><br>' + this.percentage.toFixed(0) + '%';
                    },
                    style: {
                        fontSize: '8px',
                        fontWeight: 'bold',
                        color: lblColor,
                        textOutline: 'none'
                    },
                    connectorWidth: 1,
                    connectorColor: lblColor
                },
                showInLegend: false
            }
        },
        credits: { enabled: false },
        series: [{ name: 'Days', data: pieData }]
    });

    // ✅ Bar chart stays using Chart.js (income + OT trend)
    if (chartInstanceBar3D) chartInstanceBar3D.destroy();
    const axisColor = isDark ? '#94a3b8' : '#475569';
    chartInstanceBar3D = new Chart(document.getElementById('trend3dChart').getContext('2d'), {
        type: 'bar',
        data: {
            labels: labelsBar,
            datasets: [
                { label: 'Income Base (₱)', data: dataRateBar, backgroundColor: '#3b82f6' },
                { label: 'OT Hours', data: dataOTBar, backgroundColor: '#f59e0b', yAxisID: 'y1' }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y:  { position: 'left',  ticks: { color: axisColor } },
                y1: { position: 'right', grid: { drawOnChartArea: false }, ticks: { color: axisColor } },
                x:  { ticks: { color: axisColor } }
            },
            plugins: { legend: { labels: { color: axisColor } } }
        }
    });
}

// ====================================================================
// 📥 EXPORT EXTENSION UTILITIES
// ====================================================================
function exportToCSV() {
    const empNames = Object.keys(parsedEmployeesData); if(empNames.length === 0) return;
    let maxRows = 0; empNames.forEach(name => { if(parsedEmployeesData[name].length > maxRows) maxRows = parsedEmployeesData[name].length; });
    let csvLines = [];
    
    let r1 = []; empNames.forEach(n => r1.push(`${n} ATTENDANCE`, "", "", "", "", "", "", "")); csvLines.push(r1.join(','));
    let r2 = []; empNames.forEach(() => r2.push("UPDATED ATTENDANCE RECORD", "", "", "", "", "", "", "")); csvLines.push(r2.join(','));
    let r3 = []; empNames.forEach(() => r3.push("Date", "Day", "Shift", "Remarks", "Daily rate", "Allowance", "Number of OT", "")); csvLines.push(r3.join(','));

    for(let j = 0; j < maxRows; j++) {
        let dRow = [];
        empNames.forEach(name => {
            const rec = parsedEmployeesData[name][j];
            if(rec) dRow.push(convertToCSVDate(rec.date), rec.day, rec.shift, rec.remarks, rec.dailyRate, (rec.allowance || 0), rec.otHours, "");
            else dRow.push("", "", "", "", "", "", "", "");
        });
        csvLines.push(dRow.join(',')); 
    }

    const blob = new Blob([csvLines.join("\n")], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob); const link = document.createElement("a");
    link.setAttribute("href", url); link.setAttribute("download", "MTO_UPDATED_ATTENDANCE.csv");
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
}

function handleDashboardExport() {
    if (!globalData || globalData.length === 0) { alert("Walang data na pwedeng i-export sa ngayon. Pakiusap, mag-upload muna ng file."); return; }
    let csvLines = []; let fileName = "Dashboard_Export.csv";
    if (fileType === "progress") {
        csvLines.push("Work Week,Month,Warehouse,MTO Batch Name,Forecast,Allocated,Total Shipped,Discrepancy,Percentage");
        globalData.forEach(row => { csvLines.push(`"${row.ww}","${row.month}","${row.wh}","${row.batch}",${row.forecast},${row.allocated},${row.shipped},${row.discrepancy},${row.percentage}%`); });
        fileName = "MTO_Progress_Matrix_2026.csv";
    } else if (fileType === "raw") {
        csvLines.push("Box / MTB ID,Type,Series / Serial Number,Destination Warehouse,Qty,Status,Shipped Date,Remarks");
        globalData.forEach(row => { csvLines.push(`"${row.mtb}","${row.type}","${row.series}","${row.wh}",${row.qty},"${row.status}","${row.date}","${row.remarks || ''}"`); });
        fileName = "MTO_Raw_Pallet_Tracking.csv";
    } else if (fileType === "crossdock") {
        csvLines.push("Date,MTB ID,SKU Name,Req Qty,Act Qty,Pallet,Checked By,Status");
        globalData.forEach(row => { csvLines.push(`"${row.date}","${row.mtb}","${row.skuName}",${row.qty},${row.actualQty},"${row.pallet}","${row.checkedBy}","${row.statusReceived}"`); });
        fileName = "CrossDock_Transmittal_Hub.csv";
    } else if (fileType === "device") {
        csvLines.push("Device ID,Device Type,Specification(Rows),Specification(Columns),Status,Device Progress,Template ID");
        globalData.forEach(row => { csvLines.push(`"${row.deviceId}","${row.deviceType}","${row.specRows}","${row.specCols}","${row.status}","${row.progress}","${row.templateId}"`); });
        fileName = "Device_Management_Export.csv";
    } else if (fileType === "generic") {
        csvLines.push(genericHeaders.join(","));
        globalData.forEach(row => {
            let line = genericHeaders.map(h => `"${row[h] || ''}"`).join(",");
            csvLines.push(line);
        });
        fileName = "Custom_Dashboard_Export.csv";
    } else if (fileType === "cancelled") {
        csvLines.push("WHs,MT Order,Picking ID,SKU,SKU Name,Qty,PICKED QTY,CHECKED QTY,LACK ITEM,Reason,Picker");
        globalData.forEach(row => {
            csvLines.push(`"${row.whs}","${row.mtOrder}","${row.pickingId}","${row.sku}","${row.skuName.replace(/"/g,"'")}",${row.qty || 0},${row.pickedQty || ''},${row.checkedQty || ''},${row.lackItem || ''},"${row.reason}","${row.picker}"`);
        });
        fileName = "MTO_Cancelled_Monitoring_Export.csv";
    }
    
    const blob = new Blob(["\ufeff" + csvLines.join("\n")], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob); const link = document.createElement("a");
    link.setAttribute("href", url); link.setAttribute("download", fileName);
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
}