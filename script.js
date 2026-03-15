// ============================================
// 履歴管理 (Undo/Redo)
// ============================================
const HistoryManager = {
    stack: [],
    currentIndex: -1,
    limit: 30,
    isRestoring: false,

    init(initialState) {
        this.stack = [JSON.stringify(initialState)];
        this.currentIndex = 0;
        this.updateButtons();
    },

    record(state) {
        if (this.isRestoring) return;
        // 現在位置より先の履歴は破棄
        if (this.currentIndex < this.stack.length - 1) {
            this.stack = this.stack.slice(0, this.currentIndex + 1);
        }
        // 新しい状態を追加
        const json = JSON.stringify(state);
        // 直前と同じなら保存しない
        if (this.stack[this.currentIndex] === json) return;

        this.stack.push(json);
        if (this.stack.length > this.limit) {
            this.stack.shift();
        } else {
            this.currentIndex++;
        }
        this.updateButtons();
    },

    undo() {
        if (this.currentIndex > 0) {
            this.currentIndex--;
            this.performRestore();
        }
    },

    redo() {
        if (this.currentIndex < this.stack.length - 1) {
            this.currentIndex++;
            this.performRestore();
        }
    },

    performRestore() {
        this.isRestoring = true;
        const data = JSON.parse(this.stack[this.currentIndex]);
        
        const scrollContainer = document.querySelector(".gantt-scroll-container");
        const savedScrollLeft = scrollContainer ? scrollContainer.scrollLeft : 0;
        const savedScrollTop = scrollContainer ? scrollContainer.scrollTop : 0;

        restoreFromData(data); 
        DataManager.save(data);

        if (scrollContainer) {
            scrollContainer.scrollLeft = savedScrollLeft;
            scrollContainer.scrollTop = savedScrollTop;
        }

        this.isRestoring = false;
        this.updateButtons();
    },

    updateButtons() {
        const undoBtn = document.getElementById("undoBtn");
        const redoBtn = document.getElementById("redoBtn");
        if(undoBtn) undoBtn.disabled = (this.currentIndex <= 0);
        if(redoBtn) redoBtn.disabled = (this.currentIndex >= this.stack.length - 1);
    }
};

// ============================================
// データ保存管理
// ============================================
const DataManager = {
    dbName: "GanttAppDB",
    storeName: "appData",
    useLocalStorage: false,
    PLAN_KEY: "main",

    async init() {
        try {
            await new Promise((resolve, reject) => {
                const req = indexedDB.open(this.dbName, 1);
                req.onupgradeneeded = (e) => {
                    const db = e.target.result;
                    if (!db.objectStoreNames.contains(this.storeName)) {
                        db.createObjectStore(this.storeName, { keyPath: "id" });
                    }
                };
                req.onsuccess = (e) => { this.db = e.target.result; resolve(); };
                req.onerror = (e) => reject(e);
            });
        } catch (err) {
            console.warn("IndexedDB fallback -> LocalStorage");
            this.useLocalStorage = true;
        }
    },

    async load() {
        if (this.useLocalStorage) {
            const json = localStorage.getItem(this.dbName + "_" + this.PLAN_KEY);
            return json ? JSON.parse(json) : null;
        }
        return new Promise((resolve) => {
            const tx = this.db.transaction([this.storeName], "readonly");
            const req = tx.objectStore(this.storeName).get(this.PLAN_KEY);
            req.onsuccess = (e) => resolve(e.target.result ? e.target.result.data : null);
            req.onerror = () => resolve(null);
        });
    },

    async save(data) {
        const ind = document.getElementById("statusIndicator");
        ind.style.opacity = 1;
        setTimeout(() => ind.style.opacity = 0, 1500);

        if (this.useLocalStorage) {
            localStorage.setItem(this.dbName + "_" + this.PLAN_KEY, JSON.stringify(data));
            return;
        }
        return new Promise((resolve) => {
            const tx = this.db.transaction([this.storeName], "readwrite");
            tx.objectStore(this.storeName).put({ id: this.PLAN_KEY, data: data });
            tx.oncomplete = () => resolve();
        });
    }
};

// ============================================
// アプリ状態・定数
// ============================================
const CELL_WIDTH = 28;
const MAIN_LINE_Y = 32;
const MAIN_DIVIDER_Y = 66;
const SUB_SCHEDULE_TOP = 103;
const BASE_ROW_HEIGHT = 129;
const SEGMENT_OFFSET_Y = 48; 

const now = new Date();
const todayISO = dateToISO(now);
const defaultStart = new Date(now.getFullYear(), now.getMonth(), 1);
const defaultEnd = new Date(now.getFullYear(), now.getMonth() + 3, 0);
const DEFAULT_TODO_COLUMNS = "項目1, 項目2, 時間, 実施内容, 計画, 実績, メモ";
const LEGACY_TODO_COLUMNS = "項目1, 項目2, 時間, 実施内容, 計画, 実績";

let appData = {
    projectName: "標準の計画",
    settings: {
        startDate: dateToISO(defaultStart),
        endDate: dateToISO(defaultEnd),
        holidays: []
    },
    headers: ["項目1", "項目2", "時間"], 
    todoColumns: DEFAULT_TODO_COLUMNS,
    columnWidths: [30, 120, 90, 40],
    tasks: [],
    memo: ""
};

let timelineDays = [];
let taskObjects = [];
let activeTaskId = null;
let activeProgressSegmentId = null; 
let activeProgressTaskId = null;
let selectionMode = 0; 
let isCtrlSelectionMode = false;
let activeDailyValueTarget = null;
let suppressNextClickAfterProgressCancel = false;
let suppressNextClickAfterMenuDismiss = false;

let currentTodoDate = new Date();
let todoSelectionState = false; 

let dragState = {
    isDragging: false,
    type: null,
    taskId: null,
    segId: null,
    milestoneId: null,
    selectedSegRefs: [],
    startX: 0,
    originalLeft: 0,
    originalWidth: 0,
    originalStartDate: null,
    originalEndDate: null,
    el: null
};

// DOM要素
const headerRow = document.getElementById("headerRow");
const rowsContainer = document.getElementById("rowsContainer");
const leftRowsContainer = document.getElementById("leftRows");
const rangeLabel = document.getElementById("rangeLabel");
const ganttRight = document.getElementById("ganttRight");
const freeMemo = document.getElementById("freeMemo");
const showHiddenCheck = document.getElementById("showHiddenCheck");
const projectNameInput = document.getElementById("projectNameInput");

const contextMenu = document.getElementById("contextMenu");
let contextMenuTargetTaskId = null;

const segmentContextMenu = document.getElementById("segmentContextMenu");
let contextMenuTargetSegId = null;
let contextMenuTargetTaskForSeg = null;

const settingsPanel = document.getElementById("settingsPanel");
const totalRow = document.getElementById("totalRow");
const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

const taskMemoPanel = document.getElementById("taskMemoPanel");
const taskMemoHeader = document.getElementById("taskMemoHeader");
const taskMemoTitle = document.getElementById("taskMemoTitle");
const taskMemoTextarea = document.getElementById("taskMemoTextarea");
const taskMemoCount = document.getElementById("taskMemoCount");
const taskMemoClose = document.getElementById("taskMemoClose");
let memoPanelTaskId = null;
let memoPanelPinned = false;

let leftColumnWidths = [30, 120, 90, 40];
let isResizingCol = false;
let resizeColIndex = null;
let resizeStartX = 0;
let resizeStartWidth = 0;

let isResizingRow = false;
let resizeRowTaskId = null;
let resizeStartY = 0;
let resizeStartHeight = 0;

// ============================================
// ヘルパー関数
// ============================================
function pad2(n) { return String(n).padStart(2, "0"); }
function dateToISO(d) { return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate()); }
function isoToDate(iso) { const [y, m, d] = iso.split("-").map(Number); return new Date(y, m - 1, d); }
function shiftDateStr(str, delta) {
    const d = isoToDate(str);
    d.setDate(d.getDate() + delta);
    return dateToISO(d);
}
function formatTimestamp(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const h = String(d.getHours()).padStart(2, '0');
    const min = String(d.getMinutes()).padStart(2, '0');
    return `${y}${m}${day}${h}${min}`;
}
function dateToIndex(str) { return timelineDays.findIndex((d) => d.iso === str); }
function centerX(index) { return index * CELL_WIDTH + CELL_WIDTH / 2; }
function getByteLength(str) {
    let len = 0;
    for (let i = 0; i < str.length; i++) {
        const c = str.charCodeAt(i);
        len += ((c >= 0x0 && c <= 0x7f) || (c >= 0xff61 && c <= 0xff9f)) ? 1 : 2;
    }
    return len;
}

function applyLeftColumnWidths() {
    const cols = leftColumnWidths.map(w => `${w}px`).join(" ");
    const header = document.querySelector(".left-header");
    if (header) header.style.gridTemplateColumns = cols;
    const rows = document.querySelectorAll(".left-row");
    rows.forEach(r => {
        r.style.gridTemplateColumns = cols;
        const handle = r.querySelector(".row-resize-handle");
        if (handle) handle.style.left = `${leftColumnWidths[0]}px`;
    });
    const left = document.querySelector(".gantt-left");
    if (left) {
        const total = leftColumnWidths.reduce((a, b) => a + b, 0);
        left.style.flex = `0 0 ${total}px`;
        left.style.width = `${total}px`;
    }
    updateLeftResizeHandles();
}

function measureMinWidthForHeader(index) {
    const ids = ["rowSelectHeader", "lh1", "lh2", "lh3"];
    const el = document.getElementById(ids[index]);
    if (!el) return 40;
    const text = el.textContent || "";
    const probe = document.createElement("span");
    probe.style.position = "absolute";
    probe.style.visibility = "hidden";
    probe.style.whiteSpace = "nowrap";
    probe.style.fontSize = "11px";
    probe.style.fontWeight = "600";
    probe.textContent = text;
    document.body.appendChild(probe);
    const width = probe.getBoundingClientRect().width;
    document.body.removeChild(probe);
    return Math.ceil(width + 16);
}

function clampColumnWidth(index, width) {
    const min = (index === 0) ? 26 : measureMinWidthForHeader(index);
    const max = (index === 0) ? 60 : 420;
    return Math.max(min, Math.min(max, width));
}

function ensureLeftResizeOverlay() {
    const left = document.querySelector(".gantt-left");
    if (!left) return null;
    let overlay = left.querySelector(".left-resize-overlay");
    if (!overlay) {
        overlay = document.createElement("div");
        overlay.className = "left-resize-overlay";
        left.appendChild(overlay);
    }
    return overlay;
}

function updateLeftResizeHandles() {
    const overlay = ensureLeftResizeOverlay();
    if (!overlay) return;
    overlay.innerHTML = "";
    let acc = 0;
    for (let i = 0; i < leftColumnWidths.length; i++) {
        acc += leftColumnWidths[i];
        const handle = document.createElement("div");
        handle.className = "left-resize-handle";
        handle.style.left = `${acc - 4}px`;
        handle.dataset.colIndex = String(i);
        handle.addEventListener("mousedown", (e) => {
            e.preventDefault();
            e.stopPropagation();
            isResizingCol = true;
            resizeColIndex = i;
            resizeStartX = e.clientX;
            resizeStartWidth = leftColumnWidths[i];
            document.body.style.cursor = "col-resize";
            document.body.style.userSelect = "none";
        });
        overlay.appendChild(handle);
    }
}

function computeTaskBaseHeight(task) {
    if (!timelineDays.length) return BASE_ROW_HEIGHT;
    if (!task.segments || task.segments.length === 0) return BASE_ROW_HEIGHT;

    const taskDates = {};
    task.segments.forEach(seg => seg._lane = 0);
    const sortedSegs = [...task.segments].sort((a, b) => (a.startDate !== b.startDate) ? (a.startDate < b.startDate ? -1 : 1) : (a.endDate < b.endDate ? -1 : 1));
    let maxLaneUsed = 0;
    sortedSegs.forEach(seg => {
        let requiredLane = 0;
        let sIdx = dateToIndex(seg.startDate);
        let eIdx = dateToIndex(seg.endDate);
        if (sIdx === -1 || eIdx === -1) return;
        while (true) {
            let overlap = false;
            for (let i = Math.min(sIdx, eIdx); i <= Math.max(sIdx, eIdx); i++) {
                const iso = timelineDays[i].iso;
                if (taskDates[iso] && taskDates[iso].includes(requiredLane)) { overlap = true; break; }
            }
            if (!overlap) break;
            requiredLane++;
        }
        seg._lane = requiredLane;
        maxLaneUsed = Math.max(maxLaneUsed, requiredLane);
        for (let i = Math.min(sIdx, eIdx); i <= Math.max(sIdx, eIdx); i++) {
            const iso = timelineDays[i].iso;
            if (!taskDates[iso]) taskDates[iso] = [];
            taskDates[iso].push(requiredLane);
        }
    });
    const laneCount = maxLaneUsed + 1;
    return Math.max(BASE_ROW_HEIGHT, SUB_SCHEDULE_TOP + ((laneCount - 1) * SEGMENT_OFFSET_Y) + 26);
}

function getSubScheduleLaneFromY(y) {
    const relativeY = y - SUB_SCHEDULE_TOP;
    if (relativeY <= 0) return 0;
    return Math.max(0, Math.floor((relativeY + (SEGMENT_OFFSET_Y / 2)) / SEGMENT_OFFSET_Y));
}

function normalizeMainSchedule(mainSchedule) {
    if (!mainSchedule || !mainSchedule.startDate || !mainSchedule.endDate) return null;
    return {
        startDate: mainSchedule.startDate,
        endDate: mainSchedule.endDate,
        label: mainSchedule.label || "メイン計画",
        startLabel: mainSchedule.startLabel || "",
        endLabel: mainSchedule.endLabel || "",
        progressEndDate: mainSchedule.progressEndDate || null,
        milestones: Array.isArray(mainSchedule.milestones)
            ? mainSchedule.milestones.map(ms => ({
                id: ms.id || ("ms_" + Date.now() + "_" + Math.random().toString(36).slice(2)),
                date: ms.date,
                label: ms.label || "マイルストーン"
            }))
            : []
    };
}

function getTaskTitle(task) {
    const t1 = task.leftRowEl.children[1].firstElementChild.textContent.trim();
    const t2 = task.leftRowEl.children[2].firstElementChild.textContent.trim();
    const t3 = task.leftRowEl.children[3].firstElementChild.textContent.trim();
    const parts = [t1, t2, t3].filter(Boolean);
    if (parts.length === 0) return "メモ";
    return "メモ: " + parts.join(" / ");
}

function updateMemoCount(text) {
    const len = text.length;
    taskMemoCount.textContent = `${len} / 3000`;
}

function positionMemoPanel(anchorEl) {
    const rect = anchorEl.getBoundingClientRect();
    const panelRect = taskMemoPanel.getBoundingClientRect();
    const padding = 12;
    let left = rect.right + 10;
    let top = rect.top;
    if (left + panelRect.width > window.innerWidth - padding) {
        left = rect.left - panelRect.width - 10;
    }
    if (left < padding) left = padding;
    if (top + panelRect.height > window.innerHeight - padding) {
        top = window.innerHeight - panelRect.height - padding;
    }
    if (top < padding) top = padding;
    taskMemoPanel.style.left = `${left}px`;
    taskMemoPanel.style.top = `${top}px`;
}

function openTaskMemo(task, anchorEl, pin = false) {
    if (!task || !taskMemoPanel) return;
    memoPanelTaskId = task.id;
    memoPanelPinned = pin;
    taskMemoTitle.textContent = getTaskTitle(task);
    const memoText = task.memo || "";
    taskMemoTextarea.value = memoText;
    updateMemoCount(memoText);
    taskMemoPanel.classList.remove("memo-hidden");
    taskMemoPanel.setAttribute("aria-hidden", "false");
    positionMemoPanel(anchorEl || task.leftRowEl);
}

function closeTaskMemoPanel() {
    memoPanelTaskId = null;
    memoPanelPinned = false;
    taskMemoPanel.classList.add("memo-hidden");
    taskMemoPanel.setAttribute("aria-hidden", "true");
}

function clearSegmentSelection() {
    taskObjects.forEach(task => {
        task.segments.forEach(seg => {
            seg.isSelected = false;
        });
    });
}

function getSelectedSubSegments() {
    const refs = [];
    taskObjects.forEach(task => {
        task.segments.forEach(seg => {
            if (seg.isSelected) refs.push({ task, seg });
        });
    });
    return refs;
}

function updateCtrlSelectionMode(enabled) {
    isCtrlSelectionMode = enabled;
    document.body.classList.toggle("ctrl-select-mode", enabled);
}

function isSubSegmentSelectableTarget(target) {
    return !!target.closest("[data-sub-selectable='true']");
}

function cancelActiveProgressSelection() {
    if (!activeProgressSegmentId) return false;
    activeProgressSegmentId = null;
    activeProgressTaskId = null;
    suppressNextClickAfterProgressCancel = true;
    renderAllSegments();
    return true;
}

function hideContextMenus() {
    contextMenu.style.display = "none";
    segmentContextMenu.style.display = "none";
}

function dismissContextMenusWithClickSuppression() {
    hideContextMenus();
    suppressNextClickAfterMenuDismiss = true;
}

function isOverlayCancelStateActive() {
    return segmentContextMenu.style.display === "block" ||
        contextMenu.style.display === "block" ||
        !!activeProgressSegmentId;
}

function findTaskForActiveProgressSelection() {
    if (!activeProgressSegmentId) return null;
    if (activeProgressTaskId) {
        return taskObjects.find(task => task.id === activeProgressTaskId) || null;
    }
    return taskObjects.find(task => task.segments.some(seg => seg.id === activeProgressSegmentId)) || null;
}

function isValidProgressClickTarget(target) {
    const task = findTaskForActiveProgressSelection();
    if (!task) return false;
    if (!task.rowEl.contains(target)) return false;
    if (!target.closest(".cell")) return false;
    if (target.closest(".segment") || target.closest(".point") || target.closest(".segment-label") || target.closest(".daily-val")) {
        return false;
    }
    return true;
}

function setActiveDailyValueTarget(task, seg, iso) {
    activeDailyValueTarget = {
        taskId: task.id,
        segId: seg.id,
        iso
    };
}

function clearActiveDailyValueTarget() {
    activeDailyValueTarget = null;
}

function getActiveDailyValueRef() {
    if (!activeDailyValueTarget) return null;
    const task = taskObjects.find(t => t.id === activeDailyValueTarget.taskId);
    const seg = task ? task.segments.find(s => s.id === activeDailyValueTarget.segId) : null;
    if (!task || !seg) return null;
    return { task, seg, iso: activeDailyValueTarget.iso };
}

function moveActiveDailyValue(delta) {
    const ref = getActiveDailyValueRef();
    if (!ref) return;
    const nextIso = shiftDateStr(ref.iso, delta);
    if (nextIso < ref.seg.startDate || nextIso > ref.seg.endDate) return;
    if (dateToIndex(nextIso) === -1) return;
    activeDailyValueTarget.iso = nextIso;
    renderAllSegments();
}

function editDailyValue(task, seg, iso) {
    setActiveDailyValueTarget(task, seg, iso);
    const curVal = (seg.dailyValues && seg.dailyValues[iso]) || "";
    let input = prompt("工数 (例: 1, 0.5) または文字:", curVal);
    if (input !== null) {
        input = input.trim();
        if (input === "") {
            if (seg.dailyValues) delete seg.dailyValues[iso];
        } else {
            if (!seg.dailyValues) seg.dailyValues = {};
            if (/^\d(\.\d)?$/.test(input) || (!isNaN(parseFloat(input)) && getByteLength(input) <= 4)) {
                seg.dailyValues[iso] = input;
            } else {
                if (getByteLength(input) <= 4) seg.dailyValues[iso] = input;
                else {
                    alert("全角2文字(半角4文字)以内で入力してください。");
                    renderAllSegments();
                    return;
                }
            }
        }
        renderAllSegments();
        triggerSave();
    } else {
        renderAllSegments();
    }
}

// ============================================
// データ同期 & 保存
// ============================================
function syncDataModel() {
    appData.tasks = taskObjects.map(t => {
        return {
            id: t.id,
            label1: t.leftRowEl.children[1].firstElementChild.textContent,
            label2: t.leftRowEl.children[2].firstElementChild.textContent,
            label3: t.leftRowEl.children[3].firstElementChild.textContent,
            mainSchedule: t.mainSchedule ? {
                startDate: t.mainSchedule.startDate,
                endDate: t.mainSchedule.endDate,
                label: t.mainSchedule.label || "",
                startLabel: t.mainSchedule.startLabel || "",
                endLabel: t.mainSchedule.endLabel || "",
                progressEndDate: t.mainSchedule.progressEndDate || null,
                milestones: (t.mainSchedule.milestones || []).map(ms => ({
                    id: ms.id,
                    date: ms.date,
                    label: ms.label || ""
                }))
            } : null,
            segments: t.segments.map(seg => ({
                id: seg.id,
                startDate: seg.startDate,
                endDate: seg.endDate,
                type: seg.type,
                label: seg.label || "",
                progressEndDate: seg.progressEndDate || null,
                dailyValues: seg.dailyValues ? { ...seg.dailyValues } : {},
                dailyResults: seg.dailyResults ? { ...seg.dailyResults } : {}
            })),
            memo: t.memo || "",
            customHeight: t.customHeight || 0,
            isDone: t.isDone || false,
            isHidden: t.isHidden || false
        };
    });
    appData.memo = freeMemo.innerHTML;
    appData.projectName = projectNameInput.value;
    
    appData.headers = [
        document.getElementById("lh1").textContent,
        document.getElementById("lh2").textContent,
        document.getElementById("lh3").textContent
    ];
    appData.todoColumns = document.getElementById("todoColumnsInput").value;
    appData.columnWidths = leftColumnWidths.slice();
}

function triggerSave() {
    syncDataModel();
    calculateTotals();
    HistoryManager.record(appData);
    DataManager.save(appData);
}

projectNameInput.addEventListener("change", () => {
    document.title = projectNameInput.value + " | 工程表";
    triggerSave();
});
freeMemo.addEventListener("input", () => {
    triggerSave();
});
["lh1", "lh2", "lh3"].forEach(id => {
    document.getElementById(id).addEventListener("blur", () => {
        const idx = id === "lh1" ? 1 : (id === "lh2" ? 2 : 3);
        leftColumnWidths[idx] = clampColumnWidth(idx, leftColumnWidths[idx]);
        applyLeftColumnWidths();
        triggerSave();
    });
});

// ============================================
// 初期化 & 復元
// ============================================
async function initializeApp() {
    await DataManager.init();
    const savedData = await DataManager.load();
    if (savedData) {
        restoreFromData(savedData);
        setTimeout(scrollToToday, 100);
    } else {
        document.getElementById("lh1").textContent = appData.headers[0];
        document.getElementById("lh2").textContent = appData.headers[1];
        document.getElementById("lh3").textContent = appData.headers[2];
        document.getElementById("todoColumnsInput").value = appData.todoColumns;
        leftColumnWidths = appData.columnWidths.slice();
        applyLeftColumnWidths();
        buildTimeline();
        buildHeader();
        addTaskRow();
        setTimeout(scrollToToday, 100);
        document.title = appData.projectName + " | 工程表";
    }
    HistoryManager.init(appData);
    setupControlEvents();
}

function restoreFromData(data) {
    appData = data;
    if (!data.settings.startDate) {
        appData.settings.startDate = dateToISO(defaultStart);
        appData.settings.endDate = dateToISO(defaultEnd);
    }
    if (!appData.headers) appData.headers = ["項目1", "項目2", "時間"];
    if (!appData.columnWidths) appData.columnWidths = [30, 120, 90, 40];
    
    if (!appData.todoColumns || appData.todoColumns === LEGACY_TODO_COLUMNS) {
        appData.todoColumns = DEFAULT_TODO_COLUMNS;
    }

    projectNameInput.value = data.projectName || "標準の計画";
    document.title = projectNameInput.value + " | 工程表";
    freeMemo.innerHTML = data.memo || "";

    document.getElementById("lh1").textContent = appData.headers[0];
    document.getElementById("lh2").textContent = appData.headers[1];
    document.getElementById("lh3").textContent = appData.headers[2];

    document.getElementById("todoColumnsInput").value = appData.todoColumns;

    leftRowsContainer.innerHTML = "";
    rowsContainer.innerHTML = "";
    taskObjects = [];

    buildTimeline();
    buildHeader();
    leftColumnWidths = appData.columnWidths.slice();
    applyLeftColumnWidths();

    if (appData.tasks && appData.tasks.length > 0) {
        appData.tasks.forEach(tData => {
            tData.mainSchedule = normalizeMainSchedule(tData.mainSchedule);
            addTaskRow(tData);
        });
    } else {
        addTaskRow();
    }
}

function scrollToToday() {
    const todayIdx = timelineDays.findIndex(d => d.iso === todayISO);
    if (todayIdx !== -1) {
        const x = todayIdx * CELL_WIDTH;
        const scrollContainer = document.querySelector(".gantt-scroll-container");
        if (scrollContainer) {
            scrollContainer.scrollLeft = x - (scrollContainer.clientWidth / 2) + 280; 
        }
    }
}

function buildTimeline() {
    timelineDays = [];
    const startDt = isoToDate(appData.settings.startDate);
    const endDt = isoToDate(appData.settings.endDate);
    const curr = new Date(startDt);

    while (curr <= endDt) {
        const iso = dateToISO(curr);
        const dow = curr.getDay();
        timelineDays.push({
            index: timelineDays.length,
            date: new Date(curr),
            iso,
            day: curr.getDate(),
            dow,
            month: curr.getMonth() + 1,
            year: curr.getFullYear(),
            isWeekend: dow === 0 || dow === 6,
            isHoliday: appData.settings.holidays.includes(iso),
            isToday: iso === todayISO
        });
        curr.setDate(curr.getDate() + 1);
    }
    updateRangeLabel();
}

function updateRangeLabel() {
    if (!timelineDays.length) { rangeLabel.textContent = ""; return; }
    rangeLabel.textContent = `${dateToISO(timelineDays[0].date)} 〜 ${dateToISO(timelineDays[timelineDays.length - 1].date)}`;
}

function buildHeader() {
    const total = timelineDays.length;
    headerRow.innerHTML = "";
    headerRow.style.gridTemplateColumns = `repeat(${total}, ${CELL_WIDTH}px)`;
    timelineDays.forEach((d) => {
        const c = document.createElement("div");
        c.className = "header-day";
        if (d.isWeekend) c.classList.add("weekend");
        if (d.isHoliday) c.classList.add("holiday");
        if (d.isToday) c.classList.add("today");
        c.innerHTML = `<div class="header-day-num">${d.month}/${d.day}</div><div class="header-day-week">${WEEKDAYS[d.dow]}</div>`;
        headerRow.appendChild(c);
    });

    totalRow.innerHTML = "";
    totalRow.style.gridTemplateColumns = `repeat(${total}, ${CELL_WIDTH}px)`;
    timelineDays.forEach((d) => {
        const c = document.createElement("div");
        c.className = "total-cell";
        if (d.isWeekend) c.classList.add("weekend");
        if (d.isHoliday) c.classList.add("holiday");
        if (d.isToday) c.classList.add("today");
        c.dataset.iso = d.iso;
        totalRow.appendChild(c);
    });
}

// ============================================
// 行操作 (Drag & Drop)
// ============================================
let dragSrcEl = null;
function handleRowDragStart(e) {
    dragSrcEl = this.closest('.left-row');
    dragSrcEl.classList.add('dragging');
    e.dataTransfer.effectAllowed = 'move';
    const rows = Array.from(leftRowsContainer.children);
    e.dataTransfer.setData('text/plain', rows.indexOf(dragSrcEl));
}
function handleRowDrop(e) {
    e.stopPropagation();
    const targetRow = this.closest('.left-row');
    if (dragSrcEl !== targetRow) {
        const rows = Array.from(leftRowsContainer.children);
        const srcIdx = parseInt(e.dataTransfer.getData('text/plain'));
        const targetIdx = rows.indexOf(targetRow);
        const movedItem = taskObjects.splice(srcIdx, 1)[0];
        taskObjects.splice(targetIdx, 0, movedItem);
        refreshRowsDOM();
        triggerSave();
    }
    return false;
}
function handleRowDragEnd(e) {
    leftRowsContainer.querySelectorAll('.left-row').forEach(r => { r.classList.remove('over'); r.classList.remove('dragging'); });
}
function refreshRowsDOM() {
    taskObjects.forEach(task => { leftRowsContainer.appendChild(task.leftRowEl); rowsContainer.appendChild(task.rowEl); });
}

function addTaskRow(initialData = null) {
    const id = initialData ? initialData.id : "task_" + Date.now() + "_" + Math.random().toString(36).slice(2);
    const total = timelineDays.length;

    const leftRow = document.createElement("div");
    leftRow.className = "left-row";
    if (initialData && initialData.isDone) leftRow.classList.add("task-done");
    if (initialData && initialData.isHidden) leftRow.classList.add("task-hidden");

    const grip = document.createElement("div");
    grip.className = "drag-handle"; 
    grip.dataset.taskId = id;

    const gripIcon = document.createElement("span");
    gripIcon.className = "grip-icon";
    gripIcon.textContent = "⠿";
    grip.appendChild(gripIcon);

    const insertBtn = document.createElement("div");
    insertBtn.className = "row-insert-btn";
    insertBtn.textContent = "+";
    insertBtn.title = "この下に行を追加";
    const stopEvt = (e) => { e.stopPropagation(); };
    insertBtn.addEventListener("mousedown", stopEvt);
    insertBtn.addEventListener("dragstart", stopEvt);
    insertBtn.addEventListener("click", (e) => { e.stopPropagation(); insertTaskAfter(id); });
    
    grip.appendChild(insertBtn);

    grip.addEventListener('dragstart', (e) => {
        if (selectionMode !== 0) {
            e.preventDefault(); 
            return;
        }
        handleRowDragStart.call(grip, e);
    });
    
    grip.addEventListener("click", (e) => {
        if (selectionMode !== 0) {
            e.stopPropagation();
            const t = taskObjects.find(task => task.id === id);
            if (t) {
                t.isSelected = !t.isSelected;
                renderGrip(t);
            }
        }
    });

    leftRow.appendChild(grip);

    leftRow.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; return false; });
    leftRow.addEventListener('dragenter', function() { this.classList.add('over'); });
    leftRow.addEventListener('dragleave', function() { this.classList.remove('over'); });
    leftRow.addEventListener('drop', handleRowDrop);
    leftRow.addEventListener('dragend', handleRowDragEnd);

    const createInput = (ph, text) => {
        const cell = document.createElement("div"); cell.className = "label-cell";
        const ed = document.createElement("div"); ed.className = "editable"; ed.contentEditable = "true"; ed.dataset.placeholder = ph;
        if (text) ed.textContent = text;
        ed.addEventListener('blur', triggerSave);
        if (ph === "項目1") {
            const openMemo = (e) => {
                e.stopPropagation();
                openTaskMemo(task, task.leftRowEl, true);
            };
            cell.addEventListener("click", openMemo);
            ed.addEventListener("click", openMemo);
        }
        cell.appendChild(ed);
        return cell;
    };
    leftRow.appendChild(createInput("項目1", initialData ? initialData.label1 : ""));
    leftRow.appendChild(createInput("項目2", initialData ? initialData.label2 : ""));
    leftRow.appendChild(createInput("時間", initialData ? initialData.label3 : ""));
    leftRow.style.gridTemplateColumns = leftColumnWidths.map(w => `${w}px`).join(" ");

    const rowResize = document.createElement("div");
    rowResize.className = "row-resize-handle";
    rowResize.style.left = `${leftColumnWidths[0]}px`;
    rowResize.addEventListener("mousedown", (e) => {
        e.preventDefault();
        e.stopPropagation();
        isResizingRow = true;
        resizeRowTaskId = id;
        resizeStartY = e.clientY;
        resizeStartHeight = leftRow.getBoundingClientRect().height;
        document.body.style.cursor = "row-resize";
        document.body.style.userSelect = "none";
    });
    leftRow.appendChild(rowResize);

    leftRow.addEventListener('contextmenu', (e) => { e.preventDefault(); showContextMenu(e, id); });
    leftRowsContainer.appendChild(leftRow);

    const row = document.createElement("div");
    row.className = "task-row";
    row.dataset.id = id;
    row.style.gridTemplateColumns = `repeat(${total}, ${CELL_WIDTH}px)`;
    if (initialData && initialData.isDone) row.classList.add("task-done");
    if (initialData && initialData.isHidden) row.classList.add("task-hidden");

    const cellRow = document.createElement("div"); cellRow.style.display = "contents";
    for (let i = 0; i < total; i++) {
        const d = timelineDays[i];
        const c = document.createElement("div"); c.className = "cell";
        if (d.isWeekend) c.classList.add("weekend");
        if (d.isHoliday) c.classList.add("holiday");
        if (d.isToday) c.classList.add("today");
        c.dataset.index = i; cellRow.appendChild(c);
    }
    row.appendChild(cellRow);
    const segLayer = document.createElement("div"); segLayer.className = "segments-layer"; row.appendChild(segLayer);
    rowsContainer.appendChild(row);

    const task = {
        id, rowEl: row, leftRowEl: leftRow, cellRowEl: cellRow, segLayerEl: segLayer,
        mainSchedule: initialData ? normalizeMainSchedule(initialData.mainSchedule) : null,
        segments: initialData ? initialData.segments.map(seg => ({
            ...seg,
            isSelected: false,
            _lane: 0
        })) : [],
        memo: initialData ? (initialData.memo || "") : "",
        customHeight: initialData ? (initialData.customHeight || 0) : 0,
        isDone: initialData ? !!initialData.isDone : false,
        isHidden: initialData ? !!initialData.isHidden : false,
        pendingMainStartIndex: null, pendingMainStartDate: null,
        pendingStartIndex: null, pendingStartDate: null, pendingStartLane: 0,
        isSelected: false
    };
    taskObjects.push(task);
    
    renderGrip(task); 

    setupRowInteraction(task);
    activeTaskId = id;
    renderAllSegments();
    if (!initialData) triggerSave();
}

function renderGrip(task) {
    const grip = task.leftRowEl.querySelector(".drag-handle");
    if (!grip) return;
    const iconSpan = grip.querySelector(".grip-icon");
    if (!iconSpan) return;

    let iconText = "⠿";
    if (selectionMode === 1 || selectionMode === 2) {
        iconText = task.isSelected ? "☑️" : "□";
    }

    iconSpan.textContent = iconText;
    grip.draggable = (selectionMode === 0);
    grip.style.cursor = (selectionMode === 0) ? "grab" : "pointer";
}

function insertTaskAfter(targetTaskId) {
    addTaskRow();
    const newTask = taskObjects[taskObjects.length - 1];
    const targetIndex = taskObjects.findIndex(t => t.id === targetTaskId);
    if (targetIndex === -1) return;

    const targetTask = taskObjects[targetIndex];
    if (targetTask.leftRowEl.nextSibling) {
        leftRowsContainer.insertBefore(newTask.leftRowEl, targetTask.leftRowEl.nextSibling);
        rowsContainer.insertBefore(newTask.rowEl, targetTask.rowEl.nextSibling);
    }
    taskObjects.pop(); 
    taskObjects.splice(targetIndex + 1, 0, newTask);
    triggerSave();
}

// ============================================
// 描画ロジック
// ============================================
function renderAllSegments() {
    if (!timelineDays.length) return;
    const rangeStart = timelineDays[0].date;
    const rangeEnd = timelineDays[timelineDays.length - 1].date;

    taskObjects.forEach((task) => {
        task.segLayerEl.innerHTML = "";
        drawScheduleDivider(task);

        const taskDates = {};
        task.segments.forEach(seg => seg._lane = 0);
        const sortedSegs = [...task.segments].sort((a, b) => (a.startDate !== b.startDate) ? (a.startDate < b.startDate ? -1 : 1) : (a.endDate < b.endDate ? -1 : 1));
        
        let maxLaneUsed = 0;
        sortedSegs.forEach(seg => {
            let requiredLane = 0;
            let sIdx = dateToIndex(seg.startDate);
            let eIdx = dateToIndex(seg.endDate);
            if (sIdx === -1 || eIdx === -1) return;
            while (true) {
                let overlap = false;
                for (let i = Math.min(sIdx, eIdx); i <= Math.max(sIdx, eIdx); i++) {
                    const iso = timelineDays[i].iso;
                    if (taskDates[iso] && taskDates[iso].includes(requiredLane)) { overlap = true; break; }
                }
                if (!overlap) break;
                requiredLane++;
            }
            seg._lane = requiredLane;
            maxLaneUsed = Math.max(maxLaneUsed, requiredLane);
            for (let i = Math.min(sIdx, eIdx); i <= Math.max(sIdx, eIdx); i++) {
                const iso = timelineDays[i].iso;
                if (!taskDates[iso]) taskDates[iso] = [];
                taskDates[iso].push(requiredLane);
            }
        });

        const draftLaneCount = task.pendingStartIndex != null ? ((task.pendingStartLane || 0) + 1) : 0;
        const laneCount = Math.max(task.segments.length > 0 ? (maxLaneUsed + 1) : 1, draftLaneCount || 1);
        const newHeight = Math.max(BASE_ROW_HEIGHT, SUB_SCHEDULE_TOP + ((laneCount - 1) * SEGMENT_OFFSET_Y) + 26);
        task.baseHeight = newHeight;
        const finalHeight = Math.max(newHeight, task.customHeight || 0);
        task.rowEl.style.height = finalHeight + "px";
        task.leftRowEl.style.height = finalHeight + "px";

        if (task.mainSchedule) {
            drawMainSchedule(task);
        }

        task.segments.forEach((seg) => {
            const lane = seg._lane || 0;
            const topPx = SUB_SCHEDULE_TOP + (lane * SEGMENT_OFFSET_Y);
            if (seg.type === "point") {
                const idx = dateToIndex(seg.startDate);
                if (idx !== -1) drawPointSegment(task, seg, idx, topPx);
            } else {
                const sdt = isoToDate(seg.startDate), edt = isoToDate(seg.endDate);
                if (edt >= rangeStart && sdt <= rangeEnd) {
                    const vs = sdt < rangeStart ? rangeStart : sdt;
                    const ve = edt > rangeEnd ? rangeEnd : edt;
                    const sIdx = dateToIndex(dateToISO(vs));
                    const eIdx = dateToIndex(dateToISO(ve));
                    if (sIdx !== -1 && eIdx !== -1) drawRangeSegment(task, seg, sIdx, eIdx, topPx);
                }
            }
        });
        if (task.pendingMainStartIndex != null) drawDraftStart(task, task.pendingMainStartIndex, MAIN_LINE_Y);
        if (task.pendingStartIndex != null) drawDraftStart(task, null, SUB_SCHEDULE_TOP + ((task.pendingStartLane || 0) * SEGMENT_OFFSET_Y));
        adjustLabelPositions(task);
    });
    calculateTotals();
    updateBottomRowBorders();
}

function drawScheduleDivider(task) {
    const divider = document.createElement("div");
    divider.className = "schedule-divider";
    divider.style.top = MAIN_DIVIDER_Y + "px";
    task.segLayerEl.appendChild(divider);
}

function drawMainSchedule(task) {
    const main = task.mainSchedule;
    if (!main) return;

    const rangeStart = timelineDays[0]?.date;
    const rangeEnd = timelineDays[timelineDays.length - 1]?.date;
    const startDate = isoToDate(main.startDate);
    const endDate = isoToDate(main.endDate);
    if (!rangeStart || !rangeEnd || endDate < rangeStart || startDate > rangeEnd) return;

    const visibleStart = startDate < rangeStart ? rangeStart : startDate;
    const visibleEnd = endDate > rangeEnd ? rangeEnd : endDate;
    const sIdx = dateToIndex(dateToISO(visibleStart));
    const eIdx = dateToIndex(dateToISO(visibleEnd));
    if (sIdx === -1 || eIdx === -1) return;

    const sc = centerX(sIdx);
    const ec = centerX(eIdx);
    const left = Math.min(sc, ec);
    const width = Math.max(1, Math.abs(sc - ec));

    const line = document.createElement("div");
    line.className = "segment main-segment";
    line.style.left = left + "px";
    line.style.width = width + "px";
    line.style.top = MAIN_LINE_Y + "px";

    const lHandle = document.createElement("div");
    lHandle.className = "resize-handle left";
    lHandle.addEventListener("mousedown", (e) => initDrag(e, task, main, "resize-left", line, "main"));
    line.appendChild(lHandle);

    const rHandle = document.createElement("div");
    rHandle.className = "resize-handle right";
    rHandle.addEventListener("mousedown", (e) => initDrag(e, task, main, "resize-right", line, "main"));
    line.appendChild(rHandle);

    line.addEventListener("mousedown", (e) => {
        if (e.target.classList.contains("resize-handle")) return;
        initDrag(e, task, main, "move", line, "main");
    });
    line.addEventListener("dblclick", (e) => {
        e.preventDefault();
        e.stopPropagation();
    });
    line.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        e.stopPropagation();
        showSegmentContextMenu(e, task, { id: "__main__", scheduleScope: "main" });
    });
    task.segLayerEl.appendChild(line);

    if (main.progressEndDate) {
        const progressDate = isoToDate(main.progressEndDate);
        const progressVisibleEnd = progressDate > visibleEnd ? visibleEnd : progressDate;
        if (progressVisibleEnd >= visibleStart) {
            const pIdx = dateToIndex(dateToISO(progressVisibleEnd));
            if (pIdx !== -1) {
                const progressLeft = centerX(sIdx);
                const progressRight = (progressVisibleEnd < endDate && pIdx < eIdx)
                    ? (pIdx + 1) * CELL_WIDTH
                    : centerX(pIdx);
                const doneWidth = progressRight - progressLeft;
                if (doneWidth > 0) {
                    const doneLine = document.createElement("div");
                    doneLine.className = "segment main-segment done";
                    doneLine.style.left = progressLeft + "px";
                    doneLine.style.width = doneWidth + "px";
                    doneLine.style.top = MAIN_LINE_Y + "px";
                    doneLine.style.pointerEvents = "none";
                    task.segLayerEl.appendChild(doneLine);
                }
            }
        }
    }

    if (startDate >= rangeStart && startDate <= rangeEnd) {
        const pt = document.createElement("div");
        const startDone = main.progressEndDate && isoToDate(main.progressEndDate).getTime() >= startDate.getTime();
        pt.className = "point main-point" + (startDone ? " done" : "");
        pt.style.left = centerX(dateToIndex(main.startDate)) + "px";
        pt.style.top = MAIN_LINE_Y + "px";
        pt.style.cursor = "grab";
        pt.addEventListener("mousedown", (e) => initDrag(e, task, main, "move", line, "main"));
        pt.addEventListener("dblclick", (e) => {
            e.preventDefault();
            e.stopPropagation();
            editMainEndpointLabel(task, "start");
        });
        pt.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            e.stopPropagation();
            showSegmentContextMenu(e, task, { id: "__main__", scheduleScope: "main" });
        });
        task.segLayerEl.appendChild(pt);
        drawMainEndpointLabel(task, "start", centerX(dateToIndex(main.startDate)), main.startLabel || "");
    }

    if (endDate >= rangeStart && endDate <= rangeEnd) {
        const pt = document.createElement("div");
        const endDone = main.progressEndDate && isoToDate(main.progressEndDate).getTime() >= endDate.getTime();
        pt.className = "point main-point" + (endDone ? " done" : "");
        pt.style.left = centerX(dateToIndex(main.endDate)) + "px";
        pt.style.top = MAIN_LINE_Y + "px";
        pt.style.cursor = "grab";
        pt.addEventListener("mousedown", (e) => initDrag(e, task, main, "move", line, "main"));
        pt.addEventListener("dblclick", (e) => {
            e.preventDefault();
            e.stopPropagation();
            editMainEndpointLabel(task, "end");
        });
        pt.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            e.stopPropagation();
            showSegmentContextMenu(e, task, { id: "__main__", scheduleScope: "main" });
        });
        task.segLayerEl.appendChild(pt);
        drawMainEndpointLabel(task, "end", centerX(dateToIndex(main.endDate)), main.endLabel || "");
    }

    const milestones = [...(main.milestones || [])].sort((a, b) => a.date.localeCompare(b.date));
    milestones.forEach((milestone, index) => drawMainMilestone(task, milestone, index));
}

function drawMainMilestone(task, milestone, index) {
    const idx = dateToIndex(milestone.date);
    if (idx === -1) return;
    const x = centerX(idx);
    const isAbove = index % 2 === 0;
    const labelTop = isAbove ? (MAIN_LINE_Y - 23) : (MAIN_LINE_Y + 9);

    const pt = document.createElement("div");
    const isDone = task.mainSchedule?.progressEndDate && isoToDate(task.mainSchedule.progressEndDate).getTime() >= isoToDate(milestone.date).getTime();
    pt.className = "point milestone-point" + (isDone ? " done" : "");
    pt.style.left = x + "px";
    pt.style.top = MAIN_LINE_Y + "px";
    pt.style.cursor = "grab";
    pt.addEventListener("mousedown", (e) => initDrag(e, task, milestone, "move", pt, "milestone"));
    pt.addEventListener("dblclick", (e) => {
        e.preventDefault();
        e.stopPropagation();
        editMilestone(task, milestone);
    });
    pt.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (confirm("このマイルストーンを削除しますか？")) {
            task.mainSchedule.milestones = (task.mainSchedule.milestones || []).filter(ms => ms.id !== milestone.id);
            renderAllSegments();
            triggerSave();
        }
    });
    task.segLayerEl.appendChild(pt);

    const label = document.createElement("div");
    label.className = "segment-label milestone-label";
    label.textContent = milestone.label || "マイルストーン";
    label.style.left = x + "px";
    label.dataset.baseTop = String(labelTop);
    label.style.top = labelTop + "px";
    label.addEventListener("dblclick", (e) => {
        e.preventDefault();
        e.stopPropagation();
        editMilestone(task, milestone);
    });
    label.addEventListener("click", (e) => {
        if (isCtrlSelectionMode || e.ctrlKey) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        e.preventDefault();
        e.stopPropagation();
        editMilestone(task, milestone);
    });
    label.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        e.stopPropagation();
        if (confirm("このマイルストーンを削除しますか？")) {
            task.mainSchedule.milestones = (task.mainSchedule.milestones || []).filter(ms => ms.id !== milestone.id);
            renderAllSegments();
            triggerSave();
        }
    });
    task.segLayerEl.appendChild(label);
}

function editMilestone(task, milestone) {
    const nextLabel = prompt("マイルストーン名:", milestone.label || "マイルストーン");
    if (nextLabel === null) return;
    milestone.label = nextLabel.trim() || "マイルストーン";
    renderAllSegments();
    triggerSave();
}

function editSubSegmentLabel(task, seg) {
    const nextLabel = window.prompt("計画内容:", seg.label || "");
    if (nextLabel === null) return;
    seg.label = nextLabel.trim();
    renderAllSegments();
    triggerSave();
}

function handleSubLabelMouseDown(e, task, seg, dragEl) {
    if (!(isCtrlSelectionMode && seg.isSelected)) return;
    initDrag(e, task, seg, "move", dragEl);
}

function drawMainEndpointLabel(task, side, x, text) {
    if (!text) return;
    const label = document.createElement("div");
    label.className = "segment-label main-endpoint-label";
    label.textContent = text;
    label.style.left = x + "px";
    label.dataset.baseTop = String(MAIN_LINE_Y - 23);
    label.style.top = (MAIN_LINE_Y - 23) + "px";
    label.addEventListener("dblclick", (e) => {
        e.preventDefault();
        e.stopPropagation();
        editMainEndpointLabel(task, side);
    });
    label.addEventListener("click", (e) => {
        if (isCtrlSelectionMode || e.ctrlKey) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        e.preventDefault();
        e.stopPropagation();
        editMainEndpointLabel(task, side);
    });
    task.segLayerEl.appendChild(label);
}

function editMainEndpointLabel(task, side) {
    const main = task.mainSchedule;
    if (!main) return;
    const key = side === "start" ? "startLabel" : "endLabel";
    const promptLabel = side === "start" ? "開始コメント" : "終了コメント";
    const nextLabel = prompt(`${promptLabel}:`, main[key] || "");
    if (nextLabel === null) return;
    main[key] = nextLabel.trim();
    renderAllSegments();
    triggerSave();
}

function calculateTotals() {
    const totals = {};
    timelineDays.forEach(d => totals[d.iso] = 0);
    taskObjects.forEach(task => {
        if (task.isHidden) return;
        task.segments.forEach(seg => {
            if (seg.dailyValues) {
                for (const [iso, val] of Object.entries(seg.dailyValues)) {
                    if (iso < seg.startDate || iso > seg.endDate) continue;
                    const num = parseFloat(val);
                    if (!isNaN(num)) totals[iso] += num;
                }
            }
        });
    });
    const cells = totalRow.children;
    for (let i = 0; i < cells.length; i++) {
        const cell = cells[i];
        const iso = cell.dataset.iso;
        let val = totals[iso];
        if (val > 0) {
            if (val > 99.9) val = 99.9;
            cell.textContent = (val % 1 === 0) ? val : val.toFixed(1);
        } else {
            cell.textContent = "";
        }
    }
}

function drawRangeSegment(task, seg, sIdx, eIdx, topPx) {
    const sc = centerX(sIdx), ec = centerX(eIdx);
    const baseLeft = Math.min(sc, ec), baseWidth = Math.max(1, Math.abs(sc - ec));
    const isProgressSelected = activeProgressSegmentId === seg.id;
    const isSelected = !!seg.isSelected;

    const hitbox = document.createElement("div");
    hitbox.className = "segment-hitbox" + (isSelected ? " segment-selected" : "");
    hitbox.dataset.subSelectable = "true";
    hitbox.style.left = baseLeft + "px";
    hitbox.style.width = baseWidth + "px";
    hitbox.style.top = topPx + "px";
    hitbox.addEventListener("mousedown", (e) => {
        if (e.ctrlKey) return;
        if (isCtrlSelectionMode && !seg.isSelected) return;
        initDrag(e, task, seg, "move", div);
    });
    addSegEvents(hitbox, task, seg);
    task.segLayerEl.appendChild(hitbox);

    const div = document.createElement("div");
    div.className = "segment" + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
    div.dataset.subSelectable = "true";
    div.style.left = baseLeft + "px";
    div.style.width = baseWidth + "px";
    div.style.top = topPx + "px";

    if (seg.progressEndDate) {
        div.classList.add("fixed");
    }

    // [修正] 実績(progressEndDate)がある場合は開始日が固定されるため、左ハンドルは生成しない
    if (!seg.progressEndDate) {
        const lHandle = document.createElement("div"); lHandle.className = "resize-handle left";
        lHandle.addEventListener("mousedown", (e) => initDrag(e, task, seg, "resize-left", div));
        div.appendChild(lHandle);
    }

    // [修正] 完全に完了している(progress >= end)場合のみ右ハンドルを生成しない（未完了なら生成する）
    const isFullyDone = seg.progressEndDate && (isoToDate(seg.progressEndDate).getTime() >= isoToDate(seg.endDate).getTime());
    if (!isFullyDone) {
        const rHandle = document.createElement("div"); rHandle.className = "resize-handle right";
        rHandle.addEventListener("mousedown", (e) => initDrag(e, task, seg, "resize-right", div));
        div.appendChild(rHandle);
    }

    div.addEventListener("mousedown", (e) => {
        if (e.target.classList.contains("resize-handle")) return;
        initDrag(e, task, seg, "move", div);
    });

    addSegEvents(div, task, seg);
    task.segLayerEl.appendChild(div);

    const minI = Math.min(dateToIndex(seg.startDate), dateToIndex(seg.endDate));
    const maxI = Math.max(dateToIndex(seg.startDate), dateToIndex(seg.endDate));
    if (minI !== -1 && maxI !== -1) {
        for (let i = minI; i <= maxI; i++) {
            const iso = timelineDays[i].iso;
            const x = centerX(i);
            const valDiv = document.createElement("div");
            valDiv.className = "daily-val";
            if (activeDailyValueTarget && activeDailyValueTarget.segId === seg.id && activeDailyValueTarget.iso === iso) {
                valDiv.classList.add("active");
            }
            valDiv.style.left = x + "px";
            valDiv.style.top = (topPx + 4) + "px";
            if (seg.dailyValues && seg.dailyValues[iso] != null) {
                const rawV = seg.dailyValues[iso];
                const v = parseFloat(rawV);
                valDiv.textContent = (!isNaN(v) && /^\d(\.\d)?$/.test(rawV)) ? ((v % 1 === 0) ? v : v.toFixed(1)) : rawV;
            }
            valDiv.addEventListener("click", (e) => handleDailyValueClick(e, task, seg, iso));
            task.segLayerEl.appendChild(valDiv);
        }
    }

    if (seg.progressEndDate) {
        const sIdxRaw = dateToIndex(seg.startDate);
        const pIdxRaw = dateToIndex(seg.progressEndDate);
        if (sIdxRaw !== -1 && pIdxRaw !== -1 && pIdxRaw >= sIdxRaw) {
            const left = centerX(sIdxRaw);
            const eIdxRaw = dateToIndex(seg.endDate);
            let right = (eIdxRaw !== -1 && pIdxRaw < eIdxRaw) ? (pIdxRaw + 1) * CELL_WIDTH : centerX(pIdxRaw);
            const w = right - left;
            if (w > 0) {
                const dDiv = document.createElement("div");
                dDiv.className = "segment done";
                dDiv.style.left = left + "px";
                dDiv.style.width = w + "px";
                dDiv.style.pointerEvents = "none";
                dDiv.style.top = topPx + "px";
                task.segLayerEl.appendChild(dDiv);
            }
        }
    }

    const pointsData = [ { x: sc, d: isoToDate(seg.startDate), isEnd: false }, { x: ec, d: isoToDate(seg.endDate), isEnd: true } ];
    pointsData.forEach((ptData) => {
        const pt = document.createElement("div");
        let isDone = seg.progressEndDate && (isoToDate(seg.progressEndDate).getTime() >= ptData.d.getTime());
        pt.className = "point" + (isDone ? " done" : "") + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
        pt.dataset.subSelectable = "true";
        pt.style.left = ptData.x + "px"; pt.style.top = topPx + "px";
        
        pt.style.cursor = "grab";
        pt.addEventListener("mousedown", (e) => initDrag(e, task, seg, "move", div));

        addSegEvents(pt, task, seg);
        task.segLayerEl.appendChild(pt);
    });

    if (seg.label) {
        const lab = document.createElement("div");
        const isCompletedFull = seg.progressEndDate && isoToDate(seg.progressEndDate).getTime() >= isoToDate(seg.endDate).getTime();
        lab.className = "segment-label" + (isCompletedFull ? " done" : "") + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
        lab.dataset.subSelectable = "true";
        lab.textContent = seg.label;
        lab.style.left = (sc + ec) / 2 + "px";
        lab.addEventListener("mousedown", (e) => handleSubLabelMouseDown(e, task, seg, div));
        
        const baseTop = topPx - 19;
        lab.dataset.baseTop = baseTop; 
        lab.style.top = baseTop + "px";
        
        addSegEvents(lab, task, seg);
        task.segLayerEl.appendChild(lab);
    }
}

function drawPointSegment(task, seg, idx, topPx) {
    const c = centerX(idx);
    const isProgressSelected = activeProgressSegmentId === seg.id;
    const isDone = seg.progressEndDate && isoToDate(seg.progressEndDate).getTime() >= isoToDate(seg.startDate).getTime();
    const isSelected = !!seg.isSelected;

    const pt = document.createElement("div");
    pt.className = "point" + (isDone ? " done" : "") + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
    pt.dataset.subSelectable = "true";
    pt.style.left = c + "px"; pt.style.top = topPx + "px";
    
    pt.style.cursor = "grab";
    pt.addEventListener("mousedown", (e) => initDrag(e, task, seg, "move", pt));

    addSegEvents(pt, task, seg);
    task.segLayerEl.appendChild(pt);

    const iso = timelineDays[idx].iso;
    const valDiv = document.createElement("div");
    valDiv.className = "daily-val";
    if (activeDailyValueTarget && activeDailyValueTarget.segId === seg.id && activeDailyValueTarget.iso === iso) {
        valDiv.classList.add("active");
    }
    valDiv.style.left = c + "px"; 
    valDiv.style.top = (topPx + 4) + "px";
    if (seg.dailyValues && seg.dailyValues[iso] != null) {
        valDiv.textContent = seg.dailyValues[iso];
    }
    valDiv.addEventListener("click", (e) => handleDailyValueClick(e, task, seg, iso));
    task.segLayerEl.appendChild(valDiv);

    if (seg.label) {
        const lab = document.createElement("div");
        lab.className = "segment-label" + (isDone ? " done" : "") + (isProgressSelected ? " progress-active" : "") + (isSelected ? " segment-selected" : "");
        lab.dataset.subSelectable = "true";
        lab.textContent = seg.label;
        lab.style.left = c + "px"; 
        lab.addEventListener("mousedown", (e) => handleSubLabelMouseDown(e, task, seg, pt));
        
        const baseTop = topPx - 19;
        lab.dataset.baseTop = baseTop;
        lab.style.top = baseTop + "px";
        
        addSegEvents(lab, task, seg);
        task.segLayerEl.appendChild(lab);
    }
}

function drawDraftStart(task, index = null, topPx = 30) {
    const targetIndex = index == null ? task.pendingStartIndex : index;
    if (targetIndex == null) return;
    const c = centerX(targetIndex);
    const pt = document.createElement("div"); pt.className = "point draft";
    pt.style.left = c + "px"; pt.style.top = topPx + "px";
    pt.title = "キャンセル";
    pt.addEventListener("click", (e) => {
        e.stopPropagation();
        if (index == null) {
            task.pendingStartIndex = null;
            task.pendingStartDate = null;
            task.pendingStartLane = 0;
        } else {
            task.pendingMainStartIndex = null;
            task.pendingMainStartDate = null;
        }
        renderAllSegments();
    });
    task.segLayerEl.appendChild(pt);
}

function adjustLabelPositions(task) {
    const labels = Array.from(task.segLayerEl.querySelectorAll(".segment-label:not(.milestone-label):not(.main-endpoint-label)"));
    if (labels.length === 0) return;

    const groups = {};
    labels.forEach(el => {
        const baseTop = parseFloat(el.dataset.baseTop);
        const key = Math.round(baseTop);
        if (!groups[key]) groups[key] = [];
        groups[key].push({ el, baseTop, left: parseFloat(el.style.left) });
    });

    Object.values(groups).forEach(items => {
        items.sort((a, b) => a.left - b.left);
        items.forEach((item, index) => {
            if (index % 2 === 0) {
                item.el.style.top = item.baseTop + "px";
                item.el.style.zIndex = "20";
            } else {
                item.el.style.top = (item.baseTop - 13.5) + "px";
                item.el.style.zIndex = "30";
            }
        });
    });
}

// ============================================
// Drag & Drop ロジック (移動・伸縮)
// ============================================
function initDrag(e, task, seg, type, el, scope = "sub") {
    if (e.button !== 0) return;
    if (isOverlayCancelStateActive()) return;
    if (e.ctrlKey) return;
    if (scope === "sub" && isCtrlSelectionMode && !seg.isSelected) return;

    let dragType = type;
    // [修正] 完了済み(progressEndDateあり)の場合、「blocked」という状態でドラッグを開始する。
    // 即座にreturnせず、グローバルなマウスイベントをあえて設定することで、
    // ドラッグ中のマウス操作をこの機能が「乗っ取る」形にし、裏側のセルに反応させないようにする。
    if (scope === "sub" && type === "move" && seg.progressEndDate) {
        dragType = "blocked";
    }

    // デフォルト動作(テキスト選択など)と伝播を阻止
    e.preventDefault();
    e.stopPropagation();

    let selectedSegRefs = [];
    if (scope === "sub" && type === "move") {
        const currentSelected = getSelectedSubSegments();
        if (seg.isSelected && currentSelected.length > 0) {
            selectedSegRefs = currentSelected.map(ref => ({
                taskId: ref.task.id,
                segId: ref.seg.id,
                originalStartDate: ref.seg.startDate,
                originalEndDate: ref.seg.endDate
            }));
        }
    }

    dragState = {
        isDragging: true, type: dragType, taskId: task.id, segId: seg.id || null, milestoneId: scope === "milestone" ? (seg.id || null) : null, selectedSegRefs, scope, startX: e.clientX,
        originalLeft: parseFloat(el.style.left), originalWidth: parseFloat(el.style.width),
        originalStartDate: seg.startDate || seg.date, originalEndDate: seg.endDate || seg.date, el: el
    };
    el.classList.add("dragging");
    
    // blockedの場合は「掴んでいる」ことを示すカーソルにするが、位置は更新されない
    if (dragType === "blocked") {
        document.body.style.cursor = "grabbing"; // 掴んでいるが動かせない
    } else {
        document.body.style.cursor = type === "move" ? "grabbing" : "col-resize";
    }
    
    document.addEventListener("mousemove", handleGlobalMouseMove);
    document.addEventListener("mouseup", handleGlobalMouseUp);
}

function handleGlobalMouseMove(e) {
    if (!dragState.isDragging) return;
    e.preventDefault(); // これにより裏側のテキスト選択やセルの反応を防ぐ

    // [修正] blocked状態（実績ありの移動）なら、座標計算もスタイル更新もしない
    if (dragState.type === "blocked") {
        return;
    }

    const diffPx = e.clientX - dragState.startX;
    if (dragState.type === "move") {
        dragState.el.style.left = (dragState.originalLeft + diffPx) + "px";
    } else if (dragState.type === "resize-right") {
        const newW = Math.max(0, dragState.originalWidth + diffPx); 
        dragState.el.style.width = newW + "px";
    } else if (dragState.type === "resize-left") {
        const newLeft = dragState.originalLeft + diffPx;
        const newWidth = dragState.originalWidth - diffPx;
        if (newWidth >= 0) { 
            dragState.el.style.left = newLeft + "px";
            dragState.el.style.width = newWidth + "px";
        }
    }
}

function handleGlobalMouseUp(e) {
    if (!dragState.isDragging) return;

    // [修正] blocked状態なら、何も計算せずにクリーンアップへ進む
    if (dragState.type !== "blocked") {
        const diffPx = e.clientX - dragState.startX;
        const dayDelta = Math.round(diffPx / CELL_WIDTH);
        const task = taskObjects.find(t => t.id === dragState.taskId);
        const seg = !task ? null : (
            dragState.scope === "main"
                ? task.mainSchedule
                : dragState.scope === "milestone"
                    ? (task.mainSchedule?.milestones || []).find(ms => ms.id === dragState.milestoneId)
                    : task.segments.find(s => s.id === dragState.segId)
        );

        if (task && seg && dayDelta !== 0) {
            if (dragState.scope === "milestone" && dragState.type === "move") {
                const shiftedDate = shiftDateStr(dragState.originalStartDate, dayDelta);
                const main = task.mainSchedule;
                if (main) {
                    const clampedDate = shiftedDate < main.startDate
                        ? main.startDate
                        : shiftedDate > main.endDate
                            ? main.endDate
                            : shiftedDate;
                    seg.date = clampedDate;
                }
            } else if (dragState.scope === "sub" && dragState.type === "move" && dragState.selectedSegRefs.length > 0) {
                dragState.selectedSegRefs.forEach(ref => {
                    const refTask = taskObjects.find(t => t.id === ref.taskId);
                    const refSeg = refTask ? refTask.segments.find(s => s.id === ref.segId) : null;
                    if (!refSeg) return;
                    refSeg.startDate = shiftDateStr(ref.originalStartDate, dayDelta);
                    refSeg.endDate = shiftDateStr(ref.originalEndDate, dayDelta);
                    if (refSeg.dailyValues) {
                        const newVals = {};
                        Object.keys(refSeg.dailyValues).forEach(iso => newVals[shiftDateStr(iso, dayDelta)] = refSeg.dailyValues[iso]);
                        refSeg.dailyValues = newVals;
                    }
                    if (refSeg.dailyResults) {
                        const newRes = {};
                        Object.keys(refSeg.dailyResults).forEach(iso => newRes[shiftDateStr(iso, dayDelta)] = refSeg.dailyResults[iso]);
                        refSeg.dailyResults = newRes;
                    }
                });
            } else if (dragState.type === "move") {
                seg.startDate = shiftDateStr(dragState.originalStartDate, dayDelta);
                seg.endDate = shiftDateStr(dragState.originalEndDate, dayDelta);

                if (dragState.scope === "main" && Array.isArray(seg.milestones)) {
                    seg.milestones = seg.milestones.map(ms => ({
                        ...ms,
                        date: shiftDateStr(ms.date, dayDelta)
                    }));
                }

                if (dragState.scope !== "main" && seg.dailyValues) {
                    const newVals = {};
                    Object.keys(seg.dailyValues).forEach(iso => newVals[shiftDateStr(iso, dayDelta)] = seg.dailyValues[iso]);
                    seg.dailyValues = newVals;
                }
                if (dragState.scope !== "main" && seg.dailyResults) {
                    const newRes = {};
                    Object.keys(seg.dailyResults).forEach(iso => newRes[shiftDateStr(iso, dayDelta)] = seg.dailyResults[iso]);
                    seg.dailyResults = newRes;
                }

            } else if (dragState.type === "resize-right") {
                const newEnd = shiftDateStr(dragState.originalEndDate, dayDelta);
                seg.endDate = newEnd; 
                if(seg.endDate < seg.startDate) seg.endDate = seg.startDate;
            } else if (dragState.type === "resize-left") {
                const newStart = shiftDateStr(dragState.originalStartDate, dayDelta);
                seg.startDate = newStart;
                if(seg.startDate > seg.endDate) seg.startDate = seg.endDate;
            }
            triggerSave();
        }
    }

    if (dragState.el) dragState.el.classList.remove("dragging");
    document.body.style.cursor = "";
    document.removeEventListener("mousemove", handleGlobalMouseMove);
    document.removeEventListener("mouseup", handleGlobalMouseUp);
    dragState.isDragging = false; dragState.el = null;
    renderAllSegments();
}

// ============================================
// インタラクション (クリック等)
// ============================================
function setupRowInteraction(task) {
    task.rowEl.addEventListener("click", (e) => {
        if (segmentContextMenu.style.display === "block" || contextMenu.style.display === "block") return;
        if (e.target.closest(".segment") || e.target.closest(".point") || e.target.closest(".segment-label") || e.target.closest(".daily-val")) return;
        if (isCtrlSelectionMode || e.ctrlKey) return;
        clearSegmentSelection();
        const rect = task.rowEl.getBoundingClientRect();
        const idx = Math.max(0, Math.min(timelineDays.length - 1, Math.floor((e.clientX - rect.left) / CELL_WIDTH)));
        const y = e.clientY - rect.top;
        if (y <= MAIN_DIVIDER_Y) handleMainCellClick(task, idx);
        else handleCellClick(task, idx, y);
    });
    task.leftRowEl.addEventListener("click", () => {
        if (segmentContextMenu.style.display === "block" || contextMenu.style.display === "block" || activeProgressSegmentId) return;
        if (isCtrlSelectionMode) return;
        activeTaskId = task.id;
        taskObjects.forEach(t => {
            t.pendingStartDate = null;
            t.pendingStartIndex = null;
            t.pendingStartLane = 0;
            t.pendingMainStartDate = null;
            t.pendingMainStartIndex = null;
        });
        renderAllSegments();
    });
}

function updateBottomRowBorders() {
    const rows = rowsContainer.querySelectorAll(".task-row");
    if (rows.length === 0) return;
    rows.forEach(r => r.classList.remove("is-last-visible"));
    let last = null;
    for (let i = rows.length - 1; i >= 0; i--) {
        if (!rows[i].classList.contains("task-hidden")) { last = rows[i]; break; }
    }
    if (last) last.classList.add("is-last-visible");

    const leftRows = leftRowsContainer.querySelectorAll(".left-row");
    leftRows.forEach(r => r.classList.remove("is-last-visible"));
    let lastLeft = null;
    for (let i = leftRows.length - 1; i >= 0; i--) {
        if (!leftRows[i].classList.contains("task-hidden")) { lastLeft = leftRows[i]; break; }
    }
    if (lastLeft) lastLeft.classList.add("is-last-visible");
}

function handleCellClick(task, index, y = SUB_SCHEDULE_TOP) {
    if (isCtrlSelectionMode) return;
    const clickedIso = timelineDays[index].iso;
    task.pendingMainStartIndex = null;
    task.pendingMainStartDate = null;

    if (activeProgressSegmentId) {
        const targetSeg = activeProgressSegmentId === "__main__"
            ? task.mainSchedule
            : task.segments.find(s => s.id === activeProgressSegmentId);
        if (targetSeg) {
            targetSeg.progressEndDate = clickedIso;
            activeProgressSegmentId = null; 
            activeProgressTaskId = null;
            renderAllSegments(); 
            triggerSave();
        } else {
            alert("選択中のバーはこの行にありません。");
        }
    } else {
        if (task.pendingStartIndex === null) {
            task.pendingStartIndex = index;
            task.pendingStartDate = clickedIso;
            task.pendingStartLane = getSubScheduleLaneFromY(y);
            renderAllSegments();
        } else {
            const startIso = task.pendingStartDate;
            const endIso = clickedIso;
            const s = startIso < endIso ? startIso : endIso;
            const e = startIso < endIso ? endIso : startIso;
            
            const newSeg = {
                id: "seg_" + Date.now() + "_" + Math.random().toString(36).slice(2),
                startDate: s, endDate: e, type: "range", 
                label: "...", 
                progressEndDate: null, dailyValues: {}, dailyResults: {}
            };
            task.segments.push(newSeg);

            task.pendingStartIndex = null; 
            task.pendingStartDate = null;
            task.pendingStartLane = 0;
            
            renderAllSegments();

            setTimeout(() => {
                const initialLabel = "新規作業";
                const inputLabel = prompt("計画内容を入力してください:", initialLabel);
                
                if (inputLabel === null) {
                    task.segments.pop(); 
                    renderAllSegments();
                } else {
                    newSeg.label = (inputLabel.trim() === "") ? initialLabel : inputLabel;
                    renderAllSegments();
                    triggerSave();
                }
            }, 10);
        }
    }
}

function handleSegClick(task, seg, addMode) {
    if (isOverlayCancelStateActive()) return;
    if (!seg.id) return;
    activeTaskId = task.id;
    if (addMode && !isCtrlSelectionMode) {
        updateCtrlSelectionMode(true);
        seg.isSelected = true;
        renderAllSegments();
        return;
    }
    if (isCtrlSelectionMode) {
        if (addMode) {
            seg.isSelected = !seg.isSelected;
        } else if (!seg.isSelected) {
            seg.isSelected = true;
        }
        renderAllSegments();
        return;
    }
}

function addSegEvents(el, task, seg) {
    el.addEventListener("click", (e) => {
        e.stopPropagation();
        if (isOverlayCancelStateActive()) return;
        if (el.classList.contains("segment-label") && !isCtrlSelectionMode && !e.ctrlKey) {
            editSubSegmentLabel(task, seg);
            return;
        }
        handleSegClick(task, seg, e.ctrlKey || isCtrlSelectionMode);
    });
    
    el.addEventListener("dblclick", (e) => {
        if (isOverlayCancelStateActive()) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        if (isCtrlSelectionMode || e.ctrlKey) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        e.preventDefault();
        e.stopPropagation();
        
        if (window.getSelection) {
            window.getSelection().removeAllRanges();
        }
        
        editSubSegmentLabel(task, seg);
    });
    
    el.addEventListener("contextmenu", (e) => {
        if (activeProgressSegmentId) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        if (isCtrlSelectionMode || e.ctrlKey) {
            e.preventDefault();
            e.stopPropagation();
            return;
        }
        e.preventDefault();
        e.stopPropagation(); 
        showSegmentContextMenu(e, task, seg);
    });
}

function handleDailyValueClick(e, task, seg, iso) {
    e.stopPropagation();
    if (isOverlayCancelStateActive()) return;
    if (isCtrlSelectionMode || e.ctrlKey) return;
    if (
        activeDailyValueTarget &&
        activeDailyValueTarget.taskId === task.id &&
        activeDailyValueTarget.segId === seg.id &&
        activeDailyValueTarget.iso === iso
    ) {
        clearActiveDailyValueTarget();
        renderAllSegments();
        return;
    }
    editDailyValue(task, seg, iso);
}

// ============================================
// コンテキストメニューなど
// ============================================
function showContextMenu(e, taskId) {
    contextMenuTargetTaskId = taskId;
    const task = taskObjects.find(t => t.id === taskId);
    const hideBtn = document.getElementById("cmHide");
    const unhideBtn = document.getElementById("cmUnhide");
    if (task.isHidden) { hideBtn.style.display = "none"; unhideBtn.style.display = "block"; }
    else { hideBtn.style.display = "block"; unhideBtn.style.display = "none"; }
    
    segmentContextMenu.style.display = "none";
    contextMenu.style.display = "block";
    contextMenu.style.left = e.pageX + "px";
    contextMenu.style.top = e.pageY + "px";
}

function showSegmentContextMenu(e, task, seg) {
    contextMenuTargetSegId = seg.id;
    contextMenuTargetTaskForSeg = task;
    segmentContextMenu.dataset.scope = seg.scheduleScope || "sub";
    
    contextMenu.style.display = "none";
    segmentContextMenu.style.display = "block";
    segmentContextMenu.style.left = e.pageX + "px";
    segmentContextMenu.style.top = e.pageY + "px";
}

document.addEventListener("click", () => { 
    if (!isCtrlSelectionMode) {
        clearSegmentSelection();
        renderAllSegments();
    }
    hideContextMenus();
});

document.getElementById("cmComplete").addEventListener("click", () => {
    const t = taskObjects.find(t => t.id === contextMenuTargetTaskId);
    if (t) { t.isDone = !t.isDone; t.leftRowEl.classList.toggle("task-done", t.isDone); t.rowEl.classList.toggle("task-done", t.isDone); triggerSave(); }
});
document.getElementById("cmHide").addEventListener("click", () => {
    const t = taskObjects.find(t => t.id === contextMenuTargetTaskId);
    if (t) { t.isHidden = true; t.leftRowEl.classList.add("task-hidden"); t.rowEl.classList.add("task-hidden"); triggerSave(); }
});
document.getElementById("cmUnhide").addEventListener("click", () => {
    const t = taskObjects.find(t => t.id === contextMenuTargetTaskId);
    if (t) { t.isHidden = false; t.leftRowEl.classList.remove("task-hidden"); t.rowEl.classList.remove("task-hidden"); triggerSave(); }
});

document.getElementById("ctxSegProgress").addEventListener("click", () => {
    if (contextMenuTargetSegId) {
        activeProgressSegmentId = contextMenuTargetSegId;
        activeProgressTaskId = contextMenuTargetTaskForSeg ? contextMenuTargetTaskForSeg.id : null;
        renderAllSegments();
    }
});

document.getElementById("ctxSegDelete").addEventListener("click", () => {
    if (contextMenuTargetTaskForSeg && contextMenuTargetSegId) {
        if (confirm("選択の計画を削除しますか？")) {
            if (segmentContextMenu.dataset.scope === "main") {
                contextMenuTargetTaskForSeg.mainSchedule = null;
                if (activeProgressSegmentId === "__main__") {
                    activeProgressSegmentId = null;
                    activeProgressTaskId = null;
                }
            } else {
                contextMenuTargetTaskForSeg.segments = contextMenuTargetTaskForSeg.segments.filter(s => s.id !== contextMenuTargetSegId);
                if (activeProgressSegmentId === contextMenuTargetSegId) {
                    activeProgressSegmentId = null;
                    activeProgressTaskId = null;
                }
            }
            renderAllSegments(); 
            triggerSave();
        }
    }
});

showHiddenCheck.addEventListener("change", (e) => {
    document.body.classList.toggle("show-hidden-mode", e.target.checked);
});

// ============================================
// イベント設定
// ============================================
function setupControlEvents() {
    document.addEventListener("mousedown", (e) => {
        const clickedInsideSegmentMenu = segmentContextMenu.style.display === "block" && segmentContextMenu.contains(e.target);
        const clickedInsideTaskMenu = contextMenu.style.display === "block" && contextMenu.contains(e.target);
        if (clickedInsideSegmentMenu || clickedInsideTaskMenu) return;

        if (segmentContextMenu.style.display === "block" || contextMenu.style.display === "block") {
            e.preventDefault();
            e.stopPropagation();
            dismissContextMenusWithClickSuppression();
            return;
        }

        if (!activeProgressSegmentId) return;
        if (isValidProgressClickTarget(e.target)) return;
        e.preventDefault();
        e.stopPropagation();
        cancelActiveProgressSelection();
    }, true);

    document.addEventListener("click", (e) => {
        if (suppressNextClickAfterMenuDismiss) {
            e.preventDefault();
            e.stopPropagation();
            suppressNextClickAfterMenuDismiss = false;
            return;
        }
        if (suppressNextClickAfterProgressCancel) {
            e.preventDefault();
            e.stopPropagation();
            suppressNextClickAfterProgressCancel = false;
            return;
        }
        const clickedInsideSegmentMenu = segmentContextMenu.style.display === "block" && segmentContextMenu.contains(e.target);
        const clickedInsideTaskMenu = contextMenu.style.display === "block" && contextMenu.contains(e.target);
        if (clickedInsideSegmentMenu || clickedInsideTaskMenu) return;

        if (segmentContextMenu.style.display === "block" || contextMenu.style.display === "block") {
            e.preventDefault();
            e.stopPropagation();
            return;
        }

        if (!activeProgressSegmentId) return;
        if (isValidProgressClickTarget(e.target)) return;
        e.preventDefault();
        e.stopPropagation();
    }, true);

    document.addEventListener("keydown", (e) => {
        const tag = document.activeElement?.tagName;
        const isEditable = document.activeElement?.isContentEditable;
        if (tag === "INPUT" || tag === "TEXTAREA" || isEditable) return;

        const ref = getActiveDailyValueRef();
        if (!ref) return;

        if (e.key === "ArrowRight") {
            e.preventDefault();
            moveActiveDailyValue(1);
        } else if (e.key === "ArrowLeft") {
            e.preventDefault();
            moveActiveDailyValue(-1);
        } else if (e.key === "Enter") {
            e.preventDefault();
            editDailyValue(ref.task, ref.seg, ref.iso);
        } else if (e.key === "Escape") {
            e.preventDefault();
            clearActiveDailyValueTarget();
            renderAllSegments();
        }
    });

    window.addEventListener("blur", () => updateCtrlSelectionMode(false));
    document.addEventListener("click", (e) => {
        if (!isCtrlSelectionMode) return;
        if (isSubSegmentSelectableTarget(e.target)) return;
        e.preventDefault();
        e.stopPropagation();
        clearSegmentSelection();
        renderAllSegments();
        updateCtrlSelectionMode(false);
    }, true);

    document.getElementById("undoBtn").addEventListener("click", () => HistoryManager.undo());
    document.getElementById("redoBtn").addEventListener("click", () => HistoryManager.redo());

    document.getElementById("settingsButton").addEventListener("click", () => {
        document.getElementById("settingsStartDate").value = appData.settings.startDate;
        document.getElementById("settingsEndDate").value = appData.settings.endDate;
        document.getElementById("settingsHolidays").value = appData.settings.holidays.join(", ");
        settingsPanel.classList.remove("settings-hidden");
    });
    document.getElementById("settingsCancel").addEventListener("click", () => settingsPanel.classList.add("settings-hidden"));
    document.getElementById("settingsSave").addEventListener("click", () => {
        if (!confirm("期間を変更しますか？")) return;
        appData.settings.startDate = document.getElementById("settingsStartDate").value;
        appData.settings.endDate = document.getElementById("settingsEndDate").value;
        const hText = document.getElementById("settingsHolidays").value.trim();
        appData.settings.holidays = hText ? hText.split(",").map(s => s.trim()).filter(s => s) : [];
        settingsPanel.classList.add("settings-hidden");
        restoreFromData(appData); triggerSave();
    });
    document.getElementById("clearAllBtn").addEventListener("click", () => {
        if (confirm("現在の計画を破棄し、新しい空の計画を開始しますか？")) {
            restoreFromData({ projectName: "新しい計画", settings: appData.settings, tasks: [], memo: "" });
            activeTaskId = null; activeProgressSegmentId = null; triggerSave();
            settingsPanel.classList.add("settings-hidden");
        }
    });

    document.getElementById("downloadBtn").addEventListener("click", () => {
        syncDataModel();
        const blob = new Blob([JSON.stringify(appData, null, 2)], { type: "application/json" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a"); a.href = url;
        a.download = `schedule_${formatTimestamp(new Date())}.json`;
        document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    });
    const fileInput = document.getElementById("fileInput");
    document.getElementById("uploadBtn").addEventListener("click", () => { fileInput.click(); });
    fileInput.addEventListener("change", (e) => {
        const file = e.target.files[0]; if (!file) return;
        const reader = new FileReader();
        reader.onload = (evt) => {
            try {
                const data = JSON.parse(evt.target.result);
                if (confirm("データを上書きして読み込みますか？")) { restoreFromData(data); triggerSave(); alert("完了"); }
            } catch (err) { alert("読込失敗"); }
            fileInput.value = "";
        };
        reader.readAsText(file);
    });

    const rowSelectHeader = document.getElementById("rowSelectHeader");
    if (rowSelectHeader) {
        rowSelectHeader.textContent = "□";

        rowSelectHeader.addEventListener("click", () => {
            selectionMode = (selectionMode + 1) % 3;
            
            if (selectionMode === 0) rowSelectHeader.textContent = "□";
            else if (selectionMode === 1) rowSelectHeader.textContent = "☑️";
            else if (selectionMode === 2) rowSelectHeader.textContent = "🔳"; 

            const delBtn = document.getElementById("deleteSelectedBtn");
            if(delBtn) delBtn.style.display = (selectionMode !== 0) ? "inline-block" : "none";

            taskObjects.forEach(t => {
                if (selectionMode === 1) t.isSelected = true; 
                else if (selectionMode === 2) t.isSelected = false;
                else t.isSelected = false;
                renderGrip(t);
            });
        });
    }

    const deleteSelectedBtn = document.getElementById("deleteSelectedBtn");
    if (deleteSelectedBtn) {
        deleteSelectedBtn.addEventListener("click", () => {
            const selectedTasks = taskObjects.filter(t => t.isSelected);
            if (selectedTasks.length === 0) {
                alert("削除対象が選択されていません。");
                return;
            }
            if (confirm(`${selectedTasks.length} 件の行を削除しますか？`)) {
                for (let i = taskObjects.length - 1; i >= 0; i--) {
                    if (taskObjects[i].isSelected) {
                        taskObjects[i].leftRowEl.remove();
                        taskObjects[i].rowEl.remove();
                        taskObjects.splice(i, 1);
                    }
                }
                triggerSave();
            }
        });
    }

    document.getElementById("todoColumnsInput").addEventListener("change", triggerSave);
    
    document.getElementById("outlookBtn").addEventListener("click", exportTodoToOutlookCSV);
    
    document.getElementById("todoCsvBtn").addEventListener("click", exportTodoToCSV);

    const todoBtn = document.getElementById("todoBtn");
    if(todoBtn) todoBtn.addEventListener("click", () => {
         const todoPanel = document.getElementById("todoPanel");
         currentTodoDate = new Date(); 
         todoPanel.classList.remove("settings-hidden"); 
         updateTodoTable(currentTodoDate);
    });
    
    initTodoFeature();

    if (taskMemoTextarea) {
        taskMemoTextarea.addEventListener("input", () => {
            let text = taskMemoTextarea.value || "";
            if (text.length > 3000) {
                text = text.slice(0, 3000);
                taskMemoTextarea.value = text;
            }
            updateMemoCount(text);
            const task = taskObjects.find(t => t.id === memoPanelTaskId);
            if (task) {
                task.memo = text;
                triggerSave();
            }
        });
        taskMemoTextarea.addEventListener("keydown", (e) => {
            if (e.key === "Escape") closeTaskMemoPanel();
        });
    }

    updateLeftResizeHandles();

    if (taskMemoClose) {
        taskMemoClose.addEventListener("click", () => closeTaskMemoPanel());
    }

    if (taskMemoPanel && taskMemoHeader) {
        let isDragging = false;
        let startX = 0;
        let startY = 0;
        let initL = 0;
        let initT = 0;

        taskMemoHeader.addEventListener("mousedown", (e) => {
            if (e.target.closest("button")) return;
            isDragging = true;
            startX = e.clientX;
            startY = e.clientY;
            const r = taskMemoPanel.getBoundingClientRect();
            initL = r.left;
            initT = r.top;
            document.body.style.userSelect = "none";
        });
        document.addEventListener("mousemove", (e) => {
            if (!isDragging) return;
            const left = initL + (e.clientX - startX);
            const top = initT + (e.clientY - startY);
            taskMemoPanel.style.left = `${left}px`;
            taskMemoPanel.style.top = `${top}px`;
        });
        document.addEventListener("mouseup", () => {
            if (!isDragging) return;
            isDragging = false;
            document.body.style.userSelect = "";
        });
    }

    document.addEventListener("mousemove", (e) => {
        if (isResizingCol) {
            const delta = e.clientX - resizeStartX;
            const newWidth = clampColumnWidth(resizeColIndex, resizeStartWidth + delta);
            leftColumnWidths[resizeColIndex] = newWidth;
            applyLeftColumnWidths();
        } else if (isResizingRow) {
            const task = taskObjects.find(t => t.id === resizeRowTaskId);
            if (!task) return;
            const delta = e.clientY - resizeStartY;
            const baseHeight = task.baseHeight || computeTaskBaseHeight(task);
            const newHeight = Math.max(baseHeight, resizeStartHeight + delta);
            task.customHeight = newHeight;
            task.rowEl.style.height = newHeight + "px";
            task.leftRowEl.style.height = newHeight + "px";
        }
    });
    document.addEventListener("mouseup", () => {
        if (isResizingCol) {
            isResizingCol = false;
            resizeColIndex = null;
            document.body.style.cursor = "";
            document.body.style.userSelect = "";
            triggerSave();
        }
        if (isResizingRow) {
            isResizingRow = false;
            resizeRowTaskId = null;
            document.body.style.cursor = "";
            document.body.style.userSelect = "";
            renderAllSegments();
            triggerSave();
        }
    });
}

function handleMainCellClick(task, index) {
    if (isCtrlSelectionMode) return;
    const clickedIso = timelineDays[index].iso;

    if (activeProgressSegmentId === "__main__") {
        if (!task.mainSchedule) {
            alert("この行にメインスケジュールがありません。");
            return;
        }
        task.mainSchedule.progressEndDate = clickedIso;
        activeProgressSegmentId = null;
        activeProgressTaskId = null;
        renderAllSegments();
        triggerSave();
        return;
    }

    if (!task.mainSchedule) {
        if (task.pendingMainStartIndex === null) {
            task.pendingMainStartIndex = index;
            task.pendingMainStartDate = clickedIso;
            task.pendingStartIndex = null;
            task.pendingStartDate = null;
            renderAllSegments();
            return;
        }

        const startIso = task.pendingMainStartDate;
        const endIso = clickedIso;
        const s = startIso < endIso ? startIso : endIso;
        const e = startIso < endIso ? endIso : startIso;
        task.pendingMainStartIndex = null;
        task.pendingMainStartDate = null;

        task.mainSchedule = {
            startDate: s,
            endDate: e,
            label: "メイン計画",
            startLabel: "",
            endLabel: "",
            progressEndDate: null,
            milestones: []
        };
        renderAllSegments();
        triggerSave();
        return;
    }

    const main = task.mainSchedule;
    if (clickedIso < main.startDate || clickedIso > main.endDate) {
        alert("メインスケジュールの範囲内をクリックするとマイルストーンを追加できます。開始日・終了日は線のドラッグで調整できます。");
        return;
    }

    let milestone = (main.milestones || []).find(ms => ms.date === clickedIso);
    const defaultLabel = milestone ? (milestone.label || "マイルストーン") : "マイルストーン";
    const labelInput = prompt("マイルストーン名を入力してください:", defaultLabel);
    if (labelInput === null) return;

    if (!milestone) {
        milestone = {
            id: "ms_" + Date.now() + "_" + Math.random().toString(36).slice(2),
            date: clickedIso,
            label: labelInput.trim() || "マイルストーン"
        };
        if (!Array.isArray(main.milestones)) main.milestones = [];
        main.milestones.push(milestone);
        main.milestones.sort((a, b) => a.date.localeCompare(b.date));
    } else {
        milestone.label = labelInput.trim() || "マイルストーン";
    }
    renderAllSegments();
    triggerSave();
}

function updateTodoTable(dateObj) {
    const todoDateDisplay = document.getElementById("todoDateDisplay");
    todoDateDisplay.textContent = `${dateToISO(dateObj)} (${WEEKDAYS[dateObj.getDay()]})`;
    
    const table = document.querySelector(".todo-table");
    let colgroup = table.querySelector("colgroup");
    if (colgroup) table.removeChild(colgroup); 
    colgroup = document.createElement("colgroup");
    table.insertBefore(colgroup, table.firstChild);
    
    const colWidths = [
        "40px",  // 選択
        "15%",   // 項目1
        "15%",   // 項目2
        "80px",  // 時間
        "auto",  // 実施内容
        "70px",  // 計画
        "70px"   // 実績
    ];
    colWidths.forEach(w => {
        const col = document.createElement("col");
        col.style.width = w;
        colgroup.appendChild(col);
    });

    const h1 = document.getElementById("lh1").textContent;
    const h2 = document.getElementById("lh2").textContent;
    const h3 = document.getElementById("lh3").textContent;
    
    const displayCols = ["□", h1, h2, h3, "実施内容", "計画", "実績"];
    const colKeys = ["select", "h1", "h2", "h3", "desc", "plan", "actual"];

    const thead = document.getElementById("todoThead");
    thead.innerHTML = "";
    const trH = document.createElement("tr");

    displayCols.forEach((colName, idx) => {
        const th = document.createElement("th");
        th.textContent = colName;
        if (idx === 0) {
            th.style.cursor = "pointer";
            th.style.textAlign = "center";
            th.textContent = todoSelectionState ? "☑️" : "□";
            th.addEventListener("click", () => {
                todoSelectionState = !todoSelectionState; 
                th.textContent = todoSelectionState ? "☑️" : "□";
                
                const checkboxes = document.querySelectorAll(".todo-row-checkbox");
                checkboxes.forEach(cb => {
                    cb.textContent = todoSelectionState ? "☑️" : "□";
                    cb.dataset.checked = todoSelectionState ? "true" : "false";
                });
                checkTodoDeleteBtnVisibility();
            });
        }
        trH.appendChild(th);
    });
    thead.appendChild(trH);

    const iso = dateToISO(dateObj);
    const tbody = document.getElementById("todoTableBody"); 
    tbody.innerHTML = "";
    let hasItem = false;

    // [修正] 合計計算用変数
    let totalPlan = 0;
    let totalActual = 0;

    taskObjects.forEach(task => {
        if (task.isHidden) return;
        
        const editable1 = task.leftRowEl.children[1].querySelector(".editable");
        const editable2 = task.leftRowEl.children[2].querySelector(".editable");
        const editable3 = task.leftRowEl.children[3].querySelector(".editable");

        const t1 = editable1.textContent;
        const t2 = editable2.textContent;
        const t3 = editable3.textContent;

        task.segments.forEach(seg => {
            if (seg.startDate <= iso && seg.endDate >= iso) {
                const tr = document.createElement("tr");
                tr.dataset.taskId = task.id;
                tr.dataset.segId = seg.id;

                // [修正] 合計計算
                const pv = (seg.dailyValues || {})[iso];
                const av = (seg.dailyResults || {})[iso];
                if (pv && !isNaN(parseFloat(pv))) totalPlan += parseFloat(pv);
                if (av && !isNaN(parseFloat(av))) totalActual += parseFloat(av);

                colKeys.forEach((key) => {
                    const td = document.createElement("td");
                    
                    if (key === "select") {
                        td.style.textAlign = "center";
                        td.style.cursor = "pointer";
                        td.className = "todo-row-checkbox";
                        td.textContent = todoSelectionState ? "☑️" : "□"; 
                        td.dataset.checked = todoSelectionState ? "true" : "false";
                        
                        td.addEventListener("click", (e) => {
                            e.stopPropagation();
                            const isChecked = td.dataset.checked === "true";
                            td.dataset.checked = isChecked ? "false" : "true";
                            td.textContent = (td.dataset.checked === "true") ? "☑️" : "□";
                            checkTodoDeleteBtnVisibility();
                        });
                    } else {
                        const input = document.createElement("input");
                        let val = "";

                        if (key === "h1") {
                            val = t1;
                            input.addEventListener("change", (e) => { editable1.textContent = e.target.value; triggerSave(); });
                        }
                        else if (key === "h2") {
                            val = t2;
                            input.addEventListener("change", (e) => { editable2.textContent = e.target.value; triggerSave(); });
                        }
                        else if (key === "h3") {
                            val = t3;
                            input.style.textAlign = "center";
                            input.addEventListener("change", (e) => { editable3.textContent = e.target.value; triggerSave(); });
                        }
                        else if (key === "desc") {
                            val = seg.label || "";
                            input.addEventListener("change", (e) => { seg.label = e.target.value; renderAllSegments(); triggerSave(); });
                        }
                        else if (key === "plan") {
                            val = (seg.dailyValues || {})[iso] || "";
                            input.style.textAlign = "center";
                            input.addEventListener("change", (e) => {
                                if(!seg.dailyValues) seg.dailyValues = {};
                                seg.dailyValues[iso] = e.target.value;
                                if(!e.target.value) delete seg.dailyValues[iso];
                                renderAllSegments();
                                triggerSave();
                                updateTodoTable(dateObj); // 合計再計算のため
                            });
                        }
                        else if (key === "actual") {
                            val = (seg.dailyResults || {})[iso] || "";
                            input.style.textAlign = "center";
                            input.addEventListener("change", (e) => {
                                if(!seg.dailyResults) seg.dailyResults = {};
                                seg.dailyResults[iso] = e.target.value;
                                if(!e.target.value) delete seg.dailyResults[iso];
                                triggerSave();
                                updateTodoTable(dateObj); // 合計再計算のため
                            });
                        }
                        input.value = val;
                        td.appendChild(input);
                    }
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
                hasItem = true;
            }
        });
    });
    document.getElementById("todoEmptyMsg").style.display = hasItem ? "none" : "block";

    // [修正] 合計行（tfoot）の追加
    let tfoot = table.querySelector("tfoot");
    if(tfoot) table.removeChild(tfoot);
    tfoot = document.createElement("tfoot");
    const trF = document.createElement("tr");
    
    const tdLabel = document.createElement("td");
    tdLabel.colSpan = 5; 
    tdLabel.textContent = "合計";
    tdLabel.style.textAlign = "right";
    trF.appendChild(tdLabel);

    const tdPlan = document.createElement("td");
    tdPlan.textContent = (totalPlan % 1 === 0) ? totalPlan : totalPlan.toFixed(1);
    tdPlan.style.textAlign = "center";
    trF.appendChild(tdPlan);

    const tdActual = document.createElement("td");
    tdActual.textContent = (totalActual % 1 === 0) ? totalActual : totalActual.toFixed(1);
    tdActual.style.textAlign = "center";
    trF.appendChild(tdActual);

    tfoot.appendChild(trF);
    table.appendChild(tfoot);

    checkTodoDeleteBtnVisibility();
}

function checkTodoDeleteBtnVisibility() {
    const delBtn = document.getElementById("todoDeleteBtn");
    if(!delBtn) return;
    const checkedItems = document.querySelectorAll(".todo-row-checkbox[data-checked='true']");
    delBtn.style.display = (checkedItems.length > 0) ? "inline-block" : "none";
}

function exportTodoToOutlookCSV() {
    const iso = dateToISO(currentTodoDate);
    const dateStr = iso.replace(/-/g, '/');

    const items = [];
    taskObjects.forEach(task => {
        if (task.isHidden) return;
        const t1 = task.leftRowEl.children[1].querySelector(".editable").textContent;
        task.segments.forEach(seg => {
            if (seg.startDate <= iso && seg.endDate >= iso) {
                items.push({
                    item1: t1,
                    desc: seg.label || ""
                });
            }
        });
    });

    if (items.length === 0) {
        alert("出力するデータがありません");
        return;
    }

    const headers = ["件名","開始日","開始時刻","終了日","終了時刻","プライベート","公開する時間帯の種類","秘密度","優先度"];
    const rows = [];

    let currentMin = 510; 

    items.forEach((item) => {
        const subject = `${item.item1}：${item.desc}`;

        const hStart = Math.floor(currentMin / 60);
        const mStart = currentMin % 60;
        const startTimeStr = `${hStart}:${pad2(mStart)}:00`;

        const endMin = currentMin + 30;
        const hEnd = Math.floor(endMin / 60);
        const mEnd = endMin % 60;
        const endTimeStr = `${hEnd}:${pad2(mEnd)}:00`;

        currentMin += 30;

        const rowData = [
            `"${subject.replace(/"/g, '""')}"`,
            `"${dateStr}"`,
            `"${startTimeStr}"`,
            `"${dateStr}"`,
            `"${endTimeStr}"`,
            `"FALSE"`,
            `"2"`,
            `"標準"`,
            `"標準"`
        ];
        rows.push(rowData.join(","));
    });

    const csvContent = headers.map(h => `"${h}"`).join(",") + "\r\n" + rows.join("\r\n");

    downloadAsShiftJIS(csvContent, `Outlook_${formatTimestamp(new Date())}.csv`);
}

function exportTodoToCSV() {
    const filename = `ToDoList_${formatTimestamp(new Date())}.csv`;
    const tbody = document.getElementById("todoTableBody");

    const columnsRaw = document.getElementById("todoColumnsInput").value || "";
    let columns = columnsRaw.split(",").map(s => s.trim()).filter(Boolean);
    if (columns.length === 0) {
        columns = DEFAULT_TODO_COLUMNS.split(",").map(s => s.trim());
    }

    const headers = columns.map(c => '"' + c.replace(/"/g, '""') + '"');
    const iso = dateToISO(currentTodoDate);

    const rows = [];
    tbody.querySelectorAll("tr").forEach(tr => {
        const taskId = tr.dataset.taskId;
        const segId = tr.dataset.segId;
        const task = taskObjects.find(t => t.id === taskId);
        const seg = task ? task.segments.find(s => s.id === segId) : null;

        const t1 = task ? task.leftRowEl.children[1].querySelector(".editable").textContent : "";
        const t2 = task ? task.leftRowEl.children[2].querySelector(".editable").textContent : "";
        const t3 = task ? task.leftRowEl.children[3].querySelector(".editable").textContent : "";
        const desc = seg ? (seg.label || "") : "";
        const plan = seg && seg.dailyValues ? (seg.dailyValues[iso] || "") : "";
        const actual = seg && seg.dailyResults ? (seg.dailyResults[iso] || "") : "";
        const memo = task ? (task.memo || "") : "";

        const rowData = columns.map(col => {
            let v = "";
            if (col === "項目1") v = t1;
            else if (col === "項目2") v = t2;
            else if (col === "時間") v = t3;
            else if (col === "実施内容") v = desc;
            else if (col === "計画") v = plan;
            else if (col === "実績") v = actual;
            else if (col === "メモ") v = memo;
            return '"' + String(v).replace(/"/g, '""') + '"';
        });
        rows.push(rowData.join(","));
    });

    if (rows.length === 0) {
        alert("出力するデータがありません");
        return;
    }

    const csvContent = headers.join(",") + "\r\n" + rows.join("\r\n");
    downloadAsShiftJIS(csvContent, filename);
}

function downloadAsShiftJIS(content, filename) {
    if (typeof Encoding === "undefined") {
        alert("文字コード変換ライブラリが読み込まれていません。インターネット接続を確認してください。\nとりあえずUTF-8(BOM付)で出力します。");
        const blob = new Blob(["\uFEFF" + content], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        return;
    }

    const unicodeList = [];
    for (let i = 0; i < content.length; i++) {
        unicodeList.push(content.charCodeAt(i));
    }
    
    const sjisCodeList = Encoding.convert(unicodeList, {
        to: 'SJIS',
        from: 'UNICODE'
    });
    
    const u8Array = new Uint8Array(sjisCodeList);
    const blob = new Blob([u8Array], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function ensureTodoDateControlLayout() {
    const todoPanel = document.getElementById("todoPanel");
    const prevBtn = document.getElementById("todoPrevDay");
    const nextBtn = document.getElementById("todoNextDay");
    const todayBtn = document.getElementById("todoTodayBtn");
    let label = document.getElementById("todoDateLabel")
        || todoPanel?.querySelector(".todo-date-label, .todo-date, [data-todo-date]");
    const controlsRow = prevBtn?.parentElement;
    if (!controlsRow || !prevBtn || !nextBtn || !todayBtn) {
        return label || null;
    }

    let group = document.getElementById("todoDateControlGroup");
    if (!group) {
        group = document.createElement("div");
        group.id = "todoDateControlGroup";
        controlsRow.innerHTML = "";
        controlsRow.appendChild(group);
    }

    controlsRow.style.display = "flex";
    controlsRow.style.justifyContent = "center";
    controlsRow.style.alignItems = "center";
    controlsRow.style.width = "100%";
    controlsRow.style.minWidth = "0";
    controlsRow.style.padding = "0 8px";
    controlsRow.style.boxSizing = "border-box";

    group.style.display = "flex";
    group.style.alignItems = "center";
    group.style.justifyContent = "center";
    group.style.gap = "10px";
    group.style.width = "fit-content";
    group.style.maxWidth = "100%";
    group.style.flex = "0 1 auto";
    group.style.minWidth = "0";

    if (!label) {
        label = document.createElement("div");
        label.id = "todoDateLabel";
        label.className = "todo-date-label";
    }

    [prevBtn, label, nextBtn, todayBtn].forEach(el => {
        if (el.parentElement !== group) {
            group.appendChild(el);
        }
    });

    prevBtn.style.flex = "0 0 auto";
    nextBtn.style.flex = "0 0 auto";
    todayBtn.style.flex = "0 0 auto";
    prevBtn.style.margin = "0";
    nextBtn.style.margin = "0";
    todayBtn.style.margin = "0";

    label.style.margin = "0";
    label.style.fontWeight = "600";
    label.style.fontSize = "clamp(15px, 2.6vw, 18px)";
    label.style.lineHeight = "1.2";
    label.style.textAlign = "center";
    label.style.whiteSpace = "nowrap";
    label.style.flex = "0 0 auto";
    label.style.minWidth = "210px";

    return label;
}

function updateTodoTable(dateObj) {
    const todoPanel = document.getElementById("todoPanel");
    let label = ensureTodoDateControlLayout()
        || document.getElementById("todoDateLabel")
        || todoPanel?.querySelector(".todo-date-label, .todo-date, [data-todo-date]");
    const table = document.getElementById("todoTable")
        || todoPanel?.querySelector("table");
    const emptyMsg = document.getElementById("todoEmptyMsg")
        || todoPanel?.querySelector(".todo-empty, .empty-message, [data-todo-empty]");
    if (!table) return;

    let thead = table.querySelector("thead");
    if (!thead) {
        thead = document.createElement("thead");
        table.prepend(thead);
    }

    let tbody = document.getElementById("todoTableBody") || table.querySelector("tbody");
    if (!tbody) {
        tbody = document.createElement("tbody");
        table.appendChild(tbody);
    }
    if (!tbody.id) tbody.id = "todoTableBody";

    const dateText = `${dateObj.getFullYear()}年${dateObj.getMonth() + 1}月${dateObj.getDate()}日(${WEEKDAYS[dateObj.getDay()]})`;
    if (label) label.textContent = dateText;

    const columnsInput = document.getElementById("todoColumnsInput");
    const columnsRaw = columnsInput ? columnsInput.value : "";
    let displayCols = columnsRaw.split(",").map(s => s.trim()).filter(Boolean);
    if (displayCols.length === 0) {
        displayCols = DEFAULT_TODO_COLUMNS.split(",").map(s => s.trim());
    }
    displayCols = displayCols.filter(col => col !== "メモ");
    const requiredTodoCols = ["項目1", "項目2", "時間", "実施内容", "計画", "実績"];
    requiredTodoCols.forEach(col => {
        if (!displayCols.includes(col)) displayCols.push(col);
    });

    const colKeys = displayCols.map(col => {
        if (col === "項目1") return "h1";
        if (col === "項目2") return "h2";
        if (col === "時間") return "h3";
        if (col === "実施内容") return "desc";
        if (col === "計画") return "plan";
        if (col === "実績") return "actual";
        if (col === "メモ") return "memo";
        return "unknown";
    });
    const colWidths = {
        h1: "90px",
        h2: "150px",
        h3: "90px",
        desc: "120px",
        plan: "70px",
        actual: "70px",
        memo: "100px",
        unknown: "100px"
    };

    table.style.tableLayout = "fixed";
    table.style.width = "100%";
    thead.innerHTML = "";
    const trH = document.createElement("tr");
    displayCols.forEach((colName, idx) => {
        const th = document.createElement("th");
        th.textContent = colName;
        th.style.width = colWidths[colKeys[idx]] || colWidths.unknown;
        trH.appendChild(th);
    });
    thead.appendChild(trH);

    const iso = dateToISO(dateObj);
    tbody.innerHTML = "";

    let hasItem = false;
    let totalPlan = 0;
    let totalActual = 0;

    taskObjects.forEach(task => {
        if (task.isHidden) return;

        const editable1 = task.leftRowEl.children[1].querySelector(".editable");
        const editable2 = task.leftRowEl.children[2].querySelector(".editable");
        const editable3 = task.leftRowEl.children[3].querySelector(".editable");
        const t1 = editable1 ? editable1.textContent : "";
        const t2 = editable2 ? editable2.textContent : "";
        const t3 = editable3 ? editable3.textContent : "";
        const activeSegments = task.segments.filter(seg => seg.startDate <= iso && seg.endDate >= iso);
        const rowDefs = activeSegments.length
            ? activeSegments.map(seg => ({ seg, hasTask: true }))
            : [{ seg: null, hasTask: false }];

        rowDefs.forEach(({ seg, hasTask }) => {
            const tr = document.createElement("tr");
            tr.dataset.taskId = task.id;
            if (seg) tr.dataset.segId = seg.id;

            const planValue = seg ? ((seg.dailyValues || {})[iso] || "") : "";
            const actualValue = seg ? ((seg.dailyResults || {})[iso] || "") : "";
            if (planValue && !isNaN(parseFloat(planValue))) totalPlan += parseFloat(planValue);
            if (actualValue && !isNaN(parseFloat(actualValue))) totalActual += parseFloat(actualValue);

            colKeys.forEach((key) => {
                const td = document.createElement("td");
                td.style.width = colWidths[key] || colWidths.unknown;

                const input = document.createElement("input");
                let val = "";
                let isEditable = false;
                input.style.width = "100%";
                input.style.boxSizing = "border-box";

                if (key === "h1") {
                    val = t1;
                    isEditable = !!editable1;
                    if (editable1) {
                        input.addEventListener("change", (e) => {
                            editable1.textContent = e.target.value;
                            triggerSave();
                        });
                    }
                } else if (key === "h2") {
                    val = hasTask ? t2 : "";
                    if (hasTask && editable2) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            editable2.textContent = e.target.value;
                            triggerSave();
                        });
                    }
                } else if (key === "h3") {
                    val = hasTask ? t3 : "";
                    input.style.textAlign = "center";
                    if (hasTask && editable3) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            editable3.textContent = e.target.value;
                            triggerSave();
                        });
                    }
                } else if (key === "desc") {
                    val = hasTask && seg ? (seg.label || "") : "";
                    if (hasTask && seg) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            seg.label = e.target.value;
                            renderAllSegments();
                            triggerSave();
                        });
                    }
                } else if (key === "plan") {
                    val = planValue;
                    input.style.textAlign = "center";
                    if (hasTask && seg) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            if (!seg.dailyValues) seg.dailyValues = {};
                            seg.dailyValues[iso] = e.target.value;
                            if (!e.target.value) delete seg.dailyValues[iso];
                            renderAllSegments();
                            triggerSave();
                            updateTodoTable(dateObj);
                        });
                    }
                } else if (key === "actual") {
                    val = actualValue;
                    input.style.textAlign = "center";
                    if (hasTask && seg) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            if (!seg.dailyResults) seg.dailyResults = {};
                            seg.dailyResults[iso] = e.target.value;
                            if (!e.target.value) delete seg.dailyResults[iso];
                            triggerSave();
                            updateTodoTable(dateObj);
                        });
                    }
                } else if (key === "memo") {
                    val = hasTask ? (task.memo || "") : "";
                    if (hasTask) {
                        isEditable = true;
                        input.addEventListener("change", (e) => {
                            task.memo = e.target.value;
                            triggerSave();
                        });
                    }
                }

                input.value = val;
                if (!isEditable) {
                    input.readOnly = true;
                    input.tabIndex = -1;
                }
                td.appendChild(input);
                tr.appendChild(td);
            });

            tbody.appendChild(tr);
            hasItem = true;
        });
    });

    if (emptyMsg) {
        emptyMsg.style.display = hasItem ? "none" : "block";
    }

    let tfoot = table.querySelector("tfoot");
    if (tfoot) table.removeChild(tfoot);
    tfoot = document.createElement("tfoot");
    const trF = document.createElement("tr");

    const tdLabel = document.createElement("td");
    tdLabel.colSpan = Math.max(1, colKeys.length - 2);
    tdLabel.textContent = "合計";
    tdLabel.style.textAlign = "right";
    trF.appendChild(tdLabel);

    const tdPlan = document.createElement("td");
    tdPlan.textContent = (totalPlan % 1 === 0) ? totalPlan : totalPlan.toFixed(1);
    tdPlan.style.textAlign = "center";
    trF.appendChild(tdPlan);

    const tdActual = document.createElement("td");
    tdActual.textContent = (totalActual % 1 === 0) ? totalActual : totalActual.toFixed(1);
    tdActual.style.textAlign = "center";
    trF.appendChild(tdActual);

    tfoot.appendChild(trF);
    table.appendChild(tfoot);
    checkTodoDeleteBtnVisibility();
}

function exportTodoToCSV() {
    const filename = `ToDoList_${formatTimestamp(new Date())}.csv`;
    const tbody = document.getElementById("todoTableBody");

    const columnsRaw = document.getElementById("todoColumnsInput").value || "";
    let columns = columnsRaw.split(",").map(s => s.trim()).filter(Boolean);
    if (columns.length === 0) {
        columns = DEFAULT_TODO_COLUMNS.split(",").map(s => s.trim());
    }

    const headers = columns.map(c => '"' + c.replace(/"/g, '""') + '"');
    const iso = dateToISO(currentTodoDate);
    const rows = [];

    tbody.querySelectorAll("tr").forEach(tr => {
        const taskId = tr.dataset.taskId;
        const segId = tr.dataset.segId;
        const task = taskObjects.find(t => t.id === taskId);
        const seg = task ? task.segments.find(s => s.id === segId) : null;
        const hasTask = !!seg;

        const t1 = task ? task.leftRowEl.children[1].querySelector(".editable").textContent : "";
        const t2 = hasTask && task ? task.leftRowEl.children[2].querySelector(".editable").textContent : "";
        const t3 = hasTask && task ? task.leftRowEl.children[3].querySelector(".editable").textContent : "";
        const desc = hasTask && seg ? (seg.label || "") : "";
        const plan = hasTask && seg && seg.dailyValues ? (seg.dailyValues[iso] || "") : "";
        const actual = hasTask && seg && seg.dailyResults ? (seg.dailyResults[iso] || "") : "";
        const memo = hasTask && task ? (task.memo || "") : "";

        const rowData = columns.map(col => {
            let v = "";
            if (col === "項目1") v = t1;
            else if (col === "項目2") v = t2;
            else if (col === "時間") v = t3;
            else if (col === "実施内容") v = desc;
            else if (col === "計画") v = plan;
            else if (col === "実績") v = actual;
            else if (col === "メモ") v = memo;
            return '"' + String(v).replace(/"/g, '""') + '"';
        });
        rows.push(rowData.join(","));
    });

    if (rows.length === 0) {
        alert("出力するデータがありません");
        return;
    }

    const csvContent = headers.join(",") + "\r\n" + rows.join("\r\n");
    downloadAsShiftJIS(csvContent, filename);
}

function addTodoRow(dateObj) {
    addTaskRow();
    const newTask = taskObjects[taskObjects.length - 1];
    const iso = dateToISO(dateObj);
    const newSeg = {
        id: "seg_" + Date.now() + "_" + Math.random().toString(36).slice(2),
        startDate: iso,
        endDate: iso, 
        type: "point",
        label: "", 
        progressEndDate: null,
        dailyValues: {},
        dailyResults: {}
    };
    newTask.segments.push(newSeg);
    renderAllSegments();
    triggerSave();
    updateTodoTable(dateObj);
}

function initTodoFeature() {
    const todoPanel = document.getElementById("todoPanel");
    if (!todoPanel) return;
    ensureTodoDateControlLayout();
    
    const update = () => updateTodoTable(currentTodoDate);

    document.getElementById("todoCloseBtn").addEventListener("click", () => todoPanel.classList.add("settings-hidden"));
    
    document.getElementById("todoPrevDay").addEventListener("click", () => { 
        currentTodoDate.setDate(currentTodoDate.getDate() - 1); 
        todoSelectionState = false; 
        update(); 
    });
    document.getElementById("todoNextDay").addEventListener("click", () => { 
        currentTodoDate.setDate(currentTodoDate.getDate() + 1); 
        todoSelectionState = false;
        update(); 
    });
    document.getElementById("todoTodayBtn").addEventListener("click", () => { 
        currentTodoDate = new Date(); 
        todoSelectionState = false;
        update(); 
    });

    document.getElementById("todoAddRowBtn").addEventListener("click", () => {
        addTodoRow(currentTodoDate);
    });
    
    const footerControls = document.querySelector(".todo-footer > div:nth-child(2)");
    if (!document.getElementById("todoDeleteBtn")) {
        const delBtn = document.createElement("button");
        delBtn.id = "todoDeleteBtn";
        delBtn.className = "btn-secondary";
        delBtn.style.color = "#ef4444";
        delBtn.style.borderColor = "#fca5a5";
        delBtn.style.fontSize = "12px";
        delBtn.style.marginLeft = "8px";
        delBtn.textContent = "🗑️ 選択行を削除";
        delBtn.style.display = "none"; 
        
        delBtn.addEventListener("click", () => {
            const checkboxes = document.querySelectorAll(".todo-row-checkbox[data-checked='true']");
            if (checkboxes.length === 0) {
                alert("削除する項目を選択してください。");
                return;
            }
            
            if (confirm(`${checkboxes.length} 件の項目を削除しますか？`)) {
                const itemsToDelete = [];
                checkboxes.forEach(cb => {
                    const tr = cb.closest("tr");
                    itemsToDelete.push({ taskId: tr.dataset.taskId, segId: tr.dataset.segId });
                });

                let changeOccurred = false;
                itemsToDelete.forEach(item => {
                    const task = taskObjects.find(t => t.id === item.taskId);
                    if (task) {
                        const originalLen = task.segments.length;
                        task.segments = task.segments.filter(s => s.id !== item.segId);
                        if (task.segments.length !== originalLen) changeOccurred = true;
                    }
                });

                if (changeOccurred) {
                    renderAllSegments();
                    triggerSave();
                    update(); 
                }
            }
        });
        
        const addBtn = document.getElementById("todoAddRowBtn");
        if(addBtn) {
            addBtn.insertAdjacentElement('afterend', delBtn);
        }
    }

    const win = todoPanel.querySelector(".todo-window"), header = todoPanel.querySelector(".todo-header");
    if (win) {
        win.style.minWidth = "540px";
    }
    let isDragging = false, startX, startY, initL, initT;
    header.addEventListener("mousedown", (e) => {
        if(e.target.closest("button")) return;
        isDragging = true; startX = e.clientX; startY = e.clientY;
        const r = win.getBoundingClientRect(); initL = r.left; initT = r.top;
        header.style.cursor = "grabbing"; document.body.style.userSelect = "none";
    });
    document.addEventListener("mousemove", (e) => {
        if (!isDragging) return;
        win.style.left = (initL + e.clientX - startX) + "px"; 
        win.style.top = (initT + e.clientY - startY) + "px";
        
        win.style.width = win.offsetWidth + "px";
        win.style.height = win.offsetHeight + "px";
    });
    document.addEventListener("mouseup", () => { isDragging = false; header.style.cursor = "grab"; document.body.style.userSelect = ""; });
}

window.addEventListener("resize", renderAllSegments);
initializeApp();
