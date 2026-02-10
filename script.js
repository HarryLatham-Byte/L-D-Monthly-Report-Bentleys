/**
 * L&D Monthly Summary Report Dashboard
 * Core Logic & State Management - CSV Enhanced
 */

// --- State Management ---
const state = {
    theme: 'light',
    rawData: [],      // Complete array of rows from CSV
    filteredData: [], // Filtered subset based on current selections
    filters: {
        office: 'all',
        department: 'all',
        type: 'all',
        platform: 'all',
        team: 'all',
        name: '',
        startDate: '',
        endDate: ''
    },
    charts: {}
};

// Register Chart.js DataLabels plugin
Chart.register(ChartDataLabels);

// --- Theme Handling ---
function toggleTheme() {
    state.theme = state.theme === 'light' ? 'dark' : 'light';
    document.body.classList.toggle('dark-mode', state.theme === 'dark');

    document.getElementById('moonIcon').style.display = state.theme === 'light' ? 'block' : 'none';
    document.getElementById('sunIcon').style.display = state.theme === 'dark' ? 'block' : 'none';

    updateChartThemes();
}

function updateChartThemes() {
    const isDark = state.theme === 'dark';
    const gridColor = isDark ? '#334155' : '#e3e8ee';
    const textColor = isDark ? '#94a3b8' : '#4f566b';

    Object.values(state.charts).forEach(chart => {
        if (chart.options.scales) {
            if (chart.options.scales.x) {
                chart.options.scales.x.grid.color = gridColor;
                chart.options.scales.x.ticks.color = textColor;
            }
            if (chart.options.scales.y) {
                chart.options.scales.y.grid.color = gridColor;
                chart.options.scales.y.ticks.color = textColor;
            }
        }
        if (chart.options.plugins && chart.options.plugins.legend) {
            chart.options.plugins.legend.labels.color = textColor;
        }
        if (chart.options.plugins && chart.options.plugins.datalabels) {
            chart.options.plugins.datalabels.color = textColor;
        }
        chart.update();
    });
}

// --- Chart Colors ---
const COLORS = {
    primary: '#f7941d', // Orange
    secondary: '#ffad4d', // Light Orange
    accent: '#36b37e',
    backgrounds: [
        'rgba(247, 148, 29, 0.8)',
        'rgba(255, 173, 77, 0.8)',
        'rgba(54, 179, 126, 0.8)',
        'rgba(255, 171, 0, 0.8)',
        'rgba(255, 86, 48, 0.8)',
        'rgba(101, 84, 192, 0.8)',
        'rgba(255, 153, 31, 0.8)'
    ]
};

// --- Initialization ---
document.addEventListener('DOMContentLoaded', () => {
    initUI();
    loadFiltersFromURL();
    setupEventListeners();
    setupHelpTriggers();
    // No longer loading local data automatically
    // loadLocalData(); 
});

function initUI() {
    // Theme initialization
    if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
        // toggleTheme(); // Uncomment to follow OS preference
    }
}

function setupEventListeners() {
    document.getElementById('themeToggle').addEventListener('click', toggleTheme);

    // Filter Listeners
    const filterIds = ['officeFilter', 'departmentFilter', 'trainingTypeFilter', 'platformFilter', 'teamFilter'];
    filterIds.forEach(id => {
        document.getElementById(id).addEventListener('change', (e) => {
            const filterKey = id.replace('Filter', '').toLowerCase();
            state.filters[id === 'trainingTypeFilter' ? 'type' : filterKey] = e.target.value;
            applyFilters(true);
        });
    });

    // Name Search
    const nameSearch = document.getElementById('nameSearch');
    nameSearch.addEventListener('input', (e) => {
        state.filters.name = e.target.value.trim().toLowerCase();
        applyFilters(true);
    });
    nameSearch.addEventListener('focus', () => {
        nameSearch.setAttribute('placeholder', 'Type to search...');
    });

    // Date Filters
    ['startDate', 'endDate'].forEach(id => {
        document.getElementById(id).addEventListener('change', (e) => {
            state.filters[id] = e.target.value;
            applyFilters(true);
        });
    });

    // File Upload
    const uploadBtn = document.getElementById('uploadBtn');
    const excelUpload = document.getElementById('excelUpload');
    if (uploadBtn && excelUpload) {
        uploadBtn.addEventListener('click', () => excelUpload.click());
        excelUpload.addEventListener('change', handleFileUpload);
    }

    // Export & Action Buttons
    document.getElementById('exportPDF')?.addEventListener('click', () => window.print());
    document.getElementById('resetFilters')?.addEventListener('click', resetFilters);

    // Modal Close
    document.getElementById('closeModal')?.addEventListener('click', closeModal);
    window.addEventListener('click', (e) => {
        if (e.target.classList.contains('modal-overlay')) closeModal();
    });
}

function setupHelpTriggers() {
    document.querySelectorAll('.help-trigger').forEach(trigger => {
        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            const msg = trigger.getAttribute('data-help');
            alert(msg); // Simple alert for now, can be styled further
        });
    });
}

function resetFilters() {
    state.filters = {
        office: 'all',
        department: 'all',
        type: 'all',
        platform: 'all',
        team: 'all',
        name: '',
        startDate: '',
        endDate: ''
    };

    // Reset UI elements
    const filterIds = ['officeFilter', 'departmentFilter', 'trainingTypeFilter', 'platformFilter', 'teamFilter'];
    filterIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = 'all';
    });

    const nameSearch = document.getElementById('nameSearch');
    if (nameSearch) {
        nameSearch.value = '';
        nameSearch.setAttribute('placeholder', 'Type a name...');
    }

    const dateIds = ['startDate', 'endDate'];
    dateIds.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });

    applyFilters(true);
}

// --- Data Loading ---
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    showError("Reading file...");
    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];

            // Convert to JSON with empty rows handled
            const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            if (json.length === 0) {
                throw new Error("The uploaded file appears to be empty.");
            }

            state.rawData = json;
            populateFilters();
            applyFilters();
            updatePreparedDate();
            hideError();

            // Visual feedback
            const uploadBtn = document.getElementById('uploadBtn');
            if (uploadBtn) {
                uploadBtn.innerHTML = `<span>Change Report</span>`;
                uploadBtn.classList.add('btn-success');
            }
        } catch (err) {
            console.error("File processing error:", err);
            showError("Error processing file: " + err.message);
        }
    };

    reader.onerror = function () {
        showError("Failed to read file.");
    };

    reader.readAsArrayBuffer(file);
}

// --- Filter Management ---
function populateFilters() {
    if (state.rawData.length === 0) return;

    // Detect column name for "Name" (handle BOM or case variations)
    const sampleRow = state.rawData[0];
    const nameKey = Object.keys(sampleRow).find(k => k.toLowerCase().includes('name') && !k.toLowerCase().includes('manager')) || 'Name';
    state.nameKey = nameKey; // Store for later use

    const teamKey = Object.keys(sampleRow).find(k => k.toLowerCase() === 'team') || 'Team';
    state.teamKey = teamKey;

    const offices = [...new Set(state.rawData.map(row => row.Office))].filter(Boolean).sort();
    const departments = [...new Set(state.rawData.map(row => row.Department))].filter(Boolean).sort();
    const teams = [...new Set(state.rawData.map(row => row[teamKey]))].filter(Boolean).sort();
    const types = [...new Set(state.rawData.map(row => row['Training Type']))].filter(Boolean).sort();
    const platforms = [...new Set(state.rawData.map(row => row.Platform))].filter(Boolean).sort();
    const names = [...new Set(state.rawData.map(row => row[nameKey]))].filter(Boolean).sort();

    // Detect column name for "Completion" (handle BOM or case variations)
    const dateKey = Object.keys(sampleRow).find(k => k.trim().toLowerCase() === 'completion' || k.trim().toLowerCase() === 'completion date') || 'Completion';
    state.dateKey = dateKey;

    updateSelectOptions('officeFilter', offices, 'All Offices');
    updateSelectOptions('departmentFilter', departments, 'All Departments');
    updateSelectOptions('teamFilter', teams, 'All Teams');
    updateSelectOptions('trainingTypeFilter', types, 'All Types');
    updateSelectOptions('platformFilter', platforms, 'All Platforms');

    // Populate name datalist
    const nameList = document.getElementById('nameOptions');
    if (nameList) {
        nameList.innerHTML = '';
        names.forEach(name => {
            const opt = document.createElement('option');
            opt.value = name;
            nameList.appendChild(opt);
        });
    }

    initDateRangeInputs();
}

/**
 * Find the full range of dates in the data and set the input constraints
 */
function initDateRangeInputs() {
    const dates = state.rawData
        .map(row => parseDateValue(row[state.dateKey]))
        .filter(Boolean);

    if (dates.length === 0) return;

    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));

    const toLocalISODate = (d) => {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${y}-${m}-${day}`;
    };

    const startInput = document.getElementById('startDate');
    const endInput = document.getElementById('endDate');

    if (startInput && endInput) {
        startInput.min = toLocalISODate(minDate);
        startInput.max = toLocalISODate(maxDate);
        endInput.min = toLocalISODate(minDate);
        endInput.max = toLocalISODate(maxDate);

        // Reset to empty by default as requested
        startInput.value = '';
        endInput.value = '';

        // Sync state to empty
        state.filters.startDate = '';
        state.filters.endDate = '';
    }
}

function updateSelectOptions(id, options, defaultLabel) {
    const select = document.getElementById(id);
    select.innerHTML = `<option value="all">${defaultLabel}</option>`;
    options.forEach(opt => {
        const el = document.createElement('option');
        el.value = opt;
        el.textContent = opt;
        select.appendChild(el);
    });
}

function applyFilters(syncURL = false) {
    const nameKey = state.nameKey || 'Name';
    state.filteredData = state.rawData.filter(row => {
        const matchOffice = state.filters.office === 'all' || row.Office === state.filters.office;
        const matchDept = state.filters.department === 'all' || row.Department === state.filters.department;
        const matchTeam = state.filters.team === 'all' || row[state.teamKey || 'Team'] === state.filters.team;
        const matchType = state.filters.type === 'all' || row['Training Type'] === state.filters.type;
        const matchPlatform = state.filters.platform === 'all' || row.Platform === state.filters.platform;

        // Name search (case insensitive)
        const rowName = row[nameKey] ? String(row[nameKey]).toLowerCase() : '';
        const matchName = !state.filters.name || rowName.includes(state.filters.name);

        // Date range filtering
        let matchDate = true;
        if (state.filters.startDate || state.filters.endDate) {
            const rowDate = parseDateValue(row[state.dateKey]);
            if (rowDate) {
                if (state.filters.startDate) {
                    const startParts = state.filters.startDate.split('-');
                    const start = new Date(startParts[0], startParts[1] - 1, startParts[2], 0, 0, 0);
                    if (rowDate < start) matchDate = false;
                }
                if (state.filters.endDate) {
                    const endParts = state.filters.endDate.split('-');
                    const end = new Date(endParts[0], endParts[1] - 1, endParts[2], 23, 59, 59);
                    if (rowDate > end) matchDate = false;
                }
            } else {
                matchDate = false;
            }
        }

        return matchOffice && matchDept && matchTeam && matchType && matchPlatform && matchName && matchDate;
    });

    if (syncURL) syncFiltersToURL();
    renderDashboard();
}

/**
 * Persist filters to URL hash for shareable links
 */
function syncFiltersToURL() {
    const params = new URLSearchParams();
    Object.entries(state.filters).forEach(([key, val]) => {
        if (val && val !== 'all') params.set(key, val);
    });
    window.location.hash = params.toString();
}

function loadFiltersFromURL() {
    const hash = window.location.hash.substring(1);
    if (!hash) return;
    const params = new URLSearchParams(hash);
    params.forEach((val, key) => {
        if (state.filters.hasOwnProperty(key)) {
            state.filters[key] = val;
            // Update UI elements
            const el = document.getElementById(key === 'type' ? 'trainingTypeFilter' :
                (key === 'name' ? 'nameSearch' :
                    (key.endsWith('Filter') ? key : key + 'Filter')));
            if (el) el.value = val;
            if (key === 'name' && document.getElementById('nameSearch')) {
                document.getElementById('nameSearch').value = val;
            }
        }
    });
}

/**
 * Helper to turn any CSV date value into a JS Date object
 */
function parseDateValue(val) {
    if (!val) return null;
    if (val instanceof Date) return val;

    // Handle numeric Excel Serial Dates (e.g., 45321)
    const num = Number(val);
    if (!isNaN(num) && num > 40000 && num < 60000) {
        // Excel base date: Dec 30, 1899
        return new Date((num - 25569) * 86400 * 1000);
    }

    const str = String(val).trim();
    if (!str || str.startsWith('#VALUE')) return null;

    // Handle DD/MM/YYYY or D/M/YYYY (with or without time)
    const datePart = str.split(' ')[0];
    const parts = datePart.split('/');
    if (parts.length === 3) {
        const d = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        let y = parseInt(parts[2], 10);

        // Handle 2-digit years and deep sanity check
        if (y < 100) y += 2000;
        if (y > 3000) y = 2000 + (y % 100);

        const date = new Date(y, m - 1, d);
        return isNaN(date.getTime()) ? null : date;
    }

    // Handle YYYY-MM-DD
    if (datePart.includes('-')) {
        const partsYMD = datePart.split('-');
        if (partsYMD.length === 3) {
            let y = parseInt(partsYMD[0], 10);
            const m = parseInt(partsYMD[1], 10);
            const d = parseInt(partsYMD[2], 10);

            if (y > 3000) y = 2000 + (y % 100);

            const date = new Date(y, m - 1, d);
            return isNaN(date.getTime()) ? null : date;
        }
    }

    // Fallback to standard parser
    const d = new Date(str);
    return isNaN(d.getTime()) ? null : d;
}

// --- Rendering ---
function renderDashboard() {
    try { renderKPIs(); } catch (e) { console.error("KPI Render Error:", e); }
    try { renderIntervalKPIs(); } catch (e) { console.error("Interval KPI Render Error:", e); }
    try { renderCharts(); } catch (e) { console.error("Chart Render Error:", e); }
    try { renderLeaderboard(); } catch (e) { console.error("Leaderboard Render Error:", e); }
}

function renderKPIs() {
    const grid = document.getElementById('kpi-grid');
    if (!grid) return;
    grid.innerHTML = '';

    const totalCompletions = state.filteredData.length;
    const totalHoursNum = state.filteredData.reduce((sum, row) => sum + (parseFloat(row['CPD Hours']) || 0), 0);
    const totalHours = totalHoursNum.toFixed(1);
    const uniqueTrainings = new Set(state.filteredData.map(row => row['Training Name'])).size;

    // Unique learner detection using "Name" column
    const nameKey = state.nameKey || 'Name';
    const uniqueLearners = new Set(state.filteredData.map(row => row[nameKey]).filter(Boolean)).size;
    const avgHours = uniqueLearners > 0 ? (totalHoursNum / uniqueLearners).toFixed(1) : 0;

    const kpis = [
        { label: "Total Completions", value: totalCompletions },
        { label: "Total CPD Hours", value: totalHours },
        { label: "Unique Courses", value: uniqueTrainings },
        { label: "Unique Learners", value: uniqueLearners },
        { label: "Avg. Hours / Learner", value: avgHours }
    ];

    kpis.forEach(kpi => {
        const card = document.createElement('div');
        card.className = 'kpi-card';
        card.innerHTML = `
            <span class="kpi-value">${kpi.value}</span>
            <span class="kpi-label">${kpi.label}</span>
        `;
        grid.appendChild(card);
    });
}

function renderIntervalKPIs() {
    const dates = state.rawData.map(row => parseDateValue(row[state.dateKey])).filter(Boolean);
    if (dates.length === 0) return;

    const maxDate = new Date(Math.max(...dates));

    // Define intervals relative to maxDate
    const last7dLimit = new Date(maxDate.getTime() - 7 * 86400000);
    const last30dLimit = new Date(maxDate.getTime() - 30 * 86400000);
    const last90dLimit = new Date(maxDate.getTime() - 90 * 86400000);
    const last6mLimit = new Date(maxDate.getFullYear(), maxDate.getMonth() - 6, maxDate.getDate());
    const last1yLimit = new Date(maxDate.getFullYear() - 1, maxDate.getMonth(), maxDate.getDate());

    // Respect current filters (Office, Team, etc.) but IGNORE the Date filters
    const baseData = state.rawData.filter(row => {
        const matchOffice = state.filters.office === 'all' || row.Office === state.filters.office;
        const matchDept = state.filters.department === 'all' || row.Department === state.filters.department;
        const matchTeam = state.filters.team === 'all' || row[state.teamKey || 'Team'] === state.filters.team;
        const matchType = state.filters.type === 'all' || row['Training Type'] === state.filters.type;
        const matchPlatform = state.filters.platform === 'all' || row.Platform === state.filters.platform;
        const rowName = row[state.nameKey || 'Name'] ? String(row[state.nameKey || 'Name']).toLowerCase() : '';
        const matchName = !state.filters.name || rowName.includes(state.filters.name);

        return matchOffice && matchDept && matchTeam && matchType && matchPlatform && matchName;
    });

    const counts = {
        last7d: 0,
        last30d: 0,
        last90d: 0,
        last6m: 0,
        last1y: 0
    };

    baseData.forEach(row => {
        const d = parseDateValue(row[state.dateKey]);
        if (!d) return;
        if (d >= last7dLimit) counts.last7d++;
        if (d >= last30dLimit) counts.last30d++;
        if (d >= last90dLimit) counts.last90d++;
        if (d >= last6mLimit) counts.last6m++;
        if (d >= last1yLimit) counts.last1y++;
    });

    // Update UI
    const setVal = (id, val) => {
        const el = document.getElementById(id);
        if (el) el.textContent = val.toLocaleString();
    };

    setVal('last7d', counts.last7d);
    setVal('last30d', counts.last30d);
    setVal('last90d', counts.last90d);
    setVal('last6m', counts.last6m);
    setVal('last1y', counts.last1y);
}

// --- Utilities ---
function aggregateData(data, key, limit = 0) {
    const counts = {};
    data.forEach(row => {
        const val = row[key] || 'N/A';
        counts[val] = (counts[val] || 0) + 1;
    });

    let sorted = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    if (limit > 0) sorted = sorted.slice(0, limit);

    return {
        labels: sorted.map(i => i[0]),
        values: sorted.map(i => i[1])
    };
}

function aggregateSum(data, key, sumKey) {
    const sums = {};
    data.forEach(row => {
        const val = row[key] || 'N/A';
        const num = parseFloat(row[sumKey]) || 0;
        sums[val] = (sums[val] || 0) + num;
    });

    const sorted = Object.entries(sums).sort((a, b) => b[1] - a[1]);
    return {
        labels: sorted.map(i => i[0]),
        values: sorted.map(i => i[1].toFixed(1))
    };
}

function renderCharts() {
    // 1. Training Type Distribution
    const typeCounts = aggregateData(state.filteredData, 'Training Type');
    createChart('trainingTypeChart', 'doughnut', {
        labels: typeCounts.labels,
        datasets: [{
            data: typeCounts.values,
            backgroundColor: getVibrantColors(typeCounts.labels.length)
        }]
    }, {
        plugins: { legend: { position: 'bottom' } }
    });

    // 3. Top Job Titles
    const titleCounts = aggregateData(state.filteredData, 'Learner Job Title', 10);
    createChart('jobTitleChart', 'bar', {
        labels: titleCounts.labels,
        datasets: [{
            label: 'Completions',
            data: titleCounts.values,
            backgroundColor: getVibrantColors(titleCounts.labels.length)
        }]
    }, { indexAxis: 'y' });

    // 4. CPD Hours by Department
    const deptHours = aggregateSum(state.filteredData, 'Department', 'CPD Hours');
    createChart('departmentHoursChart', 'bar', {
        labels: deptHours.labels,
        datasets: [{
            label: 'Hours',
            data: deptHours.values,
            backgroundColor: getVibrantColors(deptHours.labels.length)
        }]
    });

    // 5. Completions by Office
    const officeCounts = aggregateData(state.filteredData, 'Office');
    createChart('officeCompletionsChart', 'doughnut', {
        labels: officeCounts.labels,
        datasets: [{
            data: officeCounts.values,
            backgroundColor: getVibrantColors(officeCounts.labels.length)
        }]
    }, {
        plugins: { legend: { position: 'bottom' } }
    });

    // 6. Completions by Team
    const teamCounts = aggregateData(state.filteredData, 'Team');
    createChart('teamsChart', 'bar', {
        labels: teamCounts.labels,
        datasets: [{
            label: 'Completions',
            data: teamCounts.values,
            backgroundColor: getVibrantColors(teamCounts.labels.length)
        }]
    });

    updateChartThemes();
}

function renderLeaderboard() {
    const container = document.getElementById('courseLeaderboard');
    if (!container) return;

    const courseStats = aggregateData(state.filteredData, 'Training Name', 20);

    if (courseStats.labels.length === 0) {
        container.innerHTML = '<p class="empty-state">No training data found for the current filters.</p>';
        return;
    }

    // Map to find platform for each training name
    const platformMap = {};
    state.filteredData.forEach(row => {
        if (!platformMap[row['Training Name']]) {
            platformMap[row['Training Name']] = row['Platform'] || 'N/A';
        }
    });

    const maxCompletions = courseStats.values[0] || 1;

    let html = `<div class="leaderboard-grid">`;

    courseStats.labels.forEach((course, index) => {
        const completions = courseStats.values[index];
        const platform = platformMap[course] || 'N/A';
        const platformClass = platform.toLowerCase().includes('cch') ? 'platform-cch' : 'platform-365';
        const progressWidth = (completions / maxCompletions) * 100;
        const isTop = index < 3;

        html += `
            <div class="leaderboard-card interactive" onclick="handleChartClick('courseLeaderboard', '${course.replace(/'/g, "\\'")}')">
                <div class="rank-pill ${isTop ? 'rank-top' : ''}">${index + 1}</div>
                <div class="course-info">
                    <div class="course-header">
                        <span class="course-name">${course}</span>
                        <span class="badge ${platformClass}">${platform}</span>
                    </div>
                    <div class="progress-bar-bg">
                        <div class="progress-bar-fill" style="width: ${progressWidth}%"></div>
                    </div>
                </div>
                <div class="completion-count">${completions}</div>
            </div>
        `;
    });

    html += '</div>';
    container.innerHTML = html;
}

// --- Export & Sharing ---

// --- Modal & Drill-Down ---
function openModal(title, data, columns) {
    const modal = document.getElementById('detailModal');
    const modalTitle = document.getElementById('modalTitle');
    const modalBody = document.getElementById('modalBody');

    if (!modal || !modalBody) return;

    state.lastDrillDownData = data;
    modalTitle.textContent = title;

    let html = `<table class="detail-table"><thead><tr>`;
    columns.forEach(col => html += `<th>${col}</th>`);
    html += `</tr></thead><tbody>`;

    data.forEach(row => {
        html += `<tr>`;
        columns.forEach(col => html += `<td>${row[col] || ''}</td>`);
        html += `</tr>`;
    });

    html += `</tbody></table>`;
    modalBody.innerHTML = html;
    modal.style.display = 'flex';
}

function closeModal() {
    document.getElementById('detailModal').style.display = 'none';
}

// --- Colors Diversified ---
function getVibrantColors(count) {
    const palette = [
        '#3b82f6', // Blue
        '#10b981', // Emerald
        '#f59e0b', // Amber
        '#a855f7', // Purple
        '#ef4444', // Red
        '#06b6d4', // Cyan
        '#ec4899', // Pink
        '#6366f1', // Indigo
        '#84cc16'  // Lime
    ];
    // Cycle or return slice
    return palette.slice(0, count);
}

function createChart(id, type, data, options = {}) {
    const canvas = document.getElementById(id);
    if (!canvas) return;

    if (state.charts[id]) {
        state.charts[id].destroy();
    }
    const ctx = canvas.getContext('2d');

    const defaultOptions = {
        responsive: true,
        maintainAspectRatio: false,
        onClick: (e, elements) => {
            if (elements.length > 0) {
                const index = elements[0].index;
                const label = data.labels[index];
                handleChartClick(id, label);
            }
        },
        plugins: {
            legend: {
                display: type !== 'bar' && type !== 'line',
                position: 'top',
                labels: { font: { family: 'Inter', size: 11 } }
            },
            datalabels: {
                color: state.theme === 'dark' ? '#94a3b8' : '#4f566b',
                anchor: 'end',
                align: type === 'bar' && options.indexAxis === 'y' ? 'right' : 'top',
                offset: type === 'bar' && options.indexAxis === 'y' ? 4 : -4,
                font: { family: 'Inter', weight: 'bold', size: 10 },
                formatter: (value) => value > 0 ? value : '', // Don't show zero labels
                display: (context) => type !== 'line' // Show for everything except line trend
            }
        },
        scales: type === 'doughnut' ? {} : {
            y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' } },
            x: { grid: { display: false } }
        }
    };

    // Style adjustments for doughnut labels
    if (type === 'doughnut') {
        defaultOptions.plugins.datalabels.anchor = 'center';
        defaultOptions.plugins.datalabels.align = 'center';
        defaultOptions.plugins.datalabels.color = '#fff'; // White labels on dark segments
        defaultOptions.plugins.datalabels.display = true;
    }

    state.charts[id] = new Chart(ctx, {
        type: type,
        data: data,
        options: Object.assign(defaultOptions, options)
    });
}

function handleChartClick(chartId, label) {
    const nameKey = state.nameKey || 'Name';
    let drillData = [];
    let title = "";
    // Transcript columns
    let cols = [nameKey, 'Training Name', 'Completion Date', 'CPD Hours', 'Office', 'Department'];

    if (chartId === 'jobTitleChart') {
        drillData = state.filteredData.filter(row => row['Learner Job Title'] === label);
        title = `Training for ${label}`;
    } else if (chartId === 'trainingTypeChart') {
        drillData = state.filteredData.filter(row => row['Training Type'] === label);
        title = `${label} Completions`;
    } else if (chartId === 'departmentHoursChart') {
        drillData = state.filteredData.filter(row => row['Department'] === label);
        title = `Training for Department: ${label}`;
    } else if (chartId === 'officeCompletionsChart') {
        drillData = state.filteredData.filter(row => row['Office'] === label);
        title = `Training for Office: ${label}`;
    } else if (chartId === 'teamsChart') {
        drillData = state.filteredData.filter(row => row[state.teamKey || 'Team'] === label);
        title = `Training for Team: ${label}`;
    } else if (chartId === 'courseLeaderboard') {
        drillData = state.filteredData.filter(row => row['Training Name'] === label);
        title = `Transcript: ${label}`;
    }

    if (drillData.length > 0) {
        openModal(title, drillData, cols);
    }
}

function updatePreparedDate() {
    const now = new Date();
    document.getElementById('preparedDate').textContent = now.toLocaleDateString('en-AU');
}

function showError(msg) {
    const banner = document.getElementById('errorBanner');
    banner.textContent = msg;
    banner.style.display = 'block';
}

function hideError() {
    document.getElementById('errorBanner').style.display = 'none';
}
