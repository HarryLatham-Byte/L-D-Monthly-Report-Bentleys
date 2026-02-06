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
        name: '',
        startDate: '',
        endDate: ''
    },
    charts: {}
};

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
    loadLocalData();
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
    const filterIds = ['officeFilter', 'departmentFilter', 'trainingTypeFilter', 'platformFilter'];
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

    // Export & Action Buttons
    document.getElementById('exportPDF')?.addEventListener('click', () => window.print());
    document.getElementById('copyShareLink')?.addEventListener('click', copyShareLink);

    // Modal Close
    document.getElementById('closeModal')?.addEventListener('click', closeModal);
    window.addEventListener('click', (e) => {
        if (e.target.classList.contains('modal-overlay')) closeModal();
    });
}

// --- Data Loading ---
async function loadLocalData() {
    const fileName = 'report_data.csv';
    try {
        const response = await fetch(fileName, { cache: 'no-cache' });
        if (!response.ok) {
            throw new Error(`Failed to load '${fileName}'. Status: ${response.status}`);
        }

        const csvText = await response.text();
        const workbook = XLSX.read(csvText, { type: 'string' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);

        state.rawData = json;
        populateFilters();
        applyFilters();
        updatePreparedDate();
        hideError();
    } catch (err) {
        console.error("Data load error:", err);
        let message = err.message;
        if (window.location.protocol === 'file:') {
            message = "Browser security blocks local file reading. Please use a local web server (e.g., Live Server).";
        }
        showError(message);
    }
}

// --- Filter Management ---
function populateFilters() {
    if (state.rawData.length === 0) return;

    // Detect column name for "Name" (handle BOM or case variations)
    const sampleRow = state.rawData[0];
    const nameKey = Object.keys(sampleRow).find(k => k.toLowerCase().includes('name') && !k.toLowerCase().includes('manager')) || 'Name';
    state.nameKey = nameKey; // Store for later use

    const offices = [...new Set(state.rawData.map(row => row.Office))].filter(Boolean).sort();
    const departments = [...new Set(state.rawData.map(row => row.Department))].filter(Boolean).sort();
    const types = [...new Set(state.rawData.map(row => row['Training Type']))].filter(Boolean).sort();
    const platforms = [...new Set(state.rawData.map(row => row.Platform))].filter(Boolean).sort();
    const names = [...new Set(state.rawData.map(row => row[nameKey]))].filter(Boolean).sort();

    updateSelectOptions('officeFilter', offices, 'All Offices');
    updateSelectOptions('departmentFilter', departments, 'All Departments');
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
        .map(row => parseDateValue(row['Completion Date']))
        .filter(d => d instanceof Date && !isNaN(d.getTime()));

    if (dates.length === 0) return;

    const minDate = new Date(Math.min(...dates));
    const maxDate = new Date(Math.max(...dates));

    const toISODate = (d) => d.toISOString().split('T')[0];

    const startInput = document.getElementById('startDate');
    const endInput = document.getElementById('endDate');

    if (startInput && endInput) {
        startInput.min = toISODate(minDate);
        startInput.max = toISODate(maxDate);
        endInput.min = toISODate(minDate);
        endInput.max = toISODate(maxDate);

        // Optional: Pre-fill with the full range
        // startInput.value = toISODate(minDate);
        // endInput.value = toISODate(maxDate);
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
        const matchType = state.filters.type === 'all' || row['Training Type'] === state.filters.type;
        const matchPlatform = state.filters.platform === 'all' || row.Platform === state.filters.platform;

        // Name search (case insensitive)
        const rowName = row[nameKey] ? String(row[nameKey]).toLowerCase() : '';
        const matchName = !state.filters.name || rowName.includes(state.filters.name);

        // Date range filtering
        let matchDate = true;
        if (state.filters.startDate || state.filters.endDate) {
            const rowDate = parseDateValue(row['Completion Date']);
            if (rowDate) {
                if (state.filters.startDate) {
                    const start = new Date(state.filters.startDate);
                    start.setHours(0, 0, 0, 0);
                    if (rowDate < start) matchDate = false;
                }
                if (state.filters.endDate) {
                    const end = new Date(state.filters.endDate);
                    end.setHours(23, 59, 59, 999);
                    if (rowDate > end) matchDate = false;
                }
            } else {
                matchDate = false;
            }
        }

        return matchOffice && matchDept && matchType && matchPlatform && matchName && matchDate;
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

    // Handle "DD/MM/YYYY HH:mm"
    const str = String(val).split(' ')[0];
    const parts = str.split('/');
    if (parts.length === 3) {
        // DD/MM/YYYY -> YYYY, MM-1, DD
        return new Date(parts[2], parts[1] - 1, parts[0]);
    }

    // Fallback to standard parser
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
}

// --- Rendering ---
function renderDashboard() {
    try { renderKPIs(); } catch (e) { console.error("KPI Render Error:", e); }
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

function aggregateTrend(data) {
    const trend = {};
    data.forEach(row => {
        const dateVal = row['Completion Date'];
        if (!dateVal) return;
        const date = parseDateValue(dateVal);
        if (date) {
            const year = date.getFullYear();
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const key = `${year}-${month}`;
            trend[key] = (trend[key] || 0) + 1;
        }
    });

    const sortedLabels = Object.keys(trend).sort();
    // Prettier labels for display
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const displayLabels = sortedLabels.map(label => {
        const [y, m] = label.split('-');
        return `${months[parseInt(m) - 1]} ${y}`;
    });

    return {
        keys: sortedLabels, // Used for internal matching
        labels: displayLabels,
        values: sortedLabels.map(l => trend[l])
    };
}

function renderCharts() {
    // 1. Monthly Trend
    const monthlyTrend = aggregateTrend(state.filteredData);
    createChart('monthlyTrendChart', 'line', {
        labels: monthlyTrend.labels,
        keys: monthlyTrend.keys, // Custom property for click handling
        datasets: [{
            label: 'Completions',
            data: monthlyTrend.values,
            borderColor: '#3b82f6',
            backgroundColor: 'rgba(59, 130, 246, 0.1)',
            fill: true,
            tension: 0.4
        }]
    }, {
        plugins: { legend: { display: false } }
    });

    // 2. Training Type Distribution
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
    createChart('officeCompletionsChart', 'bar', {
        labels: officeCounts.labels,
        datasets: [{
            label: 'Completions',
            data: officeCounts.values,
            backgroundColor: getVibrantColors(officeCounts.labels.length)
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
            <div class="leaderboard-card">
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
function copyShareLink() {
    const url = window.location.href;
    navigator.clipboard.writeText(url).then(() => {
        const btn = document.getElementById('copyShareLink');
        const originalText = btn.innerHTML;
        btn.innerHTML = '<span>Copied!</span>';
        setTimeout(() => btn.innerHTML = originalText, 2000);
    });
}

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
            }
        },
        scales: type === 'doughnut' ? {} : {
            y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' } },
            x: { grid: { display: false } }
        }
    };

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

    if (chartId === 'monthlyTrendChart') {
        const chart = state.charts[chartId];
        const index = chart.data.labels.indexOf(label);
        const key = chart.data.keys[index]; // Use the YYYY-MM key

        drillData = state.filteredData.filter(row => {
            const date = parseDateValue(row['Completion Date']);
            if (!date) return false;
            const monthYear = `${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}`;
            return monthYear === key;
        });
        title = `Completions in ${label}`;
    } else if (chartId === 'jobTitleChart') {
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
