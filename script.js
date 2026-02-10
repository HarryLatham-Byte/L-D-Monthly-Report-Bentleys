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

    // Export & Action Buttons
    document.getElementById('exportPDF')?.addEventListener('click', () => window.print());
    document.getElementById('copyShareLink')?.addEventListener('click', copyShareLink);
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

        // Default to last 12 months
        const defaultStart = new Date(maxDate.getFullYear(), maxDate.getMonth() - 11, 1);
        startInput.value = toLocalISODate(defaultStart < minDate ? minDate : defaultStart);
        endInput.value = toLocalISODate(maxDate);

        // Sync state
        state.filters.startDate = startInput.value;
        state.filters.endDate = endInput.value;
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

    const str = String(val).trim();
    if (!str) return null;

    // Handle DD/MM/YYYY or D/M/YYYY (with or without time)
    const datePart = str.split(' ')[0];
    const parts = datePart.split('/');
    if (parts.length === 3) {
        const d = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        const y = parseInt(parts[2], 10);
        // Ensure 4-digit year
        const fullYear = y < 100 ? 2000 + y : y;
        const date = new Date(fullYear, m - 1, d);
        return isNaN(date.getTime()) ? null : date;
    }

    // Handle YYYY-MM-DD
    if (datePart.includes('-')) {
        const partsYMD = datePart.split('-');
        if (partsYMD.length === 3) {
            const date = new Date(partsYMD[0], partsYMD[1] - 1, partsYMD[2]);
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
    const dates = state.rawData.map(row => parseDateValue(row[state.dateKey])).filter(Boolean);
    if (dates.length === 0) return { labels: [], datasets: [] };

    const maxDate = new Date(Math.max(...dates));
    const currentYearEnd = new Date(maxDate.getFullYear(), maxDate.getMonth(), maxDate.getDate());
    const currentYearStart = new Date(maxDate.getFullYear(), maxDate.getMonth() - 11, 1);
    const prevYearStart = new Date(maxDate.getFullYear(), maxDate.getMonth() - 23, 1);

    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const currentCounts = new Array(12).fill(0);
    const prevCounts = new Array(12).fill(0);

    state.rawData.forEach(row => {
        const d = parseDateValue(row[state.dateKey]);
        if (!d) return;

        const mIdx = d.getMonth(); // 0-11

        if (d >= currentYearStart && d <= currentYearEnd) {
            currentCounts[mIdx]++;
        } else if (d >= prevYearStart && d < currentYearStart) {
            prevCounts[mIdx]++;
        }
    });

    // Create dynamic labels for legend (e.g., "Feb 24 - Jan 25")
    const fmt = (d) => `${months[d.getMonth()]} ${d.getFullYear().toString().slice(-2)}`;
    const currentLabel = `${fmt(currentYearStart)} - ${fmt(currentYearEnd)}`;
    const prevLabel = `${fmt(prevYearStart)} - ${fmt(new Date(currentYearStart.getTime() - 86400000))}`;

    return {
        labels: months,
        datasets: [
            {
                label: currentLabel,
                data: currentCounts,
                borderColor: '#f7941d', // Orange
                backgroundColor: 'transparent',
                borderWidth: 3,
                tension: 0.3,
                pointRadius: 4
            },
            {
                label: prevLabel,
                data: prevCounts,
                borderColor: '#3b82f6', // Blue
                backgroundColor: 'transparent',
                borderWidth: 2,
                tension: 0.3,
                pointRadius: 3,
                borderDash: [5, 5]
            }
        ]
    };
}

function renderCharts() {
    // 1. Monthly Trend
    const monthlyTrend = aggregateTrend(state.filteredData);
    createChart('monthlyTrendChart', 'line', {
        labels: monthlyTrend.labels,
        datasets: monthlyTrend.datasets
    }, {
        plugins: {
            legend: {
                display: true,
                position: 'top',
                labels: { usePointStyle: true, boxWidth: 6, font: { family: 'Inter', size: 10 } }
            }
        },
        interaction: {
            mode: 'index',
            intersect: false
        }
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

    if (chartId === 'monthlyTrendChart') {
        const chart = state.charts[chartId];
        const datasetIndex = elements[0].datasetIndex;
        const monthIndex = elements[0].index;
        const periodLabel = chart.data.datasets[datasetIndex].label; // e.g. "Feb 24 - Jan 25"
        const monthName = chart.data.labels[monthIndex];

        // Extract year from periodLabel based on monthIndex
        // This is a bit tricky since labels cover 2 years. 
        // We'll filter all records in state.rawData that match monthName and are within the 24 month window
        drillData = state.rawData.filter(row => {
            const d = parseDateValue(row[state.dateKey]);
            if (!d) return false;
            if (d.getMonth() !== monthIndex) return false;

            // Is it the orange line or blue line?
            const dates = state.rawData.map(r => parseDateValue(r[state.dateKey])).filter(Boolean);
            const maxDate = new Date(Math.max(...dates));
            const currentYearStart = new Date(maxDate.getFullYear(), maxDate.getMonth() - 11, 1);
            const prevYearStart = new Date(maxDate.getFullYear(), maxDate.getMonth() - 23, 1);

            if (datasetIndex === 0) { // Orange (Current)
                return d >= currentYearStart;
            } else { // Blue (Prev)
                return d >= prevYearStart && d < currentYearStart;
            }
        });
        title = `Completions in ${monthName} (${periodLabel})`;
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
