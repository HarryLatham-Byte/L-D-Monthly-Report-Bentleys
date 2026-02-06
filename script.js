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

    // Select Filters
    const filterIds = ['officeFilter', 'departmentFilter', 'trainingTypeFilter', 'platformFilter'];
    filterIds.forEach(id => {
        document.getElementById(id).addEventListener('change', (e) => {
            const filterKey = id.replace('Filter', '').toLowerCase();
            state.filters[id === 'trainingTypeFilter' ? 'type' : filterKey] = e.target.value;
            applyFilters();
        });
    });

    // Name Search
    const nameSearch = document.getElementById('nameSearch');
    nameSearch.addEventListener('input', (e) => {
        state.filters.name = e.target.value.trim().toLowerCase();
        applyFilters();
    });
    // Help the dropdown appear on focus
    nameSearch.addEventListener('focus', () => {
        nameSearch.setAttribute('placeholder', 'Type to search...');
    });

    // Date Filters
    ['startDate', 'endDate'].forEach(id => {
        document.getElementById(id).addEventListener('change', (e) => {
            state.filters[id] = e.target.value;
            applyFilters();
        });
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

function applyFilters() {
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

    renderDashboard();
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

function renderCharts() {
    // 1. Monthly Trend (Now AT THE TOP)
    const monthlyTrend = aggregateTrend(state.filteredData);
    createChart('monthlyTrendChart', 'line', {
        labels: monthlyTrend.labels,
        datasets: [{
            label: 'Completions',
            data: monthlyTrend.values,
            borderColor: COLORS.primary,
            backgroundColor: 'rgba(247, 148, 29, 0.1)',
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
        datasets: [{ data: typeCounts.values, backgroundColor: COLORS.backgrounds }]
    });

    // 3. Top Job Titles
    const titleCounts = aggregateData(state.filteredData, 'Learner Job Title', 10);
    createChart('jobTitleChart', 'bar', {
        labels: titleCounts.labels,
        datasets: [{ label: 'Completions', data: titleCounts.values, backgroundColor: COLORS.secondary }]
    }, { indexAxis: 'y' });

    // 4. CPD Hours by Department
    const deptHours = aggregateSum(state.filteredData, 'Department', 'CPD Hours');
    createChart('departmentHoursChart', 'bar', {
        labels: deptHours.labels,
        datasets: [{ label: 'Hours', data: deptHours.values, backgroundColor: COLORS.backgrounds }]
    });

    // 5. Completions by Office (Now AT THE BOTTOM)
    const officeCounts = aggregateData(state.filteredData, 'Office');
    createChart('officeCompletionsChart', 'bar', {
        labels: officeCounts.labels,
        datasets: [{ label: 'Completions', data: officeCounts.values, backgroundColor: COLORS.primary }]
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

    let html = `
        <table class="leaderboard-table">
            <thead>
                <tr>
                    <th class="rank-cell">Rank</th>
                    <th class="course-cell">Course Name</th>
                    <th class="platform-cell">Platform</th>
                    <th class="count-cell">Completions</th>
                </tr>
            </thead>
            <tbody>
    `;

    courseStats.labels.forEach((course, index) => {
        const platform = platformMap[course] || 'N/A';
        const platformClass = platform.toLowerCase().includes('cch') ? 'platform-cch' : 'platform-365';

        html += `
            <tr>
                <td class="rank-cell">#${index + 1}</td>
                <td class="course-cell">${course}</td>
                <td class="platform-cell"><span class="badge ${platformClass}">${platform}</span></td>
                <td class="count-cell">${courseStats.values[index]}</td>
            </tr>
        `;
    });

    html += '</tbody></table>';
    container.innerHTML = html;
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
        try {
            let dateVal = row['Completion Date'];
            if (!dateVal) return;

            let monthYear = '';

            // Robust check for Date object
            if (dateVal instanceof Date || (dateVal && typeof dateVal.getMonth === 'function')) {
                const year = dateVal.getFullYear();
                const month = (dateVal.getMonth() + 1).toString().padStart(2, '0');
                monthYear = `${year}-${month}`;
            } else {
                // Defensive string conversion
                const dateStr = String(dateVal);
                if (typeof dateStr.split === 'function') {
                    const firstPart = dateStr.split(' ')[0];
                    if (firstPart) {
                        const parts = firstPart.split('/');
                        if (parts.length >= 3) {
                            monthYear = `${parts[2]}-${parts[1]}`; // YYYY-MM
                        }
                    }
                }
            }

            if (monthYear) {
                trend[monthYear] = (trend[monthYear] || 0) + 1;
            }
        } catch (e) {
            console.warn("Skipping row due to date parse error:", e);
        }
    });

    const sortedMonths = Object.keys(trend).sort();
    return {
        labels: sortedMonths,
        values: sortedMonths.map(m => trend[m])
    };
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
        plugins: {
            legend: {
                display: type !== 'bar',
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
        options: JSON.parse(JSON.stringify(Object.assign(defaultOptions, options)))
    });
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
