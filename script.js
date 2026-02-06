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
        platform: 'all'
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

    // Filter Listeners
    const filterIds = ['officeFilter', 'departmentFilter', 'trainingTypeFilter', 'platformFilter'];
    filterIds.forEach(id => {
        document.getElementById(id).addEventListener('change', (e) => {
            const filterKey = id.replace('Filter', '').toLowerCase();
            state.filters[id === 'trainingTypeFilter' ? 'type' : filterKey] = e.target.value;
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
    const offices = [...new Set(state.rawData.map(row => row.Office))].filter(Boolean).sort();
    const departments = [...new Set(state.rawData.map(row => row.Department))].filter(Boolean).sort();
    const types = [...new Set(state.rawData.map(row => row['Training Type']))].filter(Boolean).sort();
    const platforms = [...new Set(state.rawData.map(row => row.Platform))].filter(Boolean).sort();

    updateSelectOptions('officeFilter', offices, 'All Offices');
    updateSelectOptions('departmentFilter', departments, 'All Departments');
    updateSelectOptions('trainingTypeFilter', types, 'All Types');
    updateSelectOptions('platformFilter', platforms, 'All Platforms');
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
    state.filteredData = state.rawData.filter(row => {
        const matchOffice = state.filters.office === 'all' || row.Office === state.filters.office;
        const matchDept = state.filters.department === 'all' || row.Department === state.filters.department;
        const matchType = state.filters.type === 'all' || row['Training Type'] === state.filters.type;
        const matchPlatform = state.filters.platform === 'all' || row.Platform === state.filters.platform;
        return matchOffice && matchDept && matchType && matchPlatform;
    });

    renderDashboard();
}

// --- Rendering ---
function renderDashboard() {
    renderKPIs();
    renderCharts();
}

function renderKPIs() {
    const grid = document.getElementById('kpi-grid');
    grid.innerHTML = '';

    const totalCompletions = state.filteredData.length;
    const totalHours = state.filteredData.reduce((sum, row) => sum + (parseFloat(row['CPD Hours']) || 0), 0).toFixed(1);
    const uniqueTrainings = new Set(state.filteredData.map(row => row['Training Name'])).size;
    const uniqueLearners = new Set(state.filteredData.map(row => row['Line Manager Name'] + row['Learner Job Title'])).size; // Approximate

    const kpis = [
        { label: "Total Completions", value: totalCompletions },
        { label: "Total CPD Hours", value: totalHours },
        { label: "Unique Courses", value: uniqueTrainings },
        { label: "Active Learners", value: uniqueLearners }
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
    // 1. Completions by Office
    const officeCounts = aggregateData(state.filteredData, 'Office');
    createChart('officeCompletionsChart', 'bar', {
        labels: officeCounts.labels,
        datasets: [{ label: 'Completions', data: officeCounts.values, backgroundColor: COLORS.primary }]
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

    // 5. Monthly Trend
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
    });

    updateChartThemes();
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
        if (!row['Completion Date']) return;
        // Parse "19/12/2025 2:45"
        const parts = row['Completion Date'].split(' ')[0].split('/');
        if (parts.length < 3) return;
        const monthYear = `${parts[2]}-${parts[1]}`; // YYYY-MM
        trend[monthYear] = (trend[monthYear] || 0) + 1;
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
            y: { beginAtZero: true, grid: { color: '#e3e8ee' } },
            x: { grid: { display: false } }
        }
    };

    state.charts[id] = new Chart(ctx, {
        type: type,
        data: data,
        options: Object.assign(defaultOptions, options)
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
