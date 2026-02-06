/**
 * L&D Monthly Summary Report Dashboard
 * Core Logic & State Management
 */

// --- Default Data (December 2025) ---
const DEFAULT_DATA = {
    summary: [
        { label: "Compliance Training Completion Rate", value: "96%" },
        { label: "Attendance Rate", value: "35%" },
        { label: "Overall Satisfaction of Training Provided", value: "90%" },
        { label: "Overall Engagement", value: "95%" },
        { label: "Knowledge Before Session", value: "6.75 / 10" },
        { label: "Knowledge After Session", value: "8.75 / 10" },
        { label: "Confidence in Applying Learnt Knowledge", value: "84%" },
        { label: "Satisfaction with Learn365 Platform", value: "81%" }
    ],
    preferences: {
        labels: ["In-Person", "Virtual", "Watch Recording"],
        data: [25, 50, 0]
    },
    insights: [
        "Positive feedback: engaging videos, examples, trainer insights",
        "Improvement opportunity: more PA/TA training, handling different client types",
        "Learning objects summary: 345 Active Users, 12 New Modules"
    ],
    snapshot: {
        uptodate: "99.5%",
        subtext: "1,240 Total Learning Objects"
    },
    officeCompletions: {
        labels: ["Brisbane", "Manila"],
        enrolments: [145, 82],
        completions: [120, 75]
    },
    teamCompletions: {
        labels: ["Affinity Audit", "Team 1", "Team 2", "Team 4", "Team 5", "Team 7", "Other"],
        enrolments: [24, 18, 12, 15, 20, 10, 30],
        completions: [20, 15, 10, 12, 18, 8, 25]
    },
    positionCompletions: {
        labels: ["Grads", "Intermediate", "Senior", "Supervisor", "Manager", "AD & Above", "PA/TA", "Other"],
        enrolments: [40, 35, 30, 20, 15, 10, 25, 12],
        completions: [38, 32, 28, 18, 14, 9, 22, 10]
    },
    teamHours: {
        labels: ["Affinity Audit", "Team 1", "Team 2", "Team 4", "Team 5", "Team 7", "Other"],
        data: [15, 12, 8, 10, 14, 11, 17] // Total 87
    },
    positionHours: {
        labels: ["Grads", "Intermediate", "Senior", "Supervisor", "Manager", "AD & Above", "PA/TA", "Other"],
        data: [25, 20, 15, 10, 6, 4, 4, 3]
    },
    cchKpis: [
        { label: "Time to Completion (Median) - Month", value: "7.74 days" },
        { label: "Time to Completion (Median) - YTD", value: "14.04 days" },
        { label: "CPD Hours Earnt - Month", value: "52.5" },
        { label: "CPD Hours Earnt - YTD", value: "719.25" },
        { label: "Enrolments - Month", value: "73" },
        { label: "Enrolments - YTD", value: "1162" },
        { label: "Completions - Month", value: "51" },
        { label: "Completions - YTD", value: "663" },
        { label: "Completion % - Month", value: "69.9%" },
        { label: "Completion % - YTD", value: "57%" },
        { label: "Cost per CPD Hour - Month", value: "$54.69" },
        { label: "Cost per CPD Hour - YTD", value: "$48.13" }
    ],
    cchTrends: {
        months: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        timeToCompletion: [10, 12, 15, 11, 14, 13, 12, 11, 10, 9, 8, 7.74],
        cpdHours: [45, 60, 55, 70, 65, 80, 75, 50, 48, 52, 65, 52.5],
        enrolments: [80, 95, 110, 105, 90, 120, 115, 100, 98, 85, 92, 73],
        completions: [50, 70, 65, 80, 60, 90, 85, 75, 60, 55, 58, 51]
    }
};

// --- Empty / Zero Data State ---
const ZERO_DATA = {
    summary: [
        { label: "Compliance Training Completion Rate", value: "0%" },
        { label: "Attendance Rate", value: "0%" },
        { label: "Overall Satisfaction of Training Provided", value: "0%" },
        { label: "Overall Engagement", value: "0%" },
        { label: "Knowledge Before Session", value: "0 / 10" },
        { label: "Knowledge After Session", value: "0 / 10" },
        { label: "Confidence in Applying Learnt Knowledge", value: "0%" },
        { label: "Satisfaction with Learn365 Platform", value: "0%" }
    ],
    preferences: { labels: [], data: [] },
    insights: [],
    snapshot: { uptodate: "0%", subtext: "0 Total Learning Objects" },
    officeCompletions: { labels: [], enrolments: [], completions: [] },
    teamCompletions: { labels: [], enrolments: [], completions: [] },
    positionCompletions: { labels: [], enrolments: [], completions: [] },
    teamHours: { labels: [], data: [] },
    positionHours: { labels: [], data: [] },
    cchKpis: [
        { label: "Time to Completion (Median) - Month", value: "0 days" },
        { label: "Time to Completion (Median) - YTD", value: "0 days" },
        { label: "CPD Hours Earnt - Month", value: "0" },
        { label: "CPD Hours Earnt - YTD", value: "0" },
        { label: "Enrolments - Month", value: "0" },
        { label: "Enrolments - YTD", value: "0" },
        { label: "Completions - Month", value: "0" },
        { label: "Completions - YTD", value: "0" },
        { label: "Completion % - Month", value: "0%" },
        { label: "Completion % - YTD", value: "0%" },
        { label: "Cost per CPD Hour - Month", value: "$0.00" },
        { label: "Cost per CPD Hour - YTD", value: "$0.00" }
    ],
    cchTrends: {
        months: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        timeToCompletion: Array(12).fill(0),
        cpdHours: Array(12).fill(0),
        enrolments: Array(12).fill(0),
        completions: Array(12).fill(0)
    }
};

// --- State Management ---
const state = {
    theme: 'light',
    data: JSON.parse(JSON.stringify(ZERO_DATA)), // Start with zeroed-out data
    charts: {}
};

// --- Theme Handling ---
function toggleTheme() {
    state.theme = state.theme === 'light' ? 'dark' : 'light';
    document.body.classList.toggle('dark-mode', state.theme === 'dark');

    document.getElementById('moonIcon').style.display = state.theme === 'light' ? 'block' : 'none';
    document.getElementById('sunIcon').style.display = state.theme === 'dark' ? 'block' : 'none';

    // Update charts for theme
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
        'rgba(247, 148, 29, 0.8)', // Orange
        'rgba(255, 173, 77, 0.8)', // Light Orange
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
    renderDashboard();
    setupEventListeners();
    loadLocalExcelData(); // Automatically load the data on startup
});

function initUI() {
    // UI initializations
}

function setupEventListeners() {
    document.getElementById('themeToggle').addEventListener('click', toggleTheme);
}

// --- Automated Local Data Fetching ---
async function loadLocalExcelData() {
    const fileName = 'report_data.xlsx';
    try {
        const response = await fetch(fileName, { cache: 'no-cache' });
        if (!response.ok) {
            if (response.status === 404) {
                throw new Error(`File '${fileName}' not found. Please ensure the file is named exactly '${fileName}' (case-sensitive) in your repository.`);
            } else {
                throw new Error(`Failed to load '${fileName}'. Status: ${response.status}`);
            }
        }

        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        // Assume first sheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);

        if (validateData(json)) {
            processExcelData(json);
            updatePreparedDate();
            hideError();
        } else {
            showError("Invalid Excel content. Ensure required columns (metric_category, metric_name, metric_value) are present.");
        }
    } catch (err) {
        console.error("Auto-load error:", err);
        let message = err.message;
        if (window.location.protocol === 'file:') {
            message = "Browser security blocks local file reading. Please view dashboard via a local web server (e.g., Live Server).";
        }
        showError(message);
    }
}

// --- Render Logic ---
function renderDashboard() {
    renderKpis();
    renderCharts();
    renderCchKpis();
}

function renderKpis() {
    const grid = document.getElementById('kpi-grid');
    grid.innerHTML = '';

    state.data.summary.forEach(kpi => {
        const card = document.createElement('div');
        card.className = 'kpi-card';
        card.innerHTML = `
            <span class="kpi-value">${kpi.value}</span>
            <span class="kpi-label">${kpi.label}</span>
        `;
        grid.appendChild(card);
    });
}

function renderCchKpis() {
    const grid = document.getElementById('cch-kpi-grid');
    grid.innerHTML = '';

    state.data.cchKpis.forEach(kpi => {
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
    // 1. Office Completions vs Enrolments
    createChart('officeCompletionsChart', 'bar', {
        labels: state.data.officeCompletions.labels,
        datasets: [
            { label: 'Enrolments', data: state.data.officeCompletions.enrolments, backgroundColor: COLORS.primary },
            { label: 'Completions', data: state.data.officeCompletions.completions, backgroundColor: COLORS.secondary }
        ]
    });

    // 2. Team Completions
    createChart('teamCompletionsChart', 'bar', {
        labels: state.data.teamCompletions.labels,
        datasets: [
            { label: 'Enrolments', data: state.data.teamCompletions.enrolments, backgroundColor: COLORS.primary },
            { label: 'Completions', data: state.data.teamCompletions.completions, backgroundColor: COLORS.secondary }
        ]
    }, { indexAxis: 'y' });

    // 3. Position Completions
    createChart('positionCompletionsChart', 'bar', {
        labels: state.data.positionCompletions.labels,
        datasets: [
            { label: 'Enrolments', data: state.data.positionCompletions.enrolments, backgroundColor: COLORS.primary },
            { label: 'Completions', data: state.data.positionCompletions.completions, backgroundColor: COLORS.secondary }
        ]
    }, { indexAxis: 'y' });

    // 4. Team Hours
    createChart('teamHoursChart', 'bar', {
        labels: state.data.teamHours.labels,
        datasets: [{
            label: 'Hours',
            data: state.data.teamHours.data,
            backgroundColor: COLORS.backgrounds
        }]
    });

    // 5. Position Hours
    createChart('positionHoursChart', 'bar', {
        labels: state.data.positionHours.labels,
        datasets: [{
            label: 'Hours',
            data: state.data.positionHours.data,
            backgroundColor: COLORS.backgrounds
        }]
    });

    // 6. CCH Trends
    createChart('cchTrendsChart', 'line', {
        labels: state.data.cchTrends.months,
        datasets: [
            { label: 'Enrolments', data: state.data.cchTrends.enrolments, borderColor: COLORS.primary, tension: 0.3 },
            { label: 'Completions', data: state.data.cchTrends.completions, borderColor: COLORS.secondary, tension: 0.3 },
            { label: 'CPD Hours', data: state.data.cchTrends.cpdHours, borderColor: COLORS.accent, tension: 0.3, yAxisID: 'y1' }
        ]
    }, {
        scales: {
            y: { beginAtZero: true, title: { display: true, text: 'Counts' } },
            y1: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: 'Hours' } }
        }
    });

    updateChartThemes();
}

function createChart(id, type, data, options = {}) {
    const canvas = document.getElementById(id);
    if (!canvas) return; // For removed charts

    if (state.charts[id]) {
        state.charts[id].destroy();
    }
    const ctx = canvas.getContext('2d');
    const defaultOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                display: true,
                position: 'top',
                labels: { font: { family: 'Inter', size: 12 } }
            },
            tooltip: {
                backgroundColor: 'rgba(26, 31, 54, 0.9)',
                padding: 12,
                titleFont: { weight: 'bold' }
            }
        },
        animation: { duration: 1000, easing: 'easeOutQuart' }
    };

    state.charts[id] = new Chart(ctx, {
        type: type,
        data: data,
        options: Object.assign(defaultOptions, options)
    });
}

function updatePreparedDate() {
    const now = new Date();
    const formatted = `${now.getDate().toString().padStart(2, '0')}/${(now.getMonth() + 1).toString().padStart(2, '0')}/${now.getFullYear()}`;
    document.getElementById('preparedDate').textContent = formatted;
}

function validateData(data) {
    if (!data || data.length === 0) return false;
    const required = ['metric_category', 'metric_name', 'metric_value'];
    const headers = Object.keys(data[0]);
    return required.every(r => headers.includes(r));
}

function processExcelData(rows) {
    // Start with fresh zero data
    const newData = JSON.parse(JSON.stringify(ZERO_DATA));
    newData.summary = []; // Clear summary to rebuild from Excel
    newData.cchKpis = []; // Clear CCH KPIs to rebuild from Excel

    rows.forEach(row => {
        const cat = row.metric_category;
        const name = row.metric_name;
        const value = row.metric_value;
        const group = row.group_name;

        switch (cat) {
            case 'KPI':
                newData.summary.push({ label: name, value: value });
                break;
            case 'Preference':
                // Removed section but data might still be in CSV
                break;
            case 'Office':
                if (!newData.officeCompletions.labels.includes(group)) {
                    newData.officeCompletions.labels.push(group);
                    newData.officeCompletions.enrolments.push(0);
                    newData.officeCompletions.completions.push(0);
                }
                const oIdx = newData.officeCompletions.labels.indexOf(group);
                if (name === 'Enrolment') newData.officeCompletions.enrolments[oIdx] = value;
                else newData.officeCompletions.completions[oIdx] = value;
                break;
            case 'TeamCompletion':
                if (!newData.teamCompletions.labels.includes(group)) {
                    newData.teamCompletions.labels.push(group);
                    newData.teamCompletions.enrolments.push(0);
                    newData.teamCompletions.completions.push(0);
                }
                const tIdx = newData.teamCompletions.labels.indexOf(group);
                if (name === 'Enrolment') newData.teamCompletions.enrolments[tIdx] = value;
                else newData.teamCompletions.completions[tIdx] = value;
                break;
            case 'PositionCompletion':
                if (!newData.positionCompletions.labels.includes(group)) {
                    newData.positionCompletions.labels.push(group);
                    newData.positionCompletions.enrolments.push(0);
                    newData.positionCompletions.completions.push(0);
                }
                const pIdx = newData.positionCompletions.labels.indexOf(group);
                if (name === 'Enrolment') newData.positionCompletions.enrolments[pIdx] = value;
                else newData.positionCompletions.completions[pIdx] = value;
                break;
            case 'TeamHours':
                newData.teamHours.labels.push(group);
                newData.teamHours.data.push(value);
                break;
            case 'PositionHours':
                newData.positionHours.labels.push(group);
                newData.positionHours.data.push(value);
                break;
            case 'CCH_KPI':
                newData.cchKpis.push({ label: name, value: value });
                break;
            case 'CCH_Trend':
                if (!newData.cchTrends.months.includes(group)) newData.cchTrends.months.push(group);
                const mIdx = newData.cchTrends.months.indexOf(group);
                if (name === 'TimeToCompletion') newData.cchTrends.timeToCompletion[mIdx] = value;
                else if (name === 'CPDHours') newData.cchTrends.cpdHours[mIdx] = value;
                else if (name === 'Enrolment') newData.cchTrends.enrolments[mIdx] = value;
                else if (name === 'Completion') newData.cchTrends.completions[mIdx] = value;
                break;
        }
    });

    state.data = newData;
    renderDashboard();
}

function showError(msg) {
    const banner = document.getElementById('errorBanner');
    banner.textContent = msg;
    banner.style.display = 'block';
    setTimeout(hideError, 5000);
}

function hideError() {
    document.getElementById('errorBanner').style.display = 'none';
}
