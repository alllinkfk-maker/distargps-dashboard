/* ==========================================================================
   JavaScript Logic - GPS PM Dashboard (Robust Version)
   ========================================================================== */

// Global State
let rawData = [];
let filteredData = [];
let currentSort = { column: 'id', ascending: true };
let pagination = { currentPage: 1, pageSize: 15 };

// Chart Instances
let charts = {
    itTicket: null,
    openSv: null,
    deviceStatus: null,
    topCustomers: null
};

// Document Ready
document.addEventListener('DOMContentLoaded', () => {
    initApp();
});

// Initialize Application
async function initApp() {
    setupEventListeners();
    
    // Try to load pre-loaded PM data
    try {
        const response = await fetch('./pm_data.json');
        if (response.ok) {
            const data = await response.json();
            loadDataset(data, 'ใช้ข้อมูลเริ่มต้น (All PM DistarGPS.xlsx)');
        } else {
            showBadgeStatus('รออัปโหลดข้อมูล (ไม่พบไฟล์ pm_data.json)', 'orange');
        }
    } catch (e) {
        console.warn('Could not load pm_data.json, waiting for user upload', e);
        showBadgeStatus('รออัปโหลดข้อมูล...', 'orange');
    }
}

// Setup Event Listeners
function setupEventListeners() {
    const fileInput = document.getElementById('excel-upload');
    const dropZone = document.getElementById('drop-zone');
    
    // File upload change
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleExcelFile(file);
    });

    // Drag and drop event listeners
    dropZone.addEventListener('click', () => fileInput.click());
    
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });
    
    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });
    
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) handleExcelFile(file);
    });

    // Search input listener
    document.getElementById('search-input').addEventListener('input', () => {
        pagination.currentPage = 1;
        applyFilters();
    });

    // Dropdown filters listeners
    ['filter-customer', 'filter-model', 'filter-it-ticket', 'filter-open-sv', 'filter-device'].forEach(id => {
        document.getElementById(id).addEventListener('change', () => {
            pagination.currentPage = 1;
            applyFilters();
        });
    });

    // Pagination buttons
    document.getElementById('btn-page-prev').addEventListener('click', () => {
        if (pagination.currentPage > 1) {
            pagination.currentPage--;
            renderTable();
        }
    });

    document.getElementById('btn-page-next').addEventListener('click', () => {
        const totalPages = Math.ceil(filteredData.length / pagination.pageSize);
        if (pagination.currentPage < totalPages) {
            pagination.currentPage++;
            renderTable();
        }
    });

    // Table sorting
    const headers = document.querySelectorAll('#pm-table th');
    headers.forEach(header => {
        header.addEventListener('click', () => {
            const col = header.getAttribute('data-sort');
            if (currentSort.column === col) {
                currentSort.ascending = !currentSort.ascending;
            } else {
                currentSort.column = col;
                currentSort.ascending = true;
            }
            
            // Update sort arrows
            headers.forEach(h => {
                const icon = h.querySelector('.sort-icon');
                if (h.getAttribute('data-sort') === col) {
                    icon.textContent = currentSort.ascending ? '▲' : '▼';
                } else {
                    icon.textContent = '↕';
                }
            });

            sortData();
            renderTable();
        });
    });

    // CSV Export
    document.getElementById('btn-export').addEventListener('click', exportToCSV);
    // PDF Print / Save as PDF Link
    document.getElementById('btn-pdf').addEventListener('click', generatePDFLink);
}

// Show/Update data status indicator
function showBadgeStatus(text, themeClass) {
    const badge = document.getElementById('data-status-badge');
    const indicator = badge.querySelector('.indicator');
    const statusText = badge.querySelector('.status-text');
    
    statusText.textContent = text;
    indicator.className = 'indicator'; // Reset classes
    
    if (themeClass === 'green') {
        indicator.classList.add('pulse-green');
    } else if (themeClass === 'orange') {
        indicator.classList.add('pulse-orange');
    }
}

// Handle uploaded file
function handleExcelFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Get first sheet
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            
            // Convert to JSON array of arrays
            const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            if (jsonRows.length < 2) {
                alert('ไฟล์ Excel ไม่มีข้อมูลเพียงพอ');
                return;
            }
            
            // Parse custom Excel rows based on headers
            const parsedData = parseExcelData(jsonRows);
            if (parsedData.length === 0) {
                alert('ไม่สามารถอ่านข้อมูลได้ กรุณาตรวจสอบหัวคอลัมน์ของไฟล์');
                return;
            }
            
            loadDataset(parsedData, `อัปโหลดไฟล์: ${file.name}`);
        } catch (error) {
            console.error(error);
            alert('เกิดข้อผิดพลาดในการอ่านไฟล์ Excel: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// Map Excel columns dynamically based on Thai headers
function parseExcelData(rows) {
    // 1. Scan rows to find the actual header row dynamically (most matches)
    let headerRowIndex = 0;
    let maxMatches = 0;
    const targetHeaders = ['ลำดับ', 'บริษัทลูกค้า', 'รุ่นอุปกรณ์', 'เลขทะเบียน', 'สถานะ IT Ticket', 'เปิดงาน SV', 'สถานะอุปกรณ์'];

    for (let i = 0; i < Math.min(rows.length, 15); i++) {
        const row = rows[i];
        if (!row) continue;
        let matches = 0;
        row.forEach(cell => {
            if (cell !== undefined && cell !== null) {
                const cleanCell = cell.toString().replace(/\s+/g, '').trim();
                if (targetHeaders.some(th => cleanCell.includes(th))) {
                    matches++;
                }
            }
        });
        if (matches > maxMatches) {
            maxMatches = matches;
            headerRowIndex = i;
        }
    }

    const headerRow = rows[headerRowIndex];
    if (!headerRow) return [];

    // Helper to find column index
    const findIndex = (names) => {
        return headerRow.findIndex(h => {
            if (h === undefined || h === null) return false;
            const cleanHeader = h.toString().replace(/\s+/g, '').trim();
            return names.some(name => cleanHeader.includes(name));
        });
    };

    const idxId = findIndex(['ลำดับ', 'no']);
    const idxCustomer = findIndex(['บริษัทลูกค้า', 'ลูกค้า']);
    const idxModel = findIndex(['รุ่นอุปกรณ์', 'รุ่น']);
    const idxPlate = findIndex(['เลขทะเบียน', 'ทะเบียน']);
    const idxItTicket = findIndex(['สถานะITTicket', 'itticket']);
    const idxOpenSv = findIndex(['เปิดงานSV', 'sv']);
    const idxRepairDate = findIndex(['แจ้งซ่อมวันที่', 'เเจ้งซ่อมวันที่', 'วันที่แจ้งซ่อม']);
    let idxDevice = findIndex(['สถานะอุปกรณ์', 'สถานะเครื่อง', 'สถานะgps']);
    if (idxDevice === -1) {
        idxDevice = headerRow.findIndex(h => {
            if (h === undefined || h === null) return false;
            const cleanHeader = h.toString().replace(/\s+/g, '').trim();
            return cleanHeader.includes('อุปกรณ์') && !cleanHeader.includes('รุ่น');
        });
    }
    const idxDlt = findIndex(['DLTและMasterFile', 'dlt']);
    const idxCard = findIndex(['ค่าบัตรโดยรวม', 'ค่าบัตร']);
    const idxCamera = findIndex(['กล้องRealtime', 'กล้อง']);
    const idxFuel = findIndex(['น้ำมัน', 'fuel']);

    const getCellValue = (row, index) => {
        if (index === -1 || !row || row[index] === undefined || row[index] === null) {
            return '';
        }
        return row[index].toString().trim();
    };

    const result = [];
    
    // Parse from headerRowIndex + 1 to end
    for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        
        const plate = getCellValue(row, idxPlate);
        const customer = getCellValue(row, idxCustomer);
        
        // Skip empty rows
        if (!plate && !customer) continue;
        
        result.push({
            id: getCellValue(row, idxId) || (r - headerRowIndex).toString(),
            customer: customer || '(ว่าง)',
            model: getCellValue(row, idxModel),
            plate: plate || '(ไม่มีทะเบียน)',
            itTicket: getCellValue(row, idxItTicket),
            openSv: getCellValue(row, idxOpenSv),
            repairDate: getCellValue(row, idxRepairDate),
            deviceStatus: getCellValue(row, idxDevice),
            dltStatus: getCellValue(row, idxDlt),
            cardStatus: getCellValue(row, idxCard),
            camera: getCellValue(row, idxCamera),
            fuel: getCellValue(row, idxFuel)
        });
    }
    
    return result;
}

// Load dataset into memory, refresh interface
function loadDataset(data, statusText) {
    rawData = data;
    filteredData = [...data];
    currentSort = { column: 'id', ascending: true };
    pagination.currentPage = 1;
    
    showBadgeStatus(statusText, 'green');
    
    // Sort initially by ID
    sortData();
    
    // Populate filter options dynamically based on new dataset
    populateFilters();
    
    // Refresh components
    updateDashboard();
}

// Populate Filter Dropdowns dynamically
function populateFilters() {
    const customers = new Set();
    const models = new Set();
    const itTickets = new Set();
    const openSvs = new Set();

    rawData.forEach(item => {
        if (item.customer && item.customer !== '(ว่าง)') customers.add(item.customer);
        if (item.model) models.add(item.model);
        if (item.itTicket) itTickets.add(item.itTicket);
        if (item.openSv) openSvs.add(item.openSv);
    });

    const populateDropdown = (id, values) => {
        const select = document.getElementById(id);
        select.innerHTML = '<option value="all">ทั้งหมด</option>';
        
        Array.from(values).sort().forEach(val => {
            const opt = document.createElement('option');
            opt.value = val;
            opt.textContent = val;
            select.appendChild(opt);
        });
    };

    populateDropdown('filter-customer', customers);
    populateDropdown('filter-model', models);
    populateDropdown('filter-it-ticket', itTickets);
    populateDropdown('filter-open-sv', openSvs);
}

// Filter dataset based on inputs
function applyFilters() {
    const searchVal = document.getElementById('search-input').value.toLowerCase().trim();
    const filterCustomer = document.getElementById('filter-customer').value;
    const filterModel = document.getElementById('filter-model').value;
    const filterItTicket = document.getElementById('filter-it-ticket').value;
    const filterOpenSv = document.getElementById('filter-open-sv').value;
    const filterDevice = document.getElementById('filter-device').value;

    filteredData = rawData.filter(item => {
        const matchesSearch = !searchVal || 
            (item.plate && item.plate.toLowerCase().includes(searchVal)) || 
            (item.customer && item.customer.toLowerCase().includes(searchVal));

        const matchesCustomer = filterCustomer === 'all' || item.customer === filterCustomer;
        const matchesModel = filterModel === 'all' || item.model === filterModel;
        const matchesItTicket = filterItTicket === 'all' || item.itTicket === filterItTicket;
        const matchesOpenSv = filterOpenSv === 'all' || item.openSv === filterOpenSv;
        
        let matchesDevice = true;
        if (filterDevice === 'online') {
            matchesDevice = item.deviceStatus && item.deviceStatus.includes('ออนไลน์');
        } else if (filterDevice === 'offline') {
            matchesDevice = item.deviceStatus && item.deviceStatus.includes('ออฟไลน์');
        }

        return matchesSearch && matchesCustomer && matchesModel && matchesItTicket && matchesOpenSv && matchesDevice;
    });

    sortData();
    updateDashboardVisualsOnly();
}

// Sort data according to currentSort
function sortData() {
    const { column, ascending } = currentSort;
    
    filteredData.sort((a, b) => {
        let valA = a[column] || '';
        let valB = b[column] || '';
        
        if (column === 'id') {
            const numA = parseInt(valA, 10);
            const numB = parseInt(valB, 10);
            if (!isNaN(numA) && !isNaN(numB)) {
                return ascending ? numA - numB : numB - numA;
            }
        }
        
        valA = valA.toString().toLowerCase();
        valB = valB.toString().toLowerCase();
        
        if (valA < valB) return ascending ? -1 : 1;
        if (valA > valB) return ascending ? 1 : -1;
        return 0;
    });
}

// Update dashboard stats, charts, and table
function updateDashboard() {
    updateKPIs();
    renderCharts();
    renderTable();
}

// Update dashboard elements without rebuild charts from scratch
function updateDashboardVisualsOnly() {
    updateKPIs();
    updateChartsData();
    renderTable();
}

// Calculate and render KPI Card numbers
function updateKPIs() {
    const total = filteredData.length;
    
    const pmOpened = filteredData.filter(item => {
        if (!item.itTicket) return false;
        const t = item.itTicket.replace(/\s+/g, '');
        return t === 'แจ้งซ่อมPM' || t === 'เเจ้งซ่อมPM';
    }).length;
    
    const repairing = filteredData.filter(item => 
        item.openSv && item.openSv.includes('ตรวจสอบ')
    ).length;

    const online = filteredData.filter(item => 
        item.deviceStatus && item.deviceStatus.includes('ออนไลน์')
    ).length;

    document.getElementById('kpi-total').textContent = total.toLocaleString();
    document.getElementById('kpi-pm-opened').textContent = pmOpened.toLocaleString();
    document.getElementById('kpi-repairing').textContent = repairing.toLocaleString();
    document.getElementById('kpi-online').textContent = online.toLocaleString();

    const pmOpenedPercent = total > 0 ? ((pmOpened / total) * 100).toFixed(1) : 0;
    const repairingPercent = pmOpened > 0 ? ((repairing / pmOpened) * 100).toFixed(1) : 0;
    const onlinePercent = total > 0 ? ((online / total) * 100).toFixed(1) : 0;

    document.getElementById('pm-opened-percent').textContent = `${pmOpenedPercent}%`;
    document.getElementById('repairing-percent').textContent = `${repairingPercent}%`;
    document.getElementById('online-percent').textContent = `${onlinePercent}%`;
}

// Helper to parse dates and return formatted month key and label
function getMonthYear(dateStr) {
    if (!dateStr || dateStr.trim() === '' || dateStr === '-') return null;
    
    let month = null;
    let year = null;
    
    const partsSlash = dateStr.split('/');
    if (partsSlash.length === 3) {
        month = parseInt(partsSlash[1], 10);
        year = parseInt(partsSlash[2], 10);
    } else {
        const partsDash = dateStr.split('-');
        if (partsDash.length === 3) {
            year = parseInt(partsDash[0], 10);
            month = parseInt(partsDash[1], 10);
        }
    }
    
    if (month && year && !isNaN(month) && !isNaN(year)) {
        if (year < 100) {
            year += 2000;
        }
        return {
            key: year * 100 + month,
            month: month,
            year: year
        };
    }
    
    return null;
}

// Build counts maps for charts
function getChartStats() {
    const itTicketCounts = {};
    const openSvCounts = {};
    const deviceCounts = { 'ออนไลน์': 0, 'ออฟไลน์': 0, 'ไม่พบสถานะ/ว่าง': 0 };
    const customerCounts = {};
    const monthlyCounts = {};

    filteredData.forEach(item => {
        const tStatus = item.itTicket || '(ว่าง)';
        itTicketCounts[tStatus] = (itTicketCounts[tStatus] || 0) + 1;

        const sStatus = item.openSv || '(ว่าง)';
        openSvCounts[sStatus] = (openSvCounts[sStatus] || 0) + 1;

        if (item.deviceStatus && item.deviceStatus.includes('ออนไลน์')) {
            deviceCounts['ออนไลน์']++;
        } else if (item.deviceStatus && item.deviceStatus.includes('ออฟไลน์')) {
            deviceCounts['ออฟไลน์']++;
        } else {
            deviceCounts['ไม่พบสถานะ/ว่าง']++;
        }

        if (item.customer && item.customer !== '(ว่าง)') {
            customerCounts[item.customer] = (customerCounts[item.customer] || 0) + 1;
        }

        const my = getMonthYear(item.repairDate);
        if (my) {
            monthlyCounts[my.key] = (monthlyCounts[my.key] || 0) + 1;
        }
    });

    const topCustomers = Object.entries(customerCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    const THAI_MONTHS_SHORT = [
        'ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.',
        'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'
    ];

    const sortedMonthKeys = Object.keys(monthlyCounts).sort((a, b) => parseInt(a, 10) - parseInt(b, 10));
    
    const monthlyLabels = [];
    const monthlyValues = [];
    
    sortedMonthKeys.forEach(key => {
        const year = Math.floor(key / 100);
        const monthIdx = (key % 100) - 1;
        const monthName = THAI_MONTHS_SHORT[monthIdx] || '';
        const label = `${monthName} ${year}`;
        monthlyLabels.push(label);
        monthlyValues.push(monthlyCounts[key]);
    });

    return { itTicketCounts, openSvCounts, deviceCounts, topCustomers, monthlyLabels, monthlyValues };
}

// Render Charts using Chart.js
function renderCharts() {
    const stats = getChartStats();

    Object.keys(charts).forEach(key => {
        if (charts[key]) charts[key].destroy();
    });

    const commonOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                position: 'bottom',
                labels: { color: '#94a3b8', font: { family: 'Sarabun', size: 11 } }
            }
        }
    };

    const ctxIt = document.getElementById('chart-it-ticket').getContext('2d');
    charts.itTicket = new Chart(ctxIt, {
        type: 'pie',
        data: {
            labels: Object.keys(stats.itTicketCounts),
            datasets: [{
                data: Object.values(stats.itTicketCounts),
                backgroundColor: ['#818cf8', '#60a5fa', '#34d399', '#fbbf24', '#f87171', '#6b7280', '#f472b6', '#a78bfa'],
                borderWidth: 1.5,
                borderColor: '#151f32'
            }]
        },
        options: {
            ...commonOptions,
            plugins: { legend: { display: false } }
        }
    });

    const ctxSv = document.getElementById('chart-open-sv').getContext('2d');
    charts.openSv = new Chart(ctxSv, {
        type: 'pie',
        data: {
            labels: Object.keys(stats.openSvCounts),
            datasets: [{
                data: Object.values(stats.openSvCounts),
                backgroundColor: ['#34d399', '#fbbf24', '#6b7280', '#60a5fa', '#a78bfa'],
                borderWidth: 1.5,
                borderColor: '#151f32'
            }]
        },
        options: {
            ...commonOptions,
            plugins: { legend: { display: false } }
        }
    });

    const ctxDev = document.getElementById('chart-device-status').getContext('2d');
    charts.deviceStatus = new Chart(ctxDev, {
        type: 'doughnut',
        data: {
            labels: Object.keys(stats.deviceCounts),
            datasets: [{
                data: Object.values(stats.deviceCounts),
                backgroundColor: ['#34d399', '#f87171', '#64748b'],
                borderWidth: 1.5,
                borderColor: '#151f32'
            }]
        },
        options: {
            ...commonOptions,
            cutout: '65%',
            plugins: { legend: { display: false } }
        }
    });

    const custLabels = stats.topCustomers.map(x => x[0]);
    const custData = stats.topCustomers.map(x => x[1]);
    const ctxCust = document.getElementById('chart-top-customers').getContext('2d');
    charts.topCustomers = new Chart(ctxCust, {
        type: 'bar',
        data: {
            labels: custLabels,
            datasets: [{
                label: 'จำนวนงาน PM (คัน)',
                data: custData,
                backgroundColor: 'rgba(52, 211, 153, 0.75)',
                borderColor: '#34d399',
                borderWidth: 1.5,
                borderRadius: 4
            }]
        },
        options: {
            ...commonOptions,
            indexAxis: 'y',
            plugins: {
                legend: { display: false }
            },
            scales: {
                x: {
                    grid: { color: 'rgba(255, 255, 255, 0.06)' },
                    ticks: { color: '#94a3b8', font: { family: 'Sarabun' } }
                },
                y: {
                    grid: { display: false },
                    ticks: { color: '#94a3b8', font: { family: 'Sarabun' } }
                }
            }
        }
    });

    // Populate custom legends initially
    updateCustomLegend(charts.itTicket, 'legend-it-ticket');
    updateCustomLegend(charts.openSv, 'legend-open-sv');
    updateCustomLegend(charts.deviceStatus, 'legend-device-status');
}

// Update custom HTML legend for circular charts
function updateCustomLegend(chart, legendId) {
    const legendContainer = document.getElementById(legendId);
    if (!legendContainer || !chart) return;

    const data = chart.data;
    const dataset = data.datasets[0];
    const total = dataset.data.reduce((a, b) => a + b, 0);

    let html = '';
    data.labels.forEach((label, index) => {
        const val = dataset.data[index];
        const percent = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
        const color = dataset.backgroundColor[index] || '#cbd5e1';

        html += `
            <div class="chart-legend-item">
                <div class="chart-legend-color" style="background-color: ${color};"></div>
                <div class="chart-legend-content">
                    <div class="chart-legend-text">${label}</div>
                    <div class="chart-legend-value">${val.toLocaleString()} คัน (${percent}%)</div>
                </div>
            </div>
        `;
    });

    legendContainer.innerHTML = html;
}

// Update charts dataset dynamically
function updateChartsData() {
    const stats = getChartStats();

    if (charts.itTicket) {
        charts.itTicket.data.labels = Object.keys(stats.itTicketCounts);
        charts.itTicket.data.datasets[0].data = Object.values(stats.itTicketCounts);
        charts.itTicket.update();
        updateCustomLegend(charts.itTicket, 'legend-it-ticket');
    }

    if (charts.openSv) {
        charts.openSv.data.labels = Object.keys(stats.openSvCounts);
        charts.openSv.data.datasets[0].data = Object.values(stats.openSvCounts);
        charts.openSv.update();
        updateCustomLegend(charts.openSv, 'legend-open-sv');
    }

    if (charts.deviceStatus) {
        charts.deviceStatus.data.labels = Object.keys(stats.deviceCounts);
        charts.deviceStatus.data.datasets[0].data = Object.values(stats.deviceCounts);
        charts.deviceStatus.update();
        updateCustomLegend(charts.deviceStatus, 'legend-device-status');
    }

    if (charts.topCustomers) {
        charts.topCustomers.data.labels = stats.topCustomers.map(x => x[0]);
        charts.topCustomers.data.datasets[0].data = stats.topCustomers.map(x => x[1]);
        charts.topCustomers.update();
    }
}

// Render data table and pagination controls
function renderTable() {
    const tbody = document.getElementById('table-body');
    tbody.innerHTML = '';

    const totalRecords = filteredData.length;
    document.getElementById('pagination-total').textContent = totalRecords.toLocaleString();

    if (totalRecords === 0) {
        tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; color: var(--text-muted); padding: 2rem;">ไม่พบข้อมูลตามเงื่อนไขที่เลือก</td></tr>';
        document.getElementById('pagination-start').textContent = '0';
        document.getElementById('pagination-end').textContent = '0';
        renderPaginationControls(0);
        return;
    }

    const startIdx = (pagination.currentPage - 1) * pagination.pageSize;
    const endIdx = Math.min(startIdx + pagination.pageSize, totalRecords);
    
    document.getElementById('pagination-start').textContent = (startIdx + 1).toLocaleString();
    document.getElementById('pagination-end').textContent = endIdx.toLocaleString();

    const pageItems = filteredData.slice(startIdx, endIdx);
    pageItems.forEach(item => {
        const tr = document.createElement('tr');
        
        let itTicketClass = 'badge-secondary';
        const cleanItTicket = item.itTicket ? item.itTicket.replace(/\s+/g, '') : '';
        if (cleanItTicket === 'แจ้งซ่อมPM' || cleanItTicket === 'เเจ้งซ่อมPM') {
            itTicketClass = 'badge-indigo';
        } else if (cleanItTicket === 'ไม่แจ้งPM' || cleanItTicket === 'ไม่เเจ้งPM') {
            itTicketClass = 'badge-blue';
        } else if (item.itTicket && item.itTicket.includes('ตรวจสอบ')) {
            itTicketClass = 'badge-orange';
        } else if (item.itTicket && item.itTicket.includes('ยกเลิก')) {
            itTicketClass = 'badge-red';
        }

        let openSvClass = 'badge-secondary';
        if (item.openSv && item.openSv.includes('ตรวจสอบ')) {
            openSvClass = 'badge-orange';
        } else if (item.openSv && (item.openSv.includes('ไม่แจ้ง') || item.openSv.includes('ไม่เเจ้ง'))) {
            openSvClass = 'badge-blue';
        }

        let deviceClass = 'badge-secondary';
        if (item.deviceStatus && item.deviceStatus.includes('ออนไลน์')) {
            deviceClass = 'badge-teal';
        } else if (item.deviceStatus && item.deviceStatus.includes('ออฟไลน์')) {
            deviceClass = 'badge-red';
        }

        tr.innerHTML = `
            <td>${item.id}</td>
            <td style="font-weight: 600; color: var(--color-blue); font-family: var(--font-heading);">${item.plate}</td>
            <td>${item.customer}</td>
            <td style="font-family: var(--font-heading);">${item.model}</td>
            <td><span class="badge ${itTicketClass}"><span class="badge-dot"></span>${item.itTicket || '-'}</span></td>
            <td><span class="badge ${openSvClass}"><span class="badge-dot"></span>${item.openSv || '-'}</span></td>
            <td style="font-family: var(--font-heading);">${item.repairDate || '-'}</td>
            <td><span class="badge ${deviceClass}" title="${item.deviceStatus}"><span class="badge-dot"></span>${item.deviceStatus.split('|')[0] || '-'}</span></td>
        `;
        tbody.appendChild(tr);
    });

    renderPaginationControls(totalPagesCount());
}

// Calculate total pages count
function totalPagesCount() {
    return Math.ceil(filteredData.length / pagination.pageSize);
}

// Render pagination buttons
function renderPaginationControls(totalPages) {
    const container = document.getElementById('page-numbers-container');
    container.innerHTML = '';

    const btnPrev = document.getElementById('btn-page-prev');
    const btnNext = document.getElementById('btn-page-next');

    btnPrev.disabled = pagination.currentPage === 1;
    btnNext.disabled = pagination.currentPage === totalPages || totalPages === 0;

    if (totalPages <= 1) return;

    const maxVisiblePages = 5;
    let startPage = Math.max(1, pagination.currentPage - Math.floor(maxVisiblePages / 2));
    let endPage = Math.min(totalPages, startPage + maxVisiblePages - 1);

    if (endPage - startPage + 1 < maxVisiblePages) {
        startPage = Math.max(1, endPage - maxVisiblePages + 1);
    }

    for (let p = startPage; p <= endPage; p++) {
        const btn = document.createElement('button');
        btn.className = `btn-page ${p === pagination.currentPage ? 'active' : ''}`;
        btn.textContent = p;
        btn.addEventListener('click', () => {
            pagination.currentPage = p;
            renderTable();
        });
        container.appendChild(btn);
    }
}

// Export Filtered Table Data to CSV
function exportToCSV() {
    if (filteredData.length === 0) {
        alert('ไม่มีข้อมูลสำหรับส่งออก');
        return;
    }

    const headers = ['ลำดับ', 'เลขทะเบียน', 'บริษัทลูกค้า', 'รุ่นอุปกรณ์', 'สถานะ IT Ticket', 'เปิดงาน SV', 'วันที่เเจ้งซ่อม', 'สถานะอุปกรณ์', 'DLT และ MasterFile', 'ค่าบัตรโดยรวม', 'กล้อง Realtime', 'น้ำมัน'];
    const rows = filteredData.map(item => [
        item.id,
        item.plate,
        item.customer,
        item.model,
        item.itTicket,
        item.openSv,
        item.repairDate,
        item.deviceStatus,
        item.dltStatus,
        item.cardStatus,
        item.camera,
        item.fuel
    ]);

    let csvContent = '\uFEFF';
    csvContent += headers.map(h => `"${h.replace(/"/g, '""')}"`).join(',') + '\n';
    
    rows.forEach(row => {
        csvContent += row.map(cell => `"${(cell || '').replace(/"/g, '""')}"`).join(',') + '\n';
    });

    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', `PM_Report_Export_${new Date().toISOString().slice(0,10)}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// Listen for print events to swap chart theme colors dynamically
window.addEventListener('beforeprint', () => {
    updateChartsTheme(true);
});

window.addEventListener('afterprint', () => {
    updateChartsTheme(false);
});

function updateChartsTheme(isPrint) {
    const textColor = isPrint ? '#4b5563' : '#94a3b8';
    const gridColor = isPrint ? 'rgba(0, 0, 0, 0.05)' : 'rgba(255, 255, 255, 0.06)';
    const borderColor = isPrint ? '#ffffff' : '#151f32';

    Object.keys(charts).forEach(key => {
        const chart = charts[key];
        if (!chart) return;

        // Update slices borders for pie/doughnut charts
        if (chart.data.datasets && chart.data.datasets[0]) {
            chart.data.datasets[0].borderColor = borderColor;
        }

        // Update legend label colors
        if (chart.options.plugins && chart.options.plugins.legend) {
            if (chart.options.plugins.legend.labels) {
                chart.options.plugins.legend.labels.color = textColor;
            }
        }

        // Update scales ticks and grid lines for bar charts
        if (chart.options.scales) {
            if (chart.options.scales.x) {
                if (chart.options.scales.x.ticks) chart.options.scales.x.ticks.color = textColor;
                if (chart.options.scales.x.grid) chart.options.scales.x.grid.color = gridColor;
            }
            if (chart.options.scales.y) {
                if (chart.options.scales.y.ticks) chart.options.scales.y.ticks.color = textColor;
                if (chart.options.scales.y.grid) chart.options.scales.y.grid.color = gridColor;
            }
        }

        chart.update('none'); // Update instantly without animations
    });
}

function generatePDFLink() {
    const badge = document.getElementById('data-status-badge');
    const oldStatus = badge.querySelector('.status-text').textContent;
    const oldClass = badge.querySelector('.indicator').className;
    
    showBadgeStatus('กำลังสร้างไฟล์ PDF...', 'orange');

    // Set generation timestamp in Thai format for the footer
    const formattedTime = new Date().toLocaleString('th-TH', { timeZone: 'Asia/Bangkok' });
    document.getElementById('pdf-generation-time').textContent = 'เวลาที่พิมพ์: ' + formattedTime;

    // Add class to body to trigger dark-theme desktop capture styling
    document.body.classList.add('generating-pdf');
    updateChartsTheme(true);

    // Wait 500ms for browser layout recalculation and chart resize
    setTimeout(() => {
        const element = document.querySelector('.app-container');

        const opt = {
            margin:       10,
            filename:     `PM_Dashboard_Report_${new Date().toISOString().slice(0,10)}.pdf`,
            image:        { type: 'jpeg', quality: 0.98 },
            html2canvas:  { scale: 2, useCORS: true, logging: false, width: 1200, scrollY: 0, backgroundColor: '#0b0f19' },
            jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' },
            pagebreak:    { mode: ['css', 'legacy'] }
        };

        html2pdf().from(element).set(opt).toPdf().output('bloburl').then((pdfBlobUrl) => {
            document.body.classList.remove('generating-pdf');
            updateChartsTheme(false);
            
            badge.querySelector('.status-text').textContent = oldStatus;
            badge.querySelector('.indicator').className = oldClass;

            window.open(pdfBlobUrl, '_blank');
        }).catch(err => {
            console.error(err);
            alert('เกิดข้อผิดพลาดในการสร้างไฟล์ PDF: ' + err.message);
            document.body.classList.remove('generating-pdf');
            updateChartsTheme(false);
            
            badge.querySelector('.status-text').textContent = oldStatus;
            badge.querySelector('.indicator').className = oldClass;
        });
    }, 500);
}
