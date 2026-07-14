/* ==========================================================================
   JavaScript Logic - GPS PM Dashboard
   ========================================================================== */

// Global State
let rawData = [];
let filteredData = [];
let currentSort = { column: 'id', ascending: true };
let pagination = { currentPage: 1, pageSize: 15 };

// Chart Instances (to destroy and rebuild on data reload)
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
    const headerRow = rows[0];
    
    // Find column indexes based on header names (case-insensitive & space trimmed)
    const findIndex = (names) => {
        return headerRow.findIndex(h => {
            if (!h) return false;
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
    const idxDevice = findIndex(['สถานะอุปกรณ์', 'อุปกรณ์']);
    const idxDlt = findIndex(['DLTและMasterFile', 'dlt']);
    const idxCard = findIndex(['ค่าบัตรโดยรวม', 'ค่าบัตร']);
    const idxCamera = findIndex(['กล้องRealtime', 'กล้อง']);
    const idxFuel = findIndex(['น้ำมัน', 'fuel']);

    const result = [];
    
    for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        
        // Skip empty rows or safety checks
        const plate = idxPlate !== -1 ? (row[idxPlate] || '').toString().trim() : '';
        const customer = idxCustomer !== -1 ? (row[idxCustomer] || '').toString().trim() : '';
        
        if (!plate && !customer) continue;
        
        result.push({
            id: idxId !== -1 && row[idxId] !== undefined ? row[idxId].toString().trim() : r.toString(),
            customer: customer || '(ว่าง)',
            model: idxModel !== -1 && row[idxModel] !== undefined ? row[idxModel].toString().trim() : '',
            plate: plate || '(ไม่มีทะเบียน)',
            itTicket: idxItTicket !== -1 && row[idxItTicket] !== undefined ? row[idxItTicket].toString().trim() : '',
            openSv: idxOpenSv !== -1 && row[idxOpenSv] !== undefined ? row[idxOpenSv].toString().trim() : '',
            repairDate: idxRepairDate !== -1 && row[idxRepairDate] !== undefined ? row[idxRepairDate].toString().trim() : '',
            deviceStatus: idxDevice !== -1 && row[idxDevice] !== undefined ? row[idxDevice].toString().trim() : '',
            dltStatus: idxDlt !== -1 && row[idxDlt] !== undefined ? row[idxDlt].toString().trim() : '',
            cardStatus: idxCard !== -1 && row[idxCard] !== undefined ? row[idxCard].toString().trim() : '',
            camera: idxCamera !== -1 && row[idxCamera] !== undefined ? row[idxCamera].toString().trim() : '',
            fuel: idxFuel !== -1 && row[idxFuel] !== undefined ? row[idxFuel].toString().trim() : ''
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
    
    // Sort initially by ID (cast numerical string if possible)
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
        // Clear previous options except "All"
        select.innerHTML = '<option value="all">ทั้งหมด</option>';
        
        // Sort and add new options
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
        // Search text matching plate or customer
        const matchesSearch = !searchVal || 
            item.plate.toLowerCase().includes(searchVal) || 
            item.customer.toLowerCase().includes(searchVal);

        // Dropdown filters matching
        const matchesCustomer = filterCustomer === 'all' || item.customer === filterCustomer;
        const matchesModel = filterModel === 'all' || item.model === filterModel;
        const matchesItTicket = filterItTicket === 'all' || item.itTicket === filterItTicket;
        const matchesOpenSv = filterOpenSv === 'all' || item.openSv === filterOpenSv;
        
        let matchesDevice = true;
        if (filterDevice === 'online') {
            matchesDevice = item.deviceStatus.includes('ออนไลน์');
        } else if (filterDevice === 'offline') {
            matchesDevice = item.deviceStatus.includes('ออฟไลน์');
        }

        return matchesSearch && matchesCustomer && matchesModel && matchesItTicket && matchesOpenSv && matchesDevice;
    });

    sortData();
    updateDashboardVisualsOnly(); // Update KPI cards, Charts, and Table based on filters
}

// Sort data according to currentSort
function sortData() {
    const { column, ascending } = currentSort;
    
    filteredData.sort((a, b) => {
        let valA = a[column] || '';
        let valB = b[column] || '';
        
        // If sorting by ID/ลำดับ, convert to numbers for natural sort
        if (column === 'id') {
            const numA = parseInt(valA, 10);
            const numB = parseInt(valB, 10);
            if (!isNaN(numA) && !isNaN(numB)) {
                return ascending ? numA - numB : numB - numA;
            }
        }
        
        // String comparison
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
    
    // PM Opened = status matches "แจ้งซ่อม PM"
    const pmOpened = filteredData.filter(item => 
        item.itTicket.replace(/\s+/g, '') === 'แจ้งซ่อมPM' || 
        item.itTicket.replace(/\s+/g, '') === 'เเจ้งซ่อมPM'
    ).length;
    
    // Repairing / Checking = SV status is "อยู่ระหว่างตรวจสอบ"
    const repairing = filteredData.filter(item => 
        item.openSv.includes('ตรวจสอบ')
    ).length;

    // Device Online = status contains "ออนไลน์"
    const online = filteredData.filter(item => 
        item.deviceStatus.includes('ออนไลน์')
    ).length;

    // Inject into UI
    document.getElementById('kpi-total').textContent = total.toLocaleString();
    document.getElementById('kpi-pm-opened').textContent = pmOpened.toLocaleString();
    document.getElementById('kpi-repairing').textContent = repairing.toLocaleString();
    document.getElementById('kpi-online').textContent = online.toLocaleString();

    // Percentages
    const pmOpenedPercent = total > 0 ? ((pmOpened / total) * 100).toFixed(1) : 0;
    const repairingPercent = pmOpened > 0 ? ((repairing / pmOpened) * 100).toFixed(1) : 0; // relative to PM opened
    const onlinePercent = total > 0 ? ((online / total) * 100).toFixed(1) : 0;

    document.getElementById('pm-opened-percent').textContent = `${pmOpenedPercent}%`;
    document.getElementById('repairing-percent').textContent = `${repairingPercent}%`;
    document.getElementById('online-percent').textContent = `${onlinePercent}%`;
}

// Build counts maps for charts
function getChartStats() {
    const itTicketCounts = {};
    const openSvCounts = {};
    const deviceCounts = { 'ออนไลน์': 0, 'ออฟไลน์': 0, 'ไม่พบสถานะ/ว่าง': 0 };
    const customerCounts = {};

    filteredData.forEach(item => {
        // IT Ticket Status
        const tStatus = item.itTicket || '(ว่าง)';
        itTicketCounts[tStatus] = (itTicketCounts[tStatus] || 0) + 1;

        // Open SV Status
        const sStatus = item.openSv || '(ว่าง)';
        openSvCounts[sStatus] = (openSvCounts[sStatus] || 0) + 1;

        // Device Connection Summary
        if (item.deviceStatus.includes('ออนไลน์')) {
            deviceCounts['ออนไลน์']++;
        } else if (item.deviceStatus.includes('ออฟไลน์')) {
            deviceCounts['ออฟไลน์']++;
        } else {
            deviceCounts['ไม่พบสถานะ/ว่าง']++;
        }

        // Customer (only count those with actual customer names)
        if (item.customer && item.customer !== '(ว่าง)') {
            customerCounts[item.customer] = (customerCounts[item.customer] || 0) + 1;
        }
    });

    // Format Top 5 Customers
    const topCustomers = Object.entries(customerCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5);

    return { itTicketCounts, openSvCounts, deviceCounts, topCustomers };
}

// Render Charts using Chart.js
function renderCharts() {
    const stats = getChartStats();

    // Destroy existing chart instances first
    Object.keys(charts).forEach(key => {
        if (charts[key]) charts[key].destroy();
    });

    const commonOptions = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                position: 'bottom',
                labels: { color: '#e5e7eb', font: { family: 'Sarabun', size: 11 } }
            }
        }
    };

    // 1. IT Ticket Chart
    const itLabels = Object.keys(stats.itTicketCounts);
    const itData = Object.values(stats.itTicketCounts);
    const ctxIt = document.getElementById('chart-it-ticket').getContext('2d');
    charts.itTicket = new Chart(ctxIt, {
        type: 'pie',
        data: {
            labels: itLabels,
            datasets: [{
                data: itData,
                backgroundColor: ['#6366f1', '#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#6b7280', '#ec4899', '#8b5cf6'],
                borderWidth: 1,
                borderColor: '#1f2937'
            }]
        },
        options: commonOptions
    });

    // 2. Open SV Chart
    const svLabels = Object.keys(stats.openSvCounts);
    const svData = Object.values(stats.openSvCounts);
    const ctxSv = document.getElementById('chart-open-sv').getContext('2d');
    charts.openSv = new Chart(ctxSv, {
        type: 'pie',
        data: {
            labels: svLabels,
            datasets: [{
                data: svData,
                backgroundColor: ['#10b981', '#f59e0b', '#6b7280', '#3b82f6', '#8b5cf6'],
                borderWidth: 1,
                borderColor: '#1f2937'
            }]
        },
        options: commonOptions
    });

    // 3. Device Status Chart
    const devLabels = Object.keys(stats.deviceCounts);
    const devData = Object.values(stats.deviceCounts);
    const ctxDev = document.getElementById('chart-device-status').getContext('2d');
    charts.deviceStatus = new Chart(ctxDev, {
        type: 'doughnut',
        data: {
            labels: devLabels,
            datasets: [{
                data: devData,
                backgroundColor: ['#10b981', '#ef4444', '#4b5563'],
                borderWidth: 1,
                borderColor: '#1f2937'
            }]
        },
        options: {
            ...commonOptions,
            cutout: '65%'
        }
    });

    // 4. Top Customers Chart
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
                backgroundColor: 'rgba(59, 130, 246, 0.7)',
                borderColor: '#3b82f6',
                borderWidth: 1,
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
                    grid: { color: 'rgba(255,255,255,0.05)' },
                    ticks: { color: '#e5e7eb', font: { family: 'Sarabun' } }
                },
                y: {
                    grid: { display: false },
                    ticks: { color: '#e5e7eb', font: { family: 'Sarabun' } }
                }
            }
        }
    });
}

// Update charts dataset dynamically (no flashing rebuild)
function updateChartsData() {
    const stats = getChartStats();

    if (charts.itTicket) {
        charts.itTicket.data.labels = Object.keys(stats.itTicketCounts);
        charts.itTicket.data.datasets[0].data = Object.values(stats.itTicketCounts);
        charts.itTicket.update();
    }

    if (charts.openSv) {
        charts.openSv.data.labels = Object.keys(stats.openSvCounts);
        charts.openSv.data.datasets[0].data = Object.values(stats.openSvCounts);
        charts.openSv.update();
    }

    if (charts.deviceStatus) {
        charts.deviceStatus.data.labels = Object.keys(stats.deviceCounts);
        charts.deviceStatus.data.datasets[0].data = Object.values(stats.deviceCounts);
        charts.deviceStatus.update();
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

    // Pagination calculations
    const startIdx = (pagination.currentPage - 1) * pagination.pageSize;
    const endIdx = Math.min(startIdx + pagination.pageSize, totalRecords);
    
    document.getElementById('pagination-start').textContent = (startIdx + 1).toLocaleString();
    document.getElementById('pagination-end').textContent = endIdx.toLocaleString();

    // Render rows
    const pageItems = filteredData.slice(startIdx, endIdx);
    pageItems.forEach(item => {
        const tr = document.createElement('tr');
        
        // IT Ticket Status badge formatting
        let itTicketClass = 'badge-secondary';
        const cleanItTicket = item.itTicket.replace(/\s+/g, '');
        if (cleanItTicket === 'แจ้งซ่อมPM' || cleanItTicket === 'เเจ้งซ่อมPM') {
            itTicketClass = 'badge-indigo';
        } else if (cleanItTicket === 'ไม่แจ้งPM' || cleanItTicket === 'ไม่เเจ้งPM') {
            itTicketClass = 'badge-blue';
        } else if (item.itTicket.includes('ตรวจสอบ')) {
            itTicketClass = 'badge-orange';
        } else if (item.itTicket.includes('ยกเลิก')) {
            itTicketClass = 'badge-red';
        }

        // Open SV status badge formatting
        let openSvClass = 'badge-secondary';
        if (item.openSv.includes('ตรวจสอบ')) {
            openSvClass = 'badge-orange';
        } else if (item.openSv.includes('ไม่แจ้ง') || item.openSv.includes('ไม่เเจ้ง')) {
            openSvClass = 'badge-blue';
        }

        // Device Connection badge formatting
        let deviceClass = 'badge-secondary';
        if (item.deviceStatus.includes('ออนไลน์')) {
            deviceClass = 'badge-teal';
        } else if (item.deviceStatus.includes('ออฟไลน์')) {
            deviceClass = 'badge-red';
        }

        tr.innerHTML = `
            <td>${item.id}</td>
            <td style="font-weight: 600; color: var(--color-blue); font-family: var(--font-heading);">${item.plate}</td>
            <td>${item.customer}</td>
            <td style="font-family: var(--font-heading);">${item.model}</td>
            <td><span class="badge ${itTicketClass}">${item.itTicket || '-'}</span></td>
            <td><span class="badge ${openSvClass}">${item.openSv || '-'}</span></td>
            <td style="font-family: var(--font-heading);">${item.repairDate || '-'}</td>
            <td><span class="badge ${deviceClass}" title="${item.deviceStatus}">${item.deviceStatus.split('|')[0] || '-'}</span></td>
        `;
        tbody.appendChild(tr);
    });

    renderPaginationControls(totalPagesCount());
}

// Calculate total pages count
function totalPagesCount() {
    return Math.ceil(filteredData.length / pagination.pageSize);
}

// Render pagination buttons dynamic sizing
function renderPaginationControls(totalPages) {
    const container = document.getElementById('page-numbers-container');
    container.innerHTML = '';

    const btnPrev = document.getElementById('btn-page-prev');
    const btnNext = document.getElementById('btn-page-next');

    // Toggle disabled state
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

    // CSV format generation
    let csvContent = '\uFEFF'; // Add UTF-8 BOM to prevent Thai garbled characters in Excel
    csvContent += headers.map(h => `"${h.replace(/"/g, '""')}"`).join(',') + '\n';
    
    rows.forEach(row => {
        csvContent += row.map(cell => `"${(cell || '').replace(/"/g, '""')}"`).join(',') + '\n';
    });

    // Download trigger
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
