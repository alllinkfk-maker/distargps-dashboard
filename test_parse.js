const XLSX = require('./xlsx.full.min.js');
const fs = require('fs');

try {
    const filePath = 'C:/Users/it-support706.dsi/Desktop/All PM DistarGPS.xlsx';
    if (!fs.existsSync(filePath)) {
        console.error('Excel file does not exist at ' + filePath);
        process.exit(1);
    }

    const fileBuffer = fs.readFileSync(filePath);
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    console.log('Total rows loaded:', rows.length);
    if (rows.length < 2) {
        console.error('Not enough rows in Excel');
        process.exit(1);
    }

    // 1. Scan rows to find actual header row
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
                if (targetHeaders.some(th => cleanCell.includes(th.replace(/\s+/g, '')))) {
                    matches++;
                }
            }
        });
        if (matches > maxMatches) {
            maxMatches = matches;
            headerRowIndex = i;
        }
    }

    console.log('Detected Header Row Index:', headerRowIndex);
    console.log('Header Row Content:', rows[headerRowIndex]);

    const headerRow = rows[headerRowIndex];
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

    console.log('Detected Indices:');
    console.log('- Id:', idxId);
    console.log('- Customer:', idxCustomer);
    console.log('- Model:', idxModel);
    console.log('- Plate:', idxPlate);
    console.log('- IT Ticket:', idxItTicket);
    console.log('- Open SV:', idxOpenSv);
    console.log('- Repair Date:', idxRepairDate);
    console.log('- Device Status:', idxDevice);

    const result = [];
    const getCellValue = (row, index) => {
        if (index === -1 || !row || row[index] === undefined || row[index] === null) {
            return '';
        }
        return row[index].toString().trim();
    };

    for (let r = headerRowIndex + 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        const plate = getCellValue(row, idxPlate);
        const customer = getCellValue(row, idxCustomer);
        if (!plate && !customer) continue;
        result.push({
            id: getCellValue(row, idxId) || (r - headerRowIndex).toString(),
            customer: customer || '(ว่าง)',
            model: getCellValue(row, idxModel),
            plate: plate || '(ไม่มีทะเบียน)'
        });
    }

    console.log('Successfully parsed rows count:', result.length);
    console.log('First parsed item sample:', result[0]);

} catch (err) {
    console.error('Error running parser test:', err);
}
