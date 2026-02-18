// ============================================
// Google Apps Script untuk RSVP & Ucapan
// Paste kode ini di Extensions > Apps Script
// ============================================

const SHEET_NAME = 'RSVP';

function doGet(e) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ data: [] }))
            .setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);

    const result = rows.map(row => ({
        name: row[0],
        text: row[1],
        attendance: row[2],
        count: row[3],
        timestamp: row[4]
    })).reverse(); // Terbaru di atas

    return ContentService.createTextOutput(JSON.stringify({ data: result }))
        .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

    if (!sheet) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const newSheet = ss.insertSheet(SHEET_NAME);
        newSheet.appendRow(['Nama', 'Ucapan', 'Kehadiran', 'Jumlah', 'Waktu']);
        return processPost(newSheet, e);
    }

    return processPost(sheet, e);
}

function processPost(sheet, e) {
    try {
        const data = JSON.parse(e.postData.contents);
        const timestamp = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });

        sheet.appendRow([
            data.name || '',
            data.message || '',
            data.attendance || '',
            data.count || '',
            timestamp
        ]);

        return ContentService.createTextOutput(JSON.stringify({
            status: 'success',
            message: 'Data berhasil disimpan'
        })).setMimeType(ContentService.MimeType.JSON);

    } catch (error) {
        return ContentService.createTextOutput(JSON.stringify({
            status: 'error',
            message: error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

// Jalankan ini SEKALI untuk membuat sheet header
function setupSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
    }

    sheet.clear();
    sheet.appendRow(['Nama', 'Ucapan', 'Kehadiran', 'Jumlah', 'Waktu']);

    // Format header
    const headerRange = sheet.getRange(1, 1, 1, 5);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4a90d9');
    headerRange.setFontColor('#ffffff');

    // Set column widths
    sheet.setColumnWidth(1, 200); // Nama
    sheet.setColumnWidth(2, 400); // Ucapan
    sheet.setColumnWidth(3, 150); // Kehadiran
    sheet.setColumnWidth(4, 100); // Jumlah
    sheet.setColumnWidth(5, 200); // Waktu
}
