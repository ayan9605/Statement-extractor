const { parentPort, workerData } = require('worker_threads');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');

async function processPDF() {
    try {
        const { pdfPath, outputPath } = workerData;
        const dataBuffer = fs.readFileSync(pdfPath);

        // 1. Extract Text
        const pdfData = await pdfParse(dataBuffer);
        const rawText = pdfData.text;

        // 2. Data Cleaning (The Raw Logic to Build Upon)
        // This splits the text and filters out empty space. 
        const rows = rawText.split('\n').filter(line => line.trim() !== '');

        // 3. Build Excel File Memory Stream
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Transactions');

        // Professional styling for headers
        worksheet.columns = [
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Transaction Details', key: 'desc', width: 60 },
            { header: 'Debit/Credit', key: 'amount', width: 20 }
        ];
        
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF007BFF' } };

        // Insert data (Replace this with your custom Regex parsing logic)
        rows.forEach(row => {
            worksheet.addRow({ date: 'DD/MM/YYYY', desc: row, amount: '0.00' });
        });

        // 4. Write to disk
        await workbook.xlsx.writeFile(outputPath);

        // Notify main thread of success
        parentPort.postMessage({ success: true, outputPath });

    } catch (error) {
        console.error("Worker Exception:", error);
        parentPort.postMessage({ success: false, error: error.message });
    }
}

processPDF();
