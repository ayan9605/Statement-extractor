const { parentPort, workerData } = require('worker_threads');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const ExcelJS = require('exceljs');

async function processPDF() {
    try {
        const { pdfPath, outputPath } = workerData;
        const dataBuffer = fs.readFileSync(pdfPath);

        const pdfData = await pdfParse(dataBuffer);
        const rawText = pdfData.text;
        const rows = rawText.split('\n').filter(line => line.trim() !== '');

        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Transactions');

        // 1. Exact 5-Column Layout Matching Your Goal
        worksheet.columns = [
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Description', key: 'desc', width: 85 },
            { header: 'Debit', key: 'debit', width: 18 },
            { header: 'Credit', key: 'credit', width: 18 },
            { header: 'Balance', key: 'balance', width: 18 }
        ];
        
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF007BFF' } };

        let tableStarted = false;
        let currentTxn = null;
        
        const dateRegex = /^(\d{2}\s[A-Za-z]{3}\s'\d{2})(.*)/; 
        const amountRegex = /₹([\d,]+(?:\.\d+)?)/g;

        for (let i = 0; i < rows.length; i++) {
            let line = rows[i].trim();

            if (line.includes('DATEDETAILSREF NO.AMOUNTBALANCE') || line.includes('DATEDETAILS')) {
                tableStarted = true;
                continue;
            }
            if (!tableStarted) continue;
            if (line.includes('Need help?') || line.includes('Closing balance')) break;

            const dateMatch = line.match(dateRegex);

            if (dateMatch) {
                if (currentTxn && currentTxn.date) {
                     worksheet.addRow(currentTxn);
                }

                currentTxn = {
                    date: dateMatch[1],
                    desc: dateMatch[2], 
                    debit: '',
                    credit: '',
                    balance: ''
                };
            } else if (currentTxn) {
                const amounts = [...line.matchAll(amountRegex)];
                
                if (amounts.length >= 2) {
                    const textBeforeAmounts = line.split('₹')[0].trim();
                    currentTxn.desc += ' ' + textBeforeAmounts;

                    const txnAmount = parseFloat(amounts[0][1].replace(/,/g, ''));
                    currentTxn.balance = parseFloat(amounts[amounts.length - 1][1].replace(/,/g, '')); 

                    // 2. Precision Text Cleaning
                    // Safely removes the '3' delimiter only if it touches a letter, protecting your reference numbers.
                    let cleanDesc = currentTxn.desc.replace(/(?<=[a-zA-Z])3|3(?=[a-zA-Z])/g, ' ');
                    currentTxn.desc = cleanDesc.replace(/\s+/g, ' ').trim();

                    // 3. Keyword-First Debit/Credit Routing
                    // Checks the very first 25 characters of the description for the transaction type
                    const firstWords = currentTxn.desc.substring(0, 25).toUpperCase();
                    if (firstWords.includes('CREDIT') || firstWords.includes('REVERSAL')) {
                        currentTxn.credit = txnAmount;
                    } else if (firstWords.includes('DEBIT')) {
                        currentTxn.debit = txnAmount;
                    } else {
                        // Fallback
                        currentTxn.debit = txnAmount; 
                    }

                    worksheet.addRow(currentTxn);
                    currentTxn = null; 
                } else {
                    currentTxn.desc += ' ' + line;
                }
            }
        }

        if (currentTxn && currentTxn.date && (currentTxn.debit || currentTxn.credit)) {
            worksheet.addRow(currentTxn);
        }

        // 4. Excel Native Number Formatting
        // This forces Excel to treat the columns as real currency/accounting numbers, not just raw text.
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Skip header row
                const debitCell = row.getCell('debit');
                const creditCell = row.getCell('credit');
                const balanceCell = row.getCell('balance');
                
                if (debitCell.value) debitCell.numFmt = '#,##0.00';
                if (creditCell.value) creditCell.numFmt = '#,##0.00';
                if (balanceCell.value) balanceCell.numFmt = '#,##0.00';
            }
        });

        await workbook.xlsx.writeFile(outputPath);
        parentPort.postMessage({ success: true, outputPath });

    } catch (error) {
        console.error("Worker Exception:", error);
        parentPort.postMessage({ success: false, error: error.message });
    }
}

processPDF();
