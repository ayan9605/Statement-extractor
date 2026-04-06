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
        const rows = rawText.split('\n').filter(line => line.trim() !== '');

        // 2. Build Excel File Structure
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Transactions');

        // Clean headers with separate Debit/Credit columns
        worksheet.columns = [
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Type', key: 'type', width: 10 },
            { header: 'Transaction Details', key: 'desc', width: 70 },
            { header: 'Debit (₹)', key: 'debit', width: 15 },
            { header: 'Credit (₹)', key: 'credit', width: 15 },
            { header: 'Balance (₹)', key: 'balance', width: 15 }
        ];
        
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF007BFF' } };

        // 3. The Custom Slice Bank Parsing Logic
        let tableStarted = false;
        let currentTxn = null;
        
        // Regex to catch dates like "01 Apr '26"
        const dateRegex = /^(\d{2}\s[A-Za-z]{3}\s'\d{2})(.*)/; 

        for (let i = 0; i < rows.length; i++) {
            let line = rows[i].trim();

            // Ignore everything until the actual transaction table starts
            if (line.includes('DATEDETAILSREF NO.AMOUNTBALANCE')) {
                tableStarted = true;
                continue;
            }
            if (!tableStarted) continue;
            
            // Stop parsing when we hit the footer
            if (line.includes('Need help?')) break; 

            // Check if line starts with a Date
            const dateMatch = line.match(dateRegex);

            if (dateMatch) {
                // Start a new transaction block
                currentTxn = {
                    date: dateMatch[1],
                    type: '',
                    desc: dateMatch[2], // The rest of the line after the date
                    debit: '',
                    credit: '',
                    balance: ''
                };
            } 
            // If it doesn't start with a date, we are inside a transaction block building the data
            else if (currentTxn) {
                if (line.includes('₹')) {
                    // This is the final line of the block containing the amounts
                    // Example line: "8042609112574134₹10₹15,010"
                    const parts = line.split('₹');
                    
                    // Add remaining reference numbers to description
                    currentTxn.desc += ' ' + parts[0].trim(); 
                    
                    // Convert amounts to actual numbers for Excel
                    const amount = parseFloat(parts[1].replace(/,/g, ''));
                    currentTxn.balance = parseFloat(parts[2].replace(/,/g, ''));

                    // Determine if it's a Debit or Credit based on the description
                    const descUpper = currentTxn.desc.toUpperCase();
                    if (descUpper.includes('DEBIT')) {
                        currentTxn.type = 'Debit';
                        currentTxn.debit = amount;
                    } else if (descUpper.includes('CREDIT') || descUpper.includes('REVERSAL')) {
                        currentTxn.type = 'Credit';
                        currentTxn.credit = amount;
                    }

                    // Clean up the weird '3' delimiter issue caused by the PDF extraction
                    currentTxn.desc = currentTxn.desc.replace(/3/g, ' ').replace(/\s+/g, ' ').trim();

                    // Push the completed transaction to Excel and reset
                    worksheet.addRow(currentTxn);
                    currentTxn = null;
                } else {
                    // It's just a multi-line description, keep attaching it
                    currentTxn.desc += ' ' + line;
                }
            }
        }

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
