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

        worksheet.columns = [
            { header: 'Date', key: 'date', width: 15 },
            { header: 'Type', key: 'type', width: 10 },
            { header: 'Transaction Details', key: 'desc', width: 70 },
            { header: 'Debit', key: 'debit', width: 15 },
            { header: 'Credit', key: 'credit', width: 15 },
            { header: 'Balance', key: 'balance', width: 15 }
        ];
        
        worksheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
        worksheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF007BFF' } };

        // 3. Find the Opening Balance mathematically from the summary section
        let previousBalance = null;
        for (let r of rows) {
            // Looks for the summary line (e.g., "₹15,000₹130₹0₹5,120₹10,010")
            if (r.includes('₹') && r.split('₹').length >= 4) {
                const parts = r.split('₹');
                const possibleBalance = parseFloat(parts[1].replace(/,/g, ''));
                if (!isNaN(possibleBalance)) {
                    previousBalance = possibleBalance;
                    break;
                }
            }
        }

        // 4. Ultra-Defensive Parsing Logic
        let tableStarted = false;
        let currentTxn = null;
        
        const dateRegex = /^(\d{2}\s[A-Za-z]{3}\s'?\d{2})(.*)/; 
        const amountRegex = /₹([\d,]+\.?\d*)/g;

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
                    type: '',
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

                    // ==========================================
                    // NEW: Math-Based Debit/Credit Routing
                    // ==========================================
                    if (previousBalance !== null) {
                        if (currentTxn.balance > previousBalance) {
                            currentTxn.type = 'Credit';
                            currentTxn.credit = txnAmount;
                        } else if (currentTxn.balance < previousBalance) {
                            currentTxn.type = 'Debit';
                            currentTxn.debit = txnAmount;
                        } else {
                            // Fallback if balance didn't change
                            currentTxn.type = currentTxn.desc.toUpperCase().includes('DEBIT') ? 'Debit' : 'Credit';
                            if (currentTxn.type === 'Debit') currentTxn.debit = txnAmount;
                            else currentTxn.credit = txnAmount;
                        }
                    } else {
                        // Absolute fallback if opening balance wasn't found
                        currentTxn.type = currentTxn.desc.toUpperCase().includes('DEBIT') ? 'Debit' : 'Credit';
                        if (currentTxn.type === 'Debit') currentTxn.debit = txnAmount;
                        else currentTxn.credit = txnAmount;
                    }
                    
                    // Set the current balance as the previous balance for the next row
                    previousBalance = currentTxn.balance;

                    // ==========================================
                    // NEW: Smart Delimiter Cleanup
                    // ==========================================
                    let cleanDesc = currentTxn.desc;
                    // Replace '3' only if it connects letters/numbers, but leaves numbers connected to numbers alone
                    cleanDesc = cleanDesc.replace(/([a-zA-Z])3([a-zA-Z0-9])/g, '$1 $2');
                    cleanDesc = cleanDesc.replace(/([0-9])3([a-zA-Z])/g, '$1 $2');
                    cleanDesc = cleanDesc.replace(/3$/g, ''); // Clear trailing 3
                    
                    currentTxn.desc = cleanDesc.replace(/\s+/g, ' ').trim();

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

        await workbook.xlsx.writeFile(outputPath);
        parentPort.postMessage({ success: true, outputPath });

    } catch (error) {
        console.error("Worker Exception:", error);
        parentPort.postMessage({ success: false, error: error.message });
    }
}

processPDF();
