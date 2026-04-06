const express = require('express');
const multer = require('multer');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { Worker } = require('worker_threads');

const app = express();
app.use(cors());

// Secure, temporary storage for massive files
const upload = multer({ 
    dest: 'uploads/',
    limits: { fileSize: 100 * 1024 * 1024 } // Up to 100MB
});

// ==========================================
// NEW: Serve the HTML Frontend UI
// ==========================================
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// ==========================================
// API Endpoint: Process the PDF
// ==========================================
app.post('/api/convert', upload.single('statement'), (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: 'No PDF uploaded.' });
    }

    const pdfPath = req.file.path;
    const outputPath = path.join(__dirname, `uploads/statement_${Date.now()}.xlsx`);

    // Spin up a separate worker thread for the heavy CPU processing
    const worker = new Worker(path.join(__dirname, 'pdfWorker.js'), {
        workerData: { pdfPath, outputPath }
    });

    worker.on('message', (message) => {
        if (message.success) {
            res.download(message.outputPath, 'Parsed_Statement.xlsx', () => {
                // Cleanup both files after successful download
                if (fs.existsSync(pdfPath)) fs.unlinkSync(pdfPath);
                if (fs.existsSync(message.outputPath)) fs.unlinkSync(message.outputPath);
            });
        } else {
            res.status(500).json({ error: 'Conversion failed during processing.' });
        }
    });

    worker.on('error', (err) => {
        console.error("Worker Error:", err);
        res.status(500).json({ error: 'Critical error processing PDF.' });
    });
});

app.listen(3000, () => console.log('🚀 High-Performance Server running on port 3000'));
