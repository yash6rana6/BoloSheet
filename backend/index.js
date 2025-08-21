const express = require('express');
const cors = require('cors');
const ExcelJS = require('exceljs');
const fs = require('fs').promises;
const lockfile = require('proper-lockfile');
const path = require('path');

const app = express();

app.use(cors());
app.use(express.json());

app.post('/save-entry', async (req, res) => {
  console.log('Received data:', req.body);
  const { fields, entries } = req.body;
  const workbook = new ExcelJS.Workbook();
  const filePath = path.join(__dirname, 'data.xlsx');
  console.log('Saving to:', filePath);

  try {
    const release = await lockfile.lock(filePath, { retries: 5 });
    try {
      const fileExists = await fs.access(filePath).then(() => true).catch(() => false);
      if (fileExists) {
        await workbook.xlsx.readFile(filePath);
        console.log('Existing file loaded');
      } else {
        console.log('Creating new Excel file');
        const worksheet = workbook.addWorksheet('Sheet1');
        worksheet.addRow(fields);
      }

      const worksheet = workbook.getWorksheet(1) || workbook.addWorksheet('Sheet1');
      entries.forEach(entry => {
        const row = fields.map(field => entry[field] || ''); 
        console.log('Row added:', row);
        worksheet.addRow(row);
      });

      await workbook.xlsx.writeFile(filePath);
      console.log('File saved successfully');
      res.json({ message: 'Entries saved to Excel!' });
    } finally {
      await release();
    }
  } catch (error) {
    console.error('Error saving to Excel:', error.message, error.stack);
    res.status(500).json({ error: 'Failed to save entries', details: error.message });
  }
});

app.get('/download-excel', (req, res) => {
  const filePath = path.join(__dirname, 'data.xlsx');
  res.download(filePath, 'BoloSheet.xlsx', (err) => {
    if (err) {
      console.error('Download error:', err);
      res.status(500).send('Failed to download');
    }
  });
});

app.get('/preview-excel', async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  const filePath = path.join(__dirname, 'data.xlsx');
  try {
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(1);
    if (!worksheet) throw new Error('Worksheet not found');

    const data = [];
    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) {
        data.push(row.values.slice(1)); 
      } else {
        data.push(row.values.slice(1)); 
      }
    });

    res.json({ data });
  } catch (error) {
    console.error('Error previewing Excel:', error.message);
    res.status(500).json({ error: 'Failed to preview Excel', details: error.message });
  }
});

app.listen(5000, () => console.log('Server running on port 5000'));