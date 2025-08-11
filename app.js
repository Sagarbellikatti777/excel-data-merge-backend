const express = require('express');
const multer = require('multer');
const AdmZip = require('adm-zip');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const cors = require('cors');  // Add this line

const app = express();

// Enable CORS for all routes
app.use(cors());

// Temp folder for uploads
const uploadFolder = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadFolder)) fs.mkdirSync(uploadFolder);

// Multer disk storage
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, uploadFolder),
    filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
  }),
  limits: { fileSize: 1024 * 1024 * 1024 } // 1 GB max per file
});

app.post('/merge', upload.array('files', 100), async (req, res) => {
  try {
    if (!req.files || req.files.length === 0) {
      return res.status(400).send('No files uploaded');
    }

    // Set Excel download headers
    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader('Content-Disposition', 'attachment; filename="merged.xlsx"');

    // Create streaming Excel workbook
    const workbookWriter = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res });
    const worksheet = workbookWriter.addWorksheet('Merged Data');

    let headerRow = null; // store header from first file

    const processExcelFile = async (buffer) => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(buffer);
      const ws = wb.worksheets[0];
      if (!ws) return;

      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        const rowArr = [];
        for (let i = 1; i < row.values.length; i++) {
          rowArr.push(row.values[i]);
        }

        if (!headerRow) {
          headerRow = rowArr;
          worksheet.addRow(headerRow).commit();
          return;
        }

        if (
          rowNumber === 1 &&
          JSON.stringify(rowArr) === JSON.stringify(headerRow)
        ) {
          return; // skip duplicate header
        }

        worksheet.addRow(rowArr).commit();
      });
    };

    // Process each uploaded ZIP
    for (const file of req.files) {
      try {
        const zip = new AdmZip(file.path);
        const entries = zip.getEntries();
        const excelEntry = entries.find(e => /\.(xlsx)$/i.test(e.entryName));

        if (!excelEntry) {
          // console.log(No Excel found in: ${file.originalname});
          continue;
        }

        const excelBuffer = excelEntry.getData();
        await processExcelFile(excelBuffer);

      } catch (err) {
        // console.error(Error processing zip: ${file.originalname}, err);
      } finally {
        // Remove uploaded ZIP to free space
        fs.unlinkSync(file.path);
      }
    }

    await workbookWriter.commit();
    console.log('Merge complete');

  } catch (err) {
    console.error(err);
    if (!res.headersSent) res.status(500).send('Server error');
  }
});

app.listen(4000, () => console.log('Server running on port 4000'));
