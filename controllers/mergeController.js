// const XLSX = require('xlsx');
// const AdmZip = require('adm-zip');

// const mergeExcelFiles = async (req, res) => {
//   try {
//     const files = req.files;

//     if (!files || files.length === 0) {
//       return res.status(400).json({ message: 'No files uploaded.' });
//     }

//     let mergedData = [];

//     for (const file of files) {
//       const zip = new AdmZip(file.buffer);
//       const zipEntries = zip.getEntries();

//       for (const entry of zipEntries) {
//         if (entry.entryName.endsWith('.xlsx')) {
//           const workbook = XLSX.read(entry.getData(), { type: 'buffer' });
//           const sheetName = workbook.SheetNames[0];
//           const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
//           mergedData = mergedData.concat(data);
//         }
//       }
//     }

//     const newWorkbook = XLSX.utils.book_new();
//     const newWorksheet = XLSX.utils.json_to_sheet(mergedData);
//     XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Merged');

//     const outputBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'buffer' });

//     res.setHeader('Content-Disposition', 'attachment; filename=merged.xlsx');
//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//     res.send(outputBuffer);
//   } catch (err) {
//     console.error('Merge error:', err);
//     res.status(500).json({ message: 'Error during merging files.' });
//   }
// };

// module.exports = { mergeExcelFiles }; // ✅ Important

// const XLSX = require('xlsx');
// const AdmZip = require('adm-zip');

// const mergeExcelFiles = async (req, res) => {
//   try {
//     const files = req.files;

//     if (!files || files.length === 0) {
//       return res.status(400).json({ message: 'No files uploaded.' });
//     }

//     let mergedData = [];
//     let headerAdded = false;

//     for (const file of files) {
//       const zip = new AdmZip(file.buffer);
//       const zipEntries = zip.getEntries();

//       for (const entry of zipEntries) {
//         if (entry.entryName.endsWith('.xlsx')) {
//           const workbook = XLSX.read(entry.getData(), { type: 'buffer' });
//           const sheetName = workbook.SheetNames[0];
//           const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '', header: 1 });

//           // Skip empty sheets
//           if (data.length === 0) continue;

//           if (!headerAdded) {
//             mergedData.push(...data); // include header
//             headerAdded = true;
//           } else {
//             mergedData.push(...data.slice(1)); // skip header
//           }
//         }
//       }
//     }

//     if (mergedData.length === 0) {
//       return res.status(400).json({ message: 'No Excel data found in ZIPs.' });
//     }

//     const worksheet = XLSX.utils.aoa_to_sheet(mergedData);
//     const workbook = XLSX.utils.book_new();
//     XLSX.utils.book_append_sheet(workbook, worksheet, 'Merged');

//     const outputBuffer = XLSX.write(workbook, {
//       bookType: 'xlsx',
//       type: 'buffer',
//       compression: true, // Helps reduce size
//     });

//     res.setHeader('Content-Disposition', 'attachment; filename=merged.xlsx');
//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
//     res.send(outputBuffer);
//   } catch (err) {
//     console.error('Merge error:', err);
//     res.status(500).json({ message: 'Error during merging files.' });
//   }
// };

// module.exports = { mergeExcelFiles };

const AdmZip = require('adm-zip');
const ExcelJS = require('exceljs');

const mergeExcelFiles = async (req, res) => {
  try {
    const files = req.files;

    if (!files || files.length === 0) {
      return res.status(400).json({ message: 'No files uploaded.' });
    }

    res.setHeader('Content-Disposition', 'attachment; filename=merged.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res });
    const worksheet = workbook.addWorksheet('Merged');

    let headerWritten = false;

    for (const file of files) {
      const zip = new AdmZip(file.buffer);
      const entries = zip.getEntries();

      for (const entry of entries) {
        if (entry.entryName.endsWith('.xlsx')) {
          const buffer = entry.getData();
          const tempWorkbook = new ExcelJS.Workbook();
          await tempWorkbook.xlsx.load(buffer);

          const sheet = tempWorkbook.worksheets[0];
          if (!sheet) continue;

          sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const rowValues = row.values.slice(1);
            if (rowNumber === 1 && !headerWritten) {
              worksheet.addRow(rowValues).commit();
              headerWritten = true;
            } else if (rowNumber !== 1) {
              worksheet.addRow(rowValues).commit();
            }
          });
        }
      }
    }

    worksheet.commit();
    await workbook.commit();
  } catch (err) {
    console.error('Error during merge:', err);
    res.status(500).json({ message: 'Error while merging Excel files.' });
  }
};

module.exports = { mergeExcelFiles };