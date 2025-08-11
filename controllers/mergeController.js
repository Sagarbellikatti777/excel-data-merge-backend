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

// const AdmZip = require('adm-zip');
// const ExcelJS = require('exceljs');

// const mergeExcelFiles = async (req, res) => {
//   try {
//     const files = req.files;

//     if (!files || files.length === 0) {
//       return res.status(400).json({ message: 'No files uploaded.' });
//     }

//     res.setHeader('Content-Disposition', 'attachment; filename=merged.xlsx');
//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

//     const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res });
//     const worksheet = workbook.addWorksheet('Merged');

//     let headerWritten = false;

//     for (const file of files) {
//       const zip = new AdmZip(file.buffer);
//       const entries = zip.getEntries();

//       for (const entry of entries) {
//         if (entry.entryName.endsWith('.xlsx')) {
//           const buffer = entry.getData();
//           const tempWorkbook = new ExcelJS.Workbook();
//           await tempWorkbook.xlsx.load(buffer);

//           const sheet = tempWorkbook.worksheets[0];
//           if (!sheet) continue;

//           sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
//             const rowValues = row.values.slice(1);
//             if (rowNumber === 1 && !headerWritten) {
//               worksheet.addRow(rowValues).commit();
//               headerWritten = true;
//             } else if (rowNumber !== 1) {
//               worksheet.addRow(rowValues).commit();
//             }
//           });
//         }
//       }
//     }

//     worksheet.commit();
//     await workbook.commit();
//   } catch (err) {
//     console.error('Error during merge:', err);
//     res.status(500).json({ message: 'Error while merging Excel files.' });
//   }
// };

// module.exports = { mergeExcelFiles };

// const AdmZip = require('adm-zip');
// const ExcelJS = require('exceljs');

// const mergeExcelFiles = async (req, res) => {
//   try {
//     const files = req.files;

//     if (!files || files.length === 0) {
//       return res.status(400).json({ message: 'No files uploaded.' });
//     }

//     const workbook = new ExcelJS.Workbook();
//     const worksheet = workbook.addWorksheet('Merged');

//     let headerWritten = false;

//     for (const file of files) {
//       const zip = new AdmZip(file.buffer);
//       const entries = zip.getEntries();

//       for (const entry of entries) {
//         if (entry.entryName.endsWith('.xlsx')) {
//           const buffer = entry.getData();
//           const tempWorkbook = new ExcelJS.Workbook();
//           await tempWorkbook.xlsx.load(buffer);

//           const sheet = tempWorkbook.worksheets[0];
//           if (!sheet) continue;

//           sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
//             const rowValues = row.values.slice(1); // remove empty first cell
//             if (rowNumber === 1 && !headerWritten) {
//               worksheet.addRow(rowValues);
//               headerWritten = true;
//             } else if (rowNumber !== 1) {
//               worksheet.addRow(rowValues);
//             }
//           });
//         }
//       }
//     }

//     // Set response headers before sending the file
//     res.setHeader('Content-Disposition', 'attachment; filename=merged.xlsx');
//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

//     // Write file to response stream
//     await workbook.xlsx.write(res);
//     res.end();

//   } catch (err) {
//     console.error('Error during merge:', err);
//     res.status(500).json({ message: 'Error while merging Excel files.' });
//   }
// };

// module.exports = { mergeExcelFiles };


// const AdmZip = require('adm-zip');
// const ExcelJS = require('exceljs');

// const mergeExcelFiles = async (req, res) => {
//   try {
//     const files = req.files;

//     if (!files || files.length === 0) {
//       return res.status(400).json({ message: 'No files uploaded.' });
//     }

//     const workbook = new ExcelJS.Workbook();
//     const worksheet = workbook.addWorksheet('Merged');
//     let headerWritten = false;
//     let totalRows = 0;
//     let totalFilesProcessed = 0;

//     // Process each file
//     for (const file of files) {
//       // Check if the file is a valid ZIP file
//       const zip = new AdmZip(file.buffer);
//       const entries = zip.getEntries();
//       let fileProcessed = false;

//       // Process each entry inside the ZIP
//       for (const entry of entries) {
//         if (entry.entryName.endsWith('.xlsx')) {
//           fileProcessed = true;
//           totalFilesProcessed++;

//           // Extract the .xlsx file and load it into ExcelJS
//           const buffer = entry.getData();
//           const tempWorkbook = new ExcelJS.Workbook();
//           await tempWorkbook.xlsx.load(buffer);

//           const sheets = tempWorkbook.worksheets;

//           // Iterate through all sheets in the workbook
//           for (const sheet of sheets) {
//             sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
//               const rowValues = row.values.slice(1); // remove empty first cell (index 0)
//               if (rowNumber === 1 && !headerWritten) {
//                 worksheet.addRow(rowValues);
//                 headerWritten = true;
//               } else if (rowNumber !== 1) {
//                 worksheet.addRow(rowValues);
//               }
//               totalRows++;
//             });
//           }
//         }
//       }

//       if (!fileProcessed) {
//         console.warn(`No .xlsx file found in the ZIP: ${file.originalname}`);
//       }
//     }

//     // If no rows were processed, return an error message
//     if (totalRows === 0) {
//       return res.status(400).json({ message: 'No valid Excel data found in uploaded ZIP files.' });
//     }

//     // Set response headers before sending the merged Excel file
//     res.setHeader('Content-Disposition', 'attachment; filename=merged.xlsx');
//     res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

//     // Write the workbook to the response stream
//     await workbook.xlsx.write(res);
//     res.end();

//     console.log(`Merged ${totalFilesProcessed} files and ${totalRows} rows of data.`);
//   } catch (err) {
//     console.error('Error during merge:', err);
//     res.status(500).json({ message: 'Error while merging Excel files.' });
//   }
// };

// module.exports = { mergeExcelFiles };


const ExcelJS = require('exceljs');
const AdmZip = require('adm-zip');

const mergeExcelFiles = async (req, res) => {
  try {
    const files = req.files;
    if (!files || files.length === 0) {
      return res.status(400).json({ message: 'No files uploaded.' });
    }

    // Set the response headers for file download
    res.setHeader('Content-Disposition', 'attachment; filename=merged.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: res });
    const worksheet = workbook.addWorksheet('Merged');

    let headerWritten = false;

    // Loop through all the uploaded files
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

          // Process each row one by one to minimize memory usage
          sheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            const rowValues = row.values.slice(1); // Remove empty first cell
            if (rowNumber === 1 && !headerWritten) {
              worksheet.addRow(rowValues);
              headerWritten = true;
            } else if (rowNumber !== 1) {
              worksheet.addRow(rowValues);
            }
          });
        }
      }
    }

    // Commit the worksheet and write the final file
    await workbook.commit();
  } catch (err) {
    console.error('Error during merge:', err);
    res.status(500).json({ message: 'Error while merging Excel files.' });
  }
};

module.exports = { mergeExcelFiles };
