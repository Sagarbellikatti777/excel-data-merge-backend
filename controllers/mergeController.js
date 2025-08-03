const XLSX = require('xlsx');
const AdmZip = require('adm-zip');

const mergeExcelFiles = async (req, res) => {
  try {
    const files = req.files;

    if (!files || files.length === 0) {
      return res.status(400).json({ message: 'No files uploaded.' });
    }

    let mergedData = [];

    for (const file of files) {
      const zip = new AdmZip(file.buffer);
      const zipEntries = zip.getEntries();

      for (const entry of zipEntries) {
        if (entry.entryName.endsWith('.xlsx')) {
          const workbook = XLSX.read(entry.getData(), { type: 'buffer' });
          const sheetName = workbook.SheetNames[0];
          const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });
          mergedData = mergedData.concat(data);
        }
      }
    }

    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.json_to_sheet(mergedData);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Merged');

    const outputBuffer = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'buffer' });

    res.setHeader('Content-Disposition', 'attachment; filename=merged.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(outputBuffer);
  } catch (err) {
    console.error('Merge error:', err);
    res.status(500).json({ message: 'Error during merging files.' });
  }
};

module.exports = { mergeExcelFiles }; // ✅ Important
