const express = require('express');
const router = express.Router();
const multer = require('multer');
const upload = multer(); // In-memory storage

const { mergeExcelFiles } = require('../controllers/mergeController');  // ✅ This line must match the export

router.post('/merge', upload.array('files'), mergeExcelFiles); // ✅ Using multer middleware and correct function

module.exports = router;
