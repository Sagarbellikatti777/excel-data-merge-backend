const express = require('express');
const multer = require('multer');
const { mergeExcelFiles } = require('../controllers/mergeController');

const router = express.Router();

// Use memory storage for Multer
const storage = multer.memoryStorage();
const upload = multer({ storage });

// API Route to handle file uploads and merging
router.post('/merge', upload.array('files', 100), mergeExcelFiles);

module.exports = router;
