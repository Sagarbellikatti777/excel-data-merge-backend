const express = require('express');
const multer = require('multer');
const path = require('path');
const { mergeExcelFiles } = require('../controllers/mergeController');

const router = express.Router();

const storage = multer.diskStorage({
  destination: './uploads/',
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  }
});

const upload = multer({ storage });

router.post('/', upload.array('files', 100), mergeExcelFiles);

module.exports = router;
