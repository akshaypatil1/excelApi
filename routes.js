const express = require('express');
const fs = require('fs');
const path = require('path');
const excelController = require('./controller/excel-controller')
const router = express.Router();

router
    .route('/api/excel').get(excelController.get);

module.exports = router;