import express from 'express'

const app = express();
//const _PATCH = 'D:/KHANH-THI/accountant-invoices/data/thoian.json'
const _PATCH = 'D:/KHANH-THI/accountant-invoices/data/test.json'

import { readJSON, jsonToExcel } from './handle/json-to-excel.js';

// Route
app.get('/', async (req, res) => {
    try {
        // Read data from JSON file
        const jsonData = await readJSON(_PATCH);
        
        // Convert data to Excel file
        await jsonToExcel(jsonData, 'output.xlsx');
        res.json("Create file successfully!")
    } catch (error) {
        res.json("Unable to create file!")
    }
    //res.json(await readJSON(_PATCH))
});

// Khởi động server
const port = 3000;
app.listen(port, () => {
    console.log(`Server is up on port ${port}`);
});
