import fs from 'fs/promises';
import path from 'path';
import XLSX from 'xlsx'

// Hàm để đọc nội dung file JSON
export const readJSON = async (relativePath) => {
    try {
        // Resolve an absolute path from a relative path
        const absolutePath = path.resolve(relativePath);
        console.log('Resolved Path: ', absolutePath);

        // Read JSON file
        const data = await fs.readFile(absolutePath, 'utf8');

        // Analyst json file
        return JSON.parse(data);
    } catch (error) {
        console.error('Error reading or parsing JSON:', error);
        throw error;
    }
};

// Convert Json into excel
export const jsonToExcel = async (jsonData, outputPath) => {
    try {
        // Create worksheet and workbook from JSON
        const worksheet = XLSX.utils.json_to_sheet(jsonData);
        const workbook = XLSX.utils.book_new();

        // Add worksheet into workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

        // Write workbook to file
        const workbookBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        await fs.writeFile(outputPath, workbookBuffer);
        console.log('Excel file created successfully at:', outputPath);
    } catch (error) {
        console.error('Error creating Excel file:', error);
        throw error;
    }
};

// Convert Advanced
export const jsonToExcelWithFormat = async (jsonData, outputPath) => {
    try {
        const workbook = XLSX.utils.book_new();
        const worksheet = [];
        let currentRow = 1;
        const _ROW = 18;
        const _COL = 10;

        for (const item of jsonData) {

            console.log("REMAIN = ", _ROW - currentRow - item.family_count )

            if (item.relationship === "Chu ho" && _ROW - currentRow - item.family_count < 2) {
                // Tạo 9 hàng trống và thêm vào worksheet
                const emptyRows = Array(_ROW - currentRow).fill(Array(_COL).fill(null));
                for (const row of emptyRows) {
                    worksheet.push(row);
                }
            }
            // Add data to worksheet
            worksheet.push(
                createRow(item, worksheet, _COL)
            );

            //Increase 1 row when adding successfully
            currentRow++;
        }

        // Create worksheet from data
        const ws = XLSX.utils.aoa_to_sheet(worksheet);

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(workbook, ws, 'Sheet1');

        // Write workbook into file
        const workbookBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
        await fs.writeFile(outputPath, workbookBuffer);
        console.log('Excel file created successfully at:', outputPath);
    } catch (error) {
        console.error('Error creating Excel file:', error);
        throw error;
    }
};

/**
 * Create row
 * @import jsonItem = {id, family_id, name, dob, gender, relationship, family_count}
 * @import XLSX worksheet
 * @import column to add
 * @returns XLSX row object
 */
const createRow = (data, worksheet, limit) => {
    
    //Init row
    const row = [];

    //Family key to change
    const _KEY = "Chu ho";

    if (data.relationship === _KEY) worksheet.push(Array(limit).fill(null));
    row.push(data.relationship === _KEY ? data.family_id : '');
    row.push(data.relationship === _KEY ? data.name.toUpperCase() : '');
    row.push(data.relationship !== _KEY ? data.name : '');
    row.push(data.dob || '');
    row.push(data.gender || '');
    row.push(data.relationship || '');
    row.push(data.relationship === _KEY ? data.family_count : '');
    row.push(data.relationship === _KEY ? 200000 * data.family_count : '');
    row.push({ border: { top: { style: 'medium' } } });

    return row;
}