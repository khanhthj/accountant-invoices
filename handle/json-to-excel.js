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
        let peoples = 0;
        let supplies = 0;

        for (let i = 0; i< jsonData.length; i++) {

            const item = jsonData[i];
            //Caculate total peoples and supplies

            console.log("test ==== " , _ROW - currentRow - 1);

            //If new family_count is more than remaining rows
            if (item.relationship === "Chu ho" && _ROW - currentRow - item.family_count < 2) {
                // Create blank row and add to worksheet
                const emptyRows = Array(_ROW - currentRow - 1).fill(Array(_COL).fill(null));
                for (const row of emptyRows) {
                    worksheet.push(row);
                }

                //Create page summary
                worksheet.push(
                    createSumPageRow(peoples, supplies)
                );
                peoples = 0;
                supplies = 0;
                currentRow = 1;

            }
            // Add data to worksheet
            worksheet.push(
                createRow(item, worksheet, _COL)
            );

            //Increase 1 row when adding successfully
            currentRow++;

            //If valid family member
            if(item.family_count){
                peoples ++;
                supplies +=200000;
            }
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
    row.push(data.relationship === _KEY ? {v: data.family_count} : '');
    row.push(data.relationship === _KEY ? {v: 200000 * data.family_count} : '');
    row.push({ border: { top: { style: 'medium' } } });

    return row;
}

/**
 * Add "Cộng trang"
 * @returns XLSX row object
 */
const createSumPageRow = (peoples, supplies) => {
    const row = [];
    const _TEXT = "CỘNG TRANG";

    row.push('');
    row.push(_TEXT);
    for (let i = 0; i < 4; i++) {
        row.push('')
    };
    row.push({t: 'n', v: peoples});
    row.push({t: 'n', v: supplies});
    row.push('');
    row.push('');

    return row;
}