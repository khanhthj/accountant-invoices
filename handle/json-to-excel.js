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

