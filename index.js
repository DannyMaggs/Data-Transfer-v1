const axios = require('axios');
const fs = require('fs');
const ExcelJS = require('exceljs');
const PptxGenJS = require('pptxgenjs');
const { ConfidentialClientApplication } = require('@azure/msal-node');

// Azure AD and MS Graph configuration
const config = {
    auth: {
        clientId: '3acd75e1-dbf0-4df0-88aa-2c7a4bd5ee8b',
        authority: 'https://login.microsoftonline.com/7f65e0c2-5159-471c-9af9-e57501d53752',
        clientSecret: 'MlC8Q~XZ_vLrsVb4E_afMEwZVKjQBk41PjIhObS0',
    }
};

// MSAL client application
const cca = new ConfidentialClientApplication(config);

// Authentication parameters
const authParams = {
    scopes: ['https://graph.microsoft.com/.default'],
};

async function getToken() {
    try {
        const authResult = await cca.acquireTokenByClientCredential(authParams);
        console.log('Access token acquired successfully');
        return authResult.accessToken;
    } catch (error) {
        console.error('Error acquiring token:', error);
    }
}

async function getFileContent(accessToken, siteId, itemId) {
    try {
        const response = await axios.get(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/content`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            },
            responseType: 'arraybuffer'
        });
        return response.data;
    } catch (error) {
        console.error(`Error retrieving file content: ${error.message}`);
    }
}

async function uploadFile(accessToken, siteId, itemId, fileData, fileName) {
    try {
        const response = await axios.put(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/items/${itemId}/content`, fileData, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/octet-stream'
            }
        });
        console.log(`File uploaded: ${fileName}`);
    } catch (error) {
        console.error(`Error uploading file: ${error.message}`);
    }
}

async function readExcelData(excelBuffer, sheetName, tableName) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(excelBuffer);
    const worksheet = workbook.getWorksheet(sheetName);

    if (!worksheet) {
        console.error(`Worksheet "${sheetName}" not found`);
        return [];
    }

    console.log(`Worksheet "${sheetName}" found. Checking tables...`);

    const tables = worksheet.model.tables;
    if (tables) {
        tables.forEach(t => console.log(t.name));
    } else {
        console.log("No tables found in worksheet");
    }

    const table = tables.find(t => t.name === tableName);
    const data = [];

    if (table) {
        const tableRange = worksheet.getRanges(table.ref)[0];
        tableRange.eachRow((row, rowNumber) => {
            data.push(row.values);
        });
    } else {
        console.error(`Table "${tableName}" not found`);
    }

    return data;
}

async function updatePowerPoint(pptBuffer, data) {
    const pptx = new PptxGenJS();
    const pres = pptx.load(pptBuffer);

    const slide = pres.getSlide(6); // Assuming we are updating slide 6
    const table = slide.getTable(0); // Assuming it's the first table on the slide

    if (table) {
        table.rows = data;
    } else {
        console.error(`Table not found on slide 6`);
    }

    const updatedBuffer = await pptx.write('arraybuffer');
    return updatedBuffer;
}

async function main() {
    const accessToken = await getToken();
    const siteId = 'motohaus.sharepoint.com,2c54175f-ca53-4ca4-8eab-1530b7f64072,07a87623-8134-4e67-b860-0ff98346cfc2';

    if (!accessToken) {
        console.error('Failed to acquire access token');
        return;
    }

    const { sourceFileId, destinationFileId } = JSON.parse(fs.readFileSync('file_ids.json', 'utf8'));

    const sourceFileContent = await getFileContent(accessToken, siteId, sourceFileId);
    const destinationFileContent = await getFileContent(accessToken, siteId, destinationFileId);

    const excelData = await readExcelData(sourceFileContent, 'For Monthly Reports', 'Table_02');
    if (excelData.length === 0) {
        console.error('No data found in the Excel table');
        return;
    }

    const updatedPptBuffer = await updatePowerPoint(destinationFileContent, excelData);

    await uploadFile(accessToken, siteId, destinationFileId, updatedPptBuffer, 'Updated_' + destinationFileId);
}

main();
