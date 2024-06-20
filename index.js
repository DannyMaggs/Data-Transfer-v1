const axios = require('axios');
const fs = require('fs');
const path = require('path');
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
    const table = worksheet.getTable(tableName);
    return table;
}

async function updatePowerPoint(pptBuffer, data) {
    const pptx = new PptxGenJS();
    await pptx.load(pptBuffer);

    const slide = pptx.getSlide(2); // Assuming we are updating slide 2
    // Add code here to update the slide with data

    const updatedBuffer = await pptx.write('arraybuffer');
    return updatedBuffer;
}

async function main() {
    const accessToken = await getToken();
    const siteId = 'YOUR_SITE_ID';

    if (!accessToken) {
        console.error('Failed to acquire access token');
        return;
    }

    const { sourceFileId, destinationFileId, sourceFileName, destinationFileName } = JSON.parse(fs.readFileSync('file_ids.json', 'utf8'));

    const sourceFileContent = await getFileContent(accessToken, siteId, sourceFileId);
    const destinationFileContent = await getFileContent(accessToken, siteId, destinationFileId);

    const excelData = await readExcelData(sourceFileContent, 'For Monthly Reports', 'Current Month (June)');
    const updatedPptBuffer = await updatePowerPoint(destinationFileContent, excelData);

    await uploadFile(accessToken, siteId, destinationFileId, updatedPptBuffer, `Updated_${destinationFileName}`);
}

main();
