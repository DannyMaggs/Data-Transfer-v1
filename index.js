const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const ExcelJS = require('exceljs');
const pptxgen = require('pptxgenjs'); // Use the appropriate PowerPoint library
const fs = require('fs');
const path = require('path');

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
        return authResult.accessToken;
    } catch (error) {
        console.error('Error acquiring token:', error);
    }
}

async function listSites(accessToken) {
    try {
        const response = await axios.get('https://graph.microsoft.com/v1.0/sites?search=*', {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            }
        });
        return response.data.value;
    } catch (error) {
        console.error('Error listing sites:', error.response.data);
    }
}

async function listDrives(accessToken, siteId) {
    try {
        const response = await axios.get(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            }
        });
        return response.data.value;
    } catch (error) {
        console.error('Error listing drives:', error.response.data);
    }
}

async function listItems(accessToken, driveId) {
    try {
        const response = await axios.get(`https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            }
        });
        return response.data.value;
    } catch (error) {
        console.error('Error listing items:', error.response.data);
    }
}

async function downloadFile(accessToken, driveId, fileId) {
    try {
        const response = await axios.get(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            },
            responseType: 'arraybuffer',
        });
        return response.data;
    } catch (error) {
        console.error('Error downloading file:', error.response.data);
    }
}

async function processExcelData(excelBuffer) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(excelBuffer);
    const worksheet = workbook.getWorksheet('For_Monthly_Reports'); // Replace with your actual sheet name

    let tableData = [];
    worksheet.eachRow((row, rowNumber) => {
        let rowData = [];
        row.eachCell((cell, colNumber) => {
            rowData.push(cell.value);
        });
        tableData.push(rowData);
    });
    return tableData;
}

async function updatePowerPoint(pptBuffer, tableData) {
    const prs = new pptxgen();
    prs.load(pptBuffer);

    const slide = prs.addSlide();
    const rows = tableData.length;
    const cols = tableData[0].length;

    const table = slide.addTable(tableData, {
        x: 1,
        y: 1,
        w: 8,
        h: 5
    });

    const newPptBuffer = prs.saveToBuffer();
    return newPptBuffer;
}

async function main() {
    const accessToken = await getToken();

    if (!accessToken) {
        console.error('Failed to acquire access token');
        return;
    }

    // List sites and find the 'salesandmarketing' site
    const sites = await listSites(accessToken);
    const salesAndMarketingSite = sites.find(site => site.name === 'salesandmarketing');
    if (!salesAndMarketingSite) {
        console.error('Site "salesandmarketing" not found');
        return;
    }
    const siteId = salesAndMarketingSite.id;

    // List drives and find the drive you need
    const drives = await listDrives(accessToken, siteId);
    if (drives.length === 0) {
        console.error('No drives found in the site');
        return;
    }
    const driveId = drives[0].id; // Assuming the first drive is the one you need

    // List items in the drive
    const items = await listItems(accessToken, driveId);
    items.forEach(item => {
        console.log(`Item Name: ${item.name}, Item ID: ${item.id}`);
    });

    // Replace these with actual filenames
    const excelFilename = 'Motohaus Monthly Reporting';
    const pptFilename = 'june 2024'; // Replace with your actual PPT filename

    const excelItem = items.find(item => item.name === excelFilename);
    const pptItem = items.find(item => item.name === pptFilename);

    if (!excelItem || !pptItem) {
        console.error('Required files not found');
        return;
    }

    // Download the Excel file
    const excelBuffer = await downloadFile(accessToken, driveId, excelItem.id);
    const tableData = await processExcelData(excelBuffer);

    // Download the PowerPoint file
    const pptBuffer = await downloadFile(accessToken, driveId, pptItem.id);
    const newPptBuffer = await updatePowerPoint(pptBuffer, tableData);

    // Upload the updated PowerPoint file back to SharePoint
    const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${pptItem.id}/content`;
    const uploadResponse = await axios.put(uploadUrl, newPptBuffer, {
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        }
    });

    if (uploadResponse.status === 200) {
        console.log('PowerPoint presentation updated successfully.');
    } else {
        console.error('Failed to upload the PowerPoint file:', uploadResponse.status, uploadResponse.statusText);
    }
}

main();
