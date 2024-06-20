const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const ExcelJS = require('exceljs');
const PptxGenJS = require('pptxgenjs');

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

async function listSites(accessToken) {
    try {
        const response = await axios.get('https://graph.microsoft.com/v1.0/sites?search=*', {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            }
        });
        console.log('Sites listed successfully');
        return response.data.value;
    } catch (error) {
        console.error('Error listing sites:', error.response ? error.response.data : error.message);
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
        console.log(`Drives listed successfully for site ID: ${siteId}`);
        return response.data.value;
    } catch (error) {
        console.error('Error listing drives:', error.response ? error.response.data : error.message);
    }
}

async function searchFileRecursive(accessToken, driveId, path, targetFileName) {
    try {
        const response = await axios.get(`https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${path}:/children`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            }
        });

        for (const item of response.data.value) {
            if (item.folder) {
                const found = await searchFileRecursive(accessToken, driveId, `${path}/${item.name}`, targetFileName);
                if (found) return found;
            } else {
                if (item.name.toLowerCase() === targetFileName.toLowerCase()) {
                    console.log(`File found: ${item.name}`);
                    return item.id;
                }
            }
        }
    } catch (error) {
        console.error('Error searching for files:', error.response ? error.response.data : error.message);
    }
    return null;
}

async function downloadFile(accessToken, driveId, fileId) {
    try {
        const response = await axios.get(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${fileId}/content`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            },
            responseType: 'arraybuffer'
        });
        console.log(`File downloaded successfully: ${fileId}`);
        return response.data;
    } catch (error) {
        console.error('Error downloading file:', error.response ? error.response.data : error.message);
    }
}

async function extractDataFromExcel(buffer) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.getWorksheet(1);
    const data = [];

    worksheet.eachRow((row, rowNumber) => {
        data.push(row.values);
    });

    return data;
}

async function updatePowerPoint(buffer, data) {
    const pptx = new PptxGenJS();
    const slide = pptx.addSlide();

    data.forEach((row, index) => {
        slide.addText(row.join(' '), { x: 0.5, y: 0.5 + index * 0.5, fontSize: 12 });
    });

    const pptBuffer = await pptx.write('arraybuffer');
    return pptBuffer;
}

async function uploadFile(accessToken, driveId, parentId, fileName, fileBuffer) {
    try {
        const response = await axios.put(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${parentId}:/${fileName}:/content`, fileBuffer, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            }
        });
        console.log(`File uploaded successfully: ${fileName}`);
        return response.data;
    } catch (error) {
        console.error('Error uploading file:', error.response ? error.response.data : error.message);
    }
}

async function main() {
    const accessToken = await getToken();
    if (!accessToken) {
        console.error('Failed to acquire access token');
        return;
    }

    const sites = await listSites(accessToken);
    const salesAndMarketingSite = sites.find(site => site.name.toLowerCase() === 'salesandmarketing');
    if (!salesAndMarketingSite) {
        console.error('Site "salesandmarketing" not found');
        return;
    }
    const siteId = salesAndMarketingSite.id;

    const drives = await listDrives(accessToken, siteId);
    if (!drives) {
        console.error('Failed to list drives');
        return;
    }
    const driveId = drives[0].id;

    const sourceFileName = 'Motohaus Monthly Reporting.xlsx';
    const targetFileName = 'June 2024.pptx';

    const sourceFileId = await searchFileRecursive(accessToken, driveId, '', sourceFileName);
    const targetFileId = await searchFileRecursive(accessToken, driveId, '', targetFileName);

    if (!sourceFileId) {
        console.error(`Source file "${sourceFileName}" not found`);
        return;
    }
    if (!targetFileId) {
        console.error(`Target file "${targetFileName}" not found`);
        return;
    }

    const sourceFileBuffer = await downloadFile(accessToken, driveId, sourceFileId);
    const targetFileBuffer = await downloadFile(accessToken, driveId, targetFileId);

    const data = await extractDataFromExcel(sourceFileBuffer);
    const updatedPptBuffer = await updatePowerPoint(targetFileBuffer, data);

    const parentId = targetFileId.split('!')[1];
    await uploadFile(accessToken, driveId, parentId, targetFileName, updatedPptBuffer);
}

main();
