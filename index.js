const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

// Azure AD and MS Graph configuration
const config = {
    auth: {
        clientId: 'your_client_id_here',
        authority: 'https://login.microsoftonline.com/your_tenant_id_here',
        clientSecret: 'your_client_secret_here',
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

async function main() {
    const accessToken = await getToken();

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

    // Example: Download a specific Excel file
    const excelFileId = 'your_excel_file_id_here'; // Replace with the actual file ID
    const excelFileData = await downloadFile(accessToken, driveId, excelFileId);
    console.log('Downloaded Excel file data:', excelFileData);

    // Example: Download a specific PowerPoint file
    const pptFileId = 'your_ppt_file_id_here'; // Replace with the actual file ID
    const pptFileData = await downloadFile(accessToken, driveId, pptFileId);
    console.log('Downloaded PowerPoint file data:', pptFileData);
}

main();
