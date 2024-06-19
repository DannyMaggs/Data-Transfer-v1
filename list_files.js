const axios = require('axios');
const { ConfidentialClientApplication } = require('@azure/msal-node');

// Azure AD and MS Graph configuration
const config = {
    auth: {
        clientId: 'y3acd75e1-dbf0-4df0-88aa-2c7a4bd5ee8b', // Replace with your actual Client ID
        authority: 'https://login.microsoftonline.com/7f65e0c2-5159-471c-9af9-e57501d53752', // Replace with your actual Tenant ID
        clientSecret: 'MlC8Q~XZ_vLrsVb4E_afMEwZVKjQBk41PjIhObS0', // Replace with your actual Client Secret
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

async function searchFile(accessToken, driveId, searchQuery) {
    try {
        const response = await axios.get(`https://graph.microsoft.com/v1.0/drives/${driveId}/root/search(q='${searchQuery}')`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            }
        });
        return response.data.value;
    } catch (error) {
        console.error('Error searching for file:', error.response.data);
    }
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

    // Search for specific files
    const excelSearchQuery = 'Motohaus Monthly Reporting';
    const pptSearchQuery = 'june 2024';

    const excelItems = await searchFile(accessToken, driveId, excelSearchQuery);
    const pptItems = await searchFile(accessToken, driveId, pptSearchQuery);

    if (excelItems.length === 0 || pptItems.length === 0) {
        console.error('Required files not found');
        return;
    }

    const excelItem = excelItems[0]; // Assuming the first result is the desired file
    const pptItem = pptItems[0]; // Assuming the first result is the desired file

    console.log(`Excel Item Name: ${excelItem.name}, Item ID: ${excelItem.id}`);
    console.log(`PPT Item Name: ${pptItem.name}, Item ID: ${pptItem.id}`);
}

main();