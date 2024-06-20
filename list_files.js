const axios = require('axios');
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

async function searchFiles(accessToken, siteId, searchQuery) {
    try {
        const response = await axios.get(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/search(q='${searchQuery}')`, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            }
        });
        return response.data.value;
    } catch (error) {
        console.error('Error searching for files:', error.response ? error.response.data : error.message);
    }
}

async function listSites(accessToken) {
    try {
        const response = await axios.get('https://graph.microsoft.com/v1.0/sites?search=salesandmarketing', {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            }
        });
        return response.data.value;
    } catch (error) {
        console.error('Error listing sites:', error.response ? error.response.data : error.message);
    }
}

async function main() {
    const accessToken = await getToken();

    if (!accessToken) {
        console.error('Failed to acquire access token');
        return;
    }

    // Get the site ID for "salesandmarketing"
    const sites = await listSites(accessToken);
    const salesAndMarketingSite = sites.find(site => site.name.toLowerCase() === 'salesandmarketing');
    if (!salesAndMarketingSite) {
        console.error('Site "salesandmarketing" not found');
        return;
    }
    const siteId = salesAndMarketingSite.id;

    const sourceFileName = 'Motohaus Monthly Reporting.xlsx';
    const destinationFileName = 'June 2024.pptx';

    const sourceFiles = await searchFiles(accessToken, siteId, sourceFileName);
    const destinationFiles = await searchFiles(accessToken, siteId, destinationFileName);

    if (!sourceFiles || sourceFiles.length === 0) {
        console.error(`Source file "${sourceFileName}" not found`);
        return;
    }
    if (!destinationFiles || destinationFiles.length === 0) {
        console.error(`Destination file "${destinationFileName}" not found`);
        return;
    }

    const sourceFile = sourceFiles[0];
    const destinationFile = destinationFiles[0];

    console.log(`Source file found: ${sourceFile.name} (ID: ${sourceFile.id})`);
    console.log(`Destination file found: ${destinationFile.name} (ID: ${destinationFile.id})`);

    // Here you can implement the logic to read from the source file and update the destination file
    // For example, reading data from an Excel file and updating a PowerPoint file.
}

main();
