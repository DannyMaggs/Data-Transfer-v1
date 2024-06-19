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

async function listItems(accessToken, driveId, path = '', items = []) {
    try {
        const endpoint = path ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodeURIComponent(path)}:/children` : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`;
        console.log(`Listing items from endpoint: ${endpoint}`);
        const response = await axios.get(endpoint, {
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Accept': 'application/json',
            }
        });

        const newItems = response.data.value;
        items = items.concat(newItems);
        if (response.data['@odata.nextLink']) {
            const nextLink = response.data['@odata.nextLink'];
            const nextResponse = await axios.get(nextLink, {
                headers: {
                    'Authorization': `Bearer ${accessToken}`,
                    'Accept': 'application/json',
                }
            });
            items = items.concat(nextResponse.data.value);
        }
        return items;
    } catch (error) {
        console.error('Error listing items:', error.response ? error.response.data : error.message);
    }
}

async function checkFileExistence(accessToken, driveId, filePath) {
    const parts = filePath.split('/');
    let currentPath = '';
    for (let i = 0; i < parts.length; i++) {
        currentPath += (i > 0 ? '/' : '') + encodeURIComponent(parts[i]);
        const items = await listItems(accessToken, driveId, currentPath);
        if (!items) {
            console.error(`Failed to list items at path: ${currentPath}`);
            return;
        }
        const item = items.find(item => item.name === parts[i]);
        if (!item) {
            console.error(`Item not found: ${parts[i]} at path: ${currentPath}`);
            return;
        }
        if (i === parts.length - 1 && item.file) {
            console.log(`File Found - Name: ${item.name}, ID: ${item.id}, Path: ${filePath}`);
        } else if (item.folder) {
            currentPath += '/' + item.name;
        } else {
            console.error(`Item is not a folder or file: ${item.name}`);
            return;
        }
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
    if (!sites) {
        console.error('Failed to list sites');
        return;
    }
    const salesAndMarketingSite = sites.find(site => site.name.toLowerCase() === 'salesandmarketing');
    if (!salesAndMarketingSite) {
        console.error('Site "salesandmarketing" not found');
        return;
    }
    console.log(`Sales and Marketing site found: ${salesAndMarketingSite.name}`);
    const siteId = salesAndMarketingSite.id;

    // List drives and find the drive you need
    const drives = await listDrives(accessToken, siteId);
    if (!drives) {
        console.error('Failed to list drives');
        return;
    }
    if (drives.length === 0) {
        console.error('No drives found in the site');
        return;
    }
    const driveId = drives[0].id; // Assuming the first drive is the one you need
    console.log(`Drive found: ${driveId}`);

    // Check file existence directly by path
    const filePaths = [
        'Motohaus/Marketing/Motohaus%20Monthly%20Reporting.xlsx',
        'Motohaus/Marketing/Monthly%20Marketing%20Meetings/June%202024.pptx'
    ];

    for (const filePath of filePaths) {
        await checkFileExistence(accessToken, driveId, filePath);
    }

    console.log('Finished checking all file paths.');
}

main();
