const { ConfidentialClientApplication } = require('@azure/msal-node');
const fetch = require('node-fetch');

// Azure AD Config
const config = {
    auth: {
        clientId: "3acd75e1-dbf0-4df0-88aa-2c7a4bd5ee8b",
        authority: "https://login.microsoftonline.com/7f65e0c2-5159-471c-9af9-e57501d53752",
        clientSecret: "MlC8Q~XZ_vLrsVb4E_afMEwZVKjQBk41PjIhObS0"
    }
};

const cca = new ConfidentialClientApplication(config);

async function getToken() {
    const clientCredentialRequest = {
        scopes: ["https://graph.microsoft.com/.default"],
        skipCache: false
    };

    const response = await cca.acquireTokenByClientCredential(clientCredentialRequest);
    return response.accessToken;
}

async function fetchFiles(url, token) {
    const response = await fetch(url, {
        headers: {
            'Authorization': `Bearer ${token}`
        }
    });

    const data = await response.json();
    return data;
}

async function listAllFiles(folderId, token) {
    let files = [];
    let url = `https://graph.microsoft.com/v1.0/me/drive/items/${folderId}/children`;

    while (url) {
        const data = await fetchFiles(url, token);
        files = files.concat(data.value);

        url = data["@odata.nextLink"] || null; // Get next page URL if available
    }

    return files;
}

async function listFiles(token, searchQuery) {
    const rootId = "root"; // Replace with the root folder ID if known
    const allFiles = await listAllFiles(rootId, token);

    const matchingFiles = allFiles.filter(file => 
        file.name.toLowerCase().includes(searchQuery.toLowerCase())
    );

    if (matchingFiles.length > 0) {
        console.log("Matching files:");
        matchingFiles.forEach(file => {
            console.log(`Item Name: ${file.name}, Item ID: ${file.id}`);
        });
    } else {
        console.log("No matching files found");
    }
}

async function main() {
    try {
        const token = await getToken();
        const searchQuery = process.argv[2]; // Pass the search query as a command-line argument

        if (!searchQuery) {
            console.log("Please provide a search query");
            return;
        }

        await listFiles(token, searchQuery);
    } catch (error) {
        console.error("Error:", error);
    }
}

main();
