const { AzureKeyCredential, SearchClient, SearchIndexClient } = require("@azure/search-documents");
const { createIndexIfNotExists, delay, upsertDocuments, getEmbeddingVector } = require("./utils");
const path = require("path");
const fs = require("fs");

/**
 *  Main function that creates the index and upserts the documents.
 */
async function main() {
    const index = "sharepoint-index2";  // Your index name

    if (
        !process.env.SECRET_AZURE_SEARCH_KEY ||
        !process.env.AZURE_SEARCH_ENDPOINT ||
        !process.env.SECRET_AZURE_OPENAI_API_KEY ||
        !process.env.AZURE_OPENAI_ENDPOINT ||
        !process.env.AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME
    ) {
        throw new Error(
            "Missing environment variables - please check that SECRET_AZURE_SEARCH_KEY, AZURE_SEARCH_ENDPOINT, SECRET_AZURE_OPENAI_API_KEY, AZURE_OPENAI_ENDPOINT, and AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME are set."
        );
    }

    const searchApiKey = process.env.SECRET_AZURE_SEARCH_KEY;
    const searchApiEndpoint = process.env.AZURE_SEARCH_ENDPOINT;
    const credentials = new AzureKeyCredential(searchApiKey);

    const searchIndexClient = new SearchIndexClient(searchApiEndpoint, credentials);
    await createIndexIfNotExists(searchIndexClient, index);  // Use your existing index
    await delay(5000);  // Wait 5 seconds for the index to be created

    const searchClient = new SearchClient(searchApiEndpoint, index, credentials);

    const filePath = path.join(__dirname, "./data");
    const files = fs.readdirSync(filePath);
    const data = [];
    for (let i = 1; i <= files.length; i++) {
        const content = fs.readFileSync(path.join(filePath, files[i - 1]), "utf-8");
        data.push({
            id: i + "",  // ID field from your index
            metadata_spo_item_name: files[i - 1],
            metadata_spo_item_weburi: `https://example.com/${files[i - 1]}`, // Example web URI, adjust as needed
            content: content,
            metadata_spo_item_title: files[i - 1], // Title field
        });
    }
    await upsertDocuments(searchClient, data);
}

main();

module.exports = main;
