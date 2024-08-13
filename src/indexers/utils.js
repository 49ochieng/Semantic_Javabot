const { KnownAnalyzerNames } = require("@azure/search-documents");
const { OpenAIEmbeddings } = require("@microsoft/teams-ai");

/**
 * A wrapper for setTimeout that resolves a promise after timeInMs milliseconds.
 */
function delay(timeInMs) {
    return new Promise((resolve) => setTimeout(resolve, timeInMs));
}

/**
 * Deletes the index with the given name
 */
function deleteIndex(client, name) {
    return client.deleteIndex(name);
}

/**
 * Adds or updates the given documents in the index
 */
async function upsertDocuments(client, documents) {
    return await client.mergeOrUploadDocuments(documents);
}

/**
 * Creates the index with the given name if it doesn't exist (for new indices)
 */
async function createIndexIfNotExists(client, name) {
    try {
        const index = await client.getIndex(name);
        if (index) {
            console.log(`Index ${name} already exists. Skipping creation.`);
            return;
        }
    } catch (error) {
        if (error.statusCode === 404) {
            console.log(`Index ${name} does not exist. Creating it.`);
            const MyDocumentIndex = {
                name,
                fields: [
                    { name: "id", type: "Edm.String", key: true, searchable: false },
                    { name: "metadata_spo_item_name", type: "Edm.String", searchable: true },
                    { name: "metadata_spo_item_weburi", type: "Edm.String", searchable: true },
                    { name: "content", type: "Edm.String", searchable: true },
                    { name: "metadata_spo_item_title", type: "Edm.String", searchable: true },
                ],
                corsOptions: {
                    allowedOrigins: ["*"],
                },
                suggesters: [
                    {
                        name: "my-suggester",
                        sourceFields: ["metadata_spo_item_name", "content"],
                    },
                ],
                semantic: {
                    defaultConfiguration: "my-semantic-config-default",
                    configurations: [
                        {
                            name: "my-semantic-config-default",
                            prioritizedFields: {
                                titleField: { fieldName: "metadata_spo_item_title" },
                                prioritizedContentFields: [
                                    { fieldName: "content" },
                                    { fieldName: "metadata_spo_item_name" },
                                ],
                            },
                        },
                    ],
                },
            };

            await client.createOrUpdateIndex(MyDocumentIndex);
        } else {
            throw error;
        }
    }
}

/**
 * Generate the embedding vector
 */
async function getEmbeddingVector(text) {
    const embeddings = new OpenAIEmbeddings({
        azureApiKey: process.env.SECRET_AZURE_OPENAI_API_KEY,
        azureEndpoint: process.env.AZURE_OPENAI_ENDPOINT,
        azureDeployment: process.env.AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME,
    });

    const result = await embeddings.createEmbeddings(process.env.AZURE_OPENAI_EMBEDDING_DEPLOYMENT_NAME, text);

    if (result.status !== "success" || !result.output) {
        throw new Error(`Failed to generate embeddings for description: ${text}`);
    }

    return result.output[0];
}

module.exports = {
    deleteIndex,
    createIndexIfNotExists,
    delay,
    upsertDocuments,
    getEmbeddingVector,
};
