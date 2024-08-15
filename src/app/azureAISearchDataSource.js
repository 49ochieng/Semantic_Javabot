const { AzureKeyCredential, SearchClient } = require("@azure/search-documents");

/**
 * A data source that searches through Azure AI search.
 */
class AzureAISearchDataSource {
    /**
     * Creates a new `AzureAISearchDataSource` instance.
     * @param {Object} options - Configuration options for the data source.
     */
    constructor(options) {
        if (!options.azureAISearchEndpoint || !options.indexName || !options.azureAISearchApiKey) {
            throw new Error("Missing required options to initialize AzureAISearchDataSource.");
        }

        this.name = options.name || 'armelysearchservice';
        this.options = options;

        // Initialize the SearchClient with the provided options
        this.searchClient = new SearchClient(
            options.azureAISearchEndpoint,
            options.indexName,
            new AzureKeyCredential(options.azureAISearchApiKey)
        );
    }

    /**
     * Renders the data source as a string of text.
     */
    async renderData(context, memory, tokenizer, maxTokens) {
        const query = memory.getValue("temp.input");
        if (!query) {
            return { output: "No input provided for the search.", length: 0, tooLong: false };
        }

        // Construct the vector search query
        const vectorQuery = {
            kind: "hnsw", // HNSW algorithm for vector search
            fields: [], // Leave fields empty if not using specific fields for the vector search
            vector: this.createDummyVector(), // Replace with actual vector data or logic
            kNearestNeighborsCount: 2, // Number of nearest neighbors to search for
        };

        // Perform the search
        const searchResults = await this.searchClient.search(query, {
            vectorQueries: [vectorQuery], // Use vectorQueries instead of vectorSearchOptions
            select: ["id", "metadata_spo_item_name", "metadata_spo_item_weburi", "content"],
        });

        // If no results, return an empty output
        if (!searchResults || searchResults.results.length === 0) {
            return { output: "No documents found matching the query.", length: 0, tooLong: false };
        }

        // Concatenate the documents string into a single document until the maximum token limit is reached
        let usedTokens = 0;
        let doc = "";
        for await (const result of searchResults.results) {
            const formattedResult = this.formatDocument(result.document);
            const tokens = tokenizer.encode(formattedResult).length;

            if (usedTokens + tokens > maxTokens) {
                break;
            }

            doc += formattedResult;
            usedTokens += tokens;
        }

        return { output: doc, length: usedTokens, tooLong: usedTokens > maxTokens };
    }

    /**
     * Dummy function to create a vector for the query.
     * Replace this with your actual vector generation logic.
     */
    createDummyVector() {
        return [0.1, 0.2, 0.3, 0.4]; // Example vector; replace with actual vector data
    }

    /**
     * Formats the result string.
     */
    formatDocument(result) {
        return `<context>${JSON.stringify(result)}</context>`;
    }
}

module.exports = {
    AzureAISearchDataSource,
};

// Initialize the data source with your configurations
const searchDataSource = new AzureAISearchDataSource({
    name: "armelysearchservice", // This name should match the name in config.json
    azureAISearchEndpoint: "https://armelysearchservice.search.windows.net",
    indexName: "sharepoint-index2",
    azureAISearchApiKey: "YOUR_AZURE_SEARCH_API_KEY",
});
