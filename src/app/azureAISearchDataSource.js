const { AzureKeyCredential, SearchClient } = require("@azure/search-documents");

/**
 * A data source that searches through Azure AI search using semantic search.
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

        // Perform the search using semantic search configuration
        const searchResults = await this.searchClient.search(query, {
            queryType: "semantic", // Enabling semantic search
            semanticConfiguration: this.options.semanticConfiguration, // Use the provided semantic config
            select: ["id", "metadata_spo_item_name", "metadata_spo_item_weburi", "content"], // Ensure the correct fields are selected
            queryLanguage: "en-us", // Assuming English is the query language
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
     * Formats the result string to include metadata_spo_item_weburi as a clickable link.
     */
    formatDocument(result) {
       // Set default values if the data is missing
    const webUri = result.metadata_spo_item_weburi || 'No reference available';
    const title = result.metadata_spo_item_name || 'Untitled Document';
    const content = result.content || 'No content available';

    // Log the URL for debugging purposes
    console.log("Doc URL:", webUri);

    // If the webUri is not available, you might want to handle it differently in the return statement
    if (webUri === 'No reference available') {
        return `**Title**: ${title}\n\n**Content**: ${content}\n\n(No document reference available)\n\n`;
    } else {
        // Format the result with a clickable hyperlink
        return `**Title**: ${title}\n\n**Content**: ${content}\n\n[Read more here](${webUri})\n\n`;
    }
    
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
    semanticConfiguration: "my-semantic-config-default", // Your semantic search configuration
});
