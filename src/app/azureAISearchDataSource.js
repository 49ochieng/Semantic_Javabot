const { OpenAIEmbeddings } = require("@microsoft/teams-ai");
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
        // Ensure all required options are provided
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
            return { output: "", length: 0, tooLong: false };
        }

        const selectedFields = [
            "id",
            "metadata_spo_item_name",
            "metadata_spo_item_weburi",
            "content",
        ];

        // Perform a hybrid search using the query vector
        const queryVector = await this.getEmbeddingVector(query);
        const searchResults = await this.searchClient.search(query, {
            searchFields: ["metadata_spo_item_name", "content"],
            select: selectedFields
            // ,vectorSearchOptions: {
            //     queries: [
            //         {
            //             kind: "vector",
            //             fields: ["content"], // Assuming content field is vector searchable
            //             kNearestNeighborsCount: 2,
            //             vector: queryVector
            //         }
            //     ]
            // },
        });

        // If no results, return an empty output
        if (!searchResults || searchResults.results.length === 0) {
            return { output: "", length: 0, tooLong: false };
        }

        // Concatenate the documents string into a single document until the maximum token limit is reached
        let usedTokens = 0;
        let doc = "";
        for await (const result of searchResults.results) {
            const formattedResult = this.formatDocument(result.document.content);
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
     * Formats the result string.
     */
    formatDocument(result) {
        return `<context>${result}</context>`;
    }

    /**
     * Generate embeddings for the user's input.
     */
    async getEmbeddingVector(text) {
        // Ensure all required options are provided for embeddings
        if (!this.options.azureOpenAIApiKey || !this.options.azureOpenAIEndpoint || !this.options.azureOpenAIEmbeddingDeploymentName) {
            throw new Error("Missing required options for generating embeddings.");
        }

        const embeddings = new OpenAIEmbeddings({
            azureApiKey: this.options.azureOpenAIApiKey,
            azureEndpoint: this.options.azureOpenAIEndpoint,
            azureDeployment: this.options.azureOpenAIEmbeddingDeploymentName,
        });

        const result = await embeddings.createEmbeddings(this.options.azureOpenAIEmbeddingDeploymentName, text);

        if (result.status !== "success" || !result.output || result.output.length === 0) {
            throw new Error(`Failed to generate embeddings for the input text: ${text}`);
        }

        return result.output[0];
    }
}

module.exports = {
    AzureAISearchDataSource,
};
const searchDataSource = new AzureAISearchDataSource({
    name: "armelysearchservice", // This name should match the name in config.json
    azureAISearchEndpoint: "https://armelysearchservice.search.windows.net",
    indexName: "sharepoint-index2",
    azureAISearchApiKey: "YOUR_AZURE_SEARCH_API_KEY",
    azureOpenAIApiKey: "YOUR_AZURE_OPENAI_API_KEY",
    azureOpenAIEndpoint: "YOUR_AZURE_OPENAI_ENDPOINT",
    azureOpenAIEmbeddingDeploymentName: "YOUR_AZURE_EMBEDDING_DEPLOYMENT"
});
