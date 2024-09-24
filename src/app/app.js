const { MemoryStorage } = require("botbuilder");
const path = require("path");
const config = require("../config");

// See https://aka.ms/teams-ai-library to learn more about the Teams AI library.
const { Application, ActionPlanner, OpenAIModel, PromptManager } = require("@microsoft/teams-ai");
const { AzureAISearchDataSource } = require("./azureAISearchDataSource");

// Create AI components
const model = new OpenAIModel({
  azureApiKey: config.azureOpenAIKey,
  azureDefaultDeployment: config.azureOpenAIDeploymentName,
  azureEndpoint: config.azureOpenAIEndpoint,

  useSystemMessages: true,
  logRequests: true,
});
const prompts = new PromptManager({
  promptsFolder: path.join(__dirname, "../prompts"),
});
const planner = new ActionPlanner({
  model,
  prompts,
  defaultPrompt: "chat",
});

// Register your data source with planner
planner.prompts.addDataSource(
  new AzureAISearchDataSource({
    name: "armelysearchservice",
    indexName: "sharepoint-index2",
    azureAISearchApiKey: config.azureSearchKey,
    azureAISearchEndpoint: config.azureSearchEndpoint,
    azureOpenAIApiKey: config.azureOpenAIKey,
    azureOpenAIEndpoint: config.azureOpenAIEndpoint,
    azureOpenAIEmbeddingDeploymentName: config.azureOpenAIEmbeddingDeploymentName,
    semanticConfiguration: "my-semantic-config-default", // Ensure semantic search is used
  })
);

// Define storage and application
const storage = new MemoryStorage();
const app = new Application({
  storage,
  ai: {
    planner,
    outputHandler: async (context, response) => {
      if (response.source === "llm") {
        // This indicates that the response is AI-generated
        context.sendActivity(`AI-generated response:\n${response.output}`);
      } else {
        // This indicates that the response comes from a document search
        context.sendActivity(`Here is the information you requested:\n${response.output}`);
      }
    }
  },
});

module.exports = app;
