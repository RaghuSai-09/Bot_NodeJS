const express = require('express');
const { MemoryStorage, ConversationState, UserState } = require('botbuilder');
const { CloudAdapter, ConfigurationBotFrameworkAuthentication } = require('botbuilder');
const dotenv = require('dotenv');
dotenv.config();
const path = require('path');

const ENV_FILE = path.join(__dirname, '.env');
dotenv.config({ path: ENV_FILE });

// Initialize the Express server
const app = express();
app.use(express.json());

// Create adapter.
const adapterSettings = new ConfigurationBotFrameworkAuthentication(process.env);
const adapter = new CloudAdapter(
    adapterSettings
);

//Catch-all for errors. 
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${ error }`);
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
};

// Initialize state
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

const { MyBot } = require('./bot');

// Create the bot instance
const myBot = new MyBot(conversationState, userState);

// Listen for incoming requests
app.post('/api/messages', async (req, res) => {
  await adapter.process(req, res,  (context) =>  myBot.run(context) );
});

// Start the server
const port = process.env.PORT || 3978;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});