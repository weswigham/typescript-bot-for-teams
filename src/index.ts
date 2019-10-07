// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { config } from 'dotenv';
import * as path from 'path';
import * as restify from 'restify';

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import { MemoryStorage } from 'botbuilder';

// This bot's main dialog.
import { TypescriptStandupBot } from './bot';
import { TeamsMiddleware, TeamSpecificConversationState } from 'botbuilder-teams';
import { BlobStorage } from 'botbuilder-azure';
import { PatchedTeamsAdapter } from './teamsAdapterPatched';

const ENV_FILE = path.join(__dirname, '..', '.env');
config({ path: ENV_FILE });

// Create HTTP server.
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`\n${server.name} listening to ${server.url}`);
    console.log(`\nGet Bot Framework Emulator: https://aka.ms/botframework-emulator`);
    console.log(`\nTo test your bot, see: https://aka.ms/debug-with-emulator`);
});

const storage = process.env.StorageAccount && process.env.StorageAccessKey && process.env.StorageContainerName ? new BlobStorage({
    containerName: process.env.StorageContainerName,
    storageAccountOrConnectionString: process.env.StorageAccount,
    storageAccessKey: process.env.storageAccessKey
}) : new MemoryStorage();
const botState = new TeamSpecificConversationState(storage);

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new PatchedTeamsAdapter({
    appId: process.env.MicrosoftAppID,
    appPassword: process.env.MicrosoftAppPassword,
});
adapter.use(new TeamsMiddleware());

// Catch-all for any unhandled errors in your bot.
adapter.onTurnError = async (context, error) => {
    // This check writes out errors to console log .vs. app insights.
    console.error(`\n [onTurnError]: ${ error }`);
    // Send a message to the user.
    await context.sendActivity(`Oops. Something went wrong!`);
    // Clear out state
    await botState.delete(context);
};

const bot = new TypescriptStandupBot(botState, storage, adapter);

// Listen for incoming requests.
server.post('/api/messages', (req, res) => {
    // Use the adapter to process the incoming web request into a TurnContext object.
    adapter.processActivity(req, res, bot.onTurn.bind(bot));
});
