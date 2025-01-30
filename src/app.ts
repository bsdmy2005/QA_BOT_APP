// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import * as dotenv from "dotenv";
dotenv.config(); // Load .env file from project root

import express from "express";
import favicon from "serve-favicon";
import * as bodyParser from "body-parser";
import * as path from "path";
import * as logger from "winston";
import * as winston from "winston";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as storage from "./storage";
import { TeamsBot } from "./TeamsBot";
import { MessagingExtension } from "./MessagingExtension";

const PORT = parseInt(process.env.PORT || "3978", 10);
const BASE_URI = process.env.BASE_URI || `http://localhost:${PORT}`;

// Configure logging
logger.remove(logger.transports.Console);
logger.add(logger.transports.Console, {
    level: 'debug',
    handleExceptions: true,
    timestamp: true,
    colorize: true,
    json: false,
    prettyPrint: true
});

const app = express();
app.set("port", PORT);

// Add CORS support for Teams
app.use((req, res, next) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
    next();
});

// Add request logging
app.use((req, res, next) => {
    logger.info(`Incoming request: ${req.method} ${req.url}`, {
        headers: req.headers,
        query: req.query,
        body: req.body
    });
    next();
});

// Add error handling middleware
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
    logger.error('Error handling request', err);
    res.status(500).send('Internal Server Error');
});

app.use(express.static(path.join(__dirname, "../../../public")));
app.use(bodyParser.json());

// Configure bot storage
let botStorageProvider = process.env.BOT_STORAGE;
let botStorage = null;
const mongoDbCollection = process.env.MONGODB_BOT_STATE_COLLECTION || "botstate";
const mongoDbConnection = process.env.MONGODB_CONNECTION_STRING;

if (botStorageProvider === "mongodb") {
    if (!mongoDbConnection) {
        throw new Error("MONGODB_CONNECTION_STRING environment variable is required when using mongodb storage");
    }
    botStorage = new storage.MongoDbBotStorage(
        mongoDbCollection,
        mongoDbConnection
    );
} else {
    // Default to memory storage
    botStorage = new builder.MemoryBotStorage();
}

// Create bot
let connector = new msteams.TeamsChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD,
});
let botSettings = {
    storage: botStorage,
};
let bot = new TeamsBot(connector as builder.ChatConnector, botSettings);

// Adding a messaging extension to our bot
let messagingExtension = new MessagingExtension(bot);

// Set up route for the bot to listen.
// NOTE: This endpoint cannot be changed and must be api/messages
app.post("/api/messages", connector.listen());

// Log bot errors
bot.on("error", (error: Error) => {
    logger.error(error.message);
});

// Adding tabs to our app. This will setup routes to various views
let tabs = require("./tabs");
tabs.setup(app);

// Configure ping route
app.get("/ping", (req: express.Request, res: express.Response) => {
    res.status(200).send("OK");
});

// Start our nodejs app
app.listen(PORT, "0.0.0.0", function(): void {
    logger.info(`Server Configuration:`, {
        port: PORT,
        baseUri: BASE_URI,
        environment: process.env.NODE_ENV || 'development'
});

    logger.info(`Server is running on port ${PORT}`);
    logger.info(`Bot messaging endpoint: ${BASE_URI}/api/messages`);
});

function initLogger(): void {

    logger.addColors({
        error: "red",
        warn:  "yellow",
        info:  "green",
        verbose: "cyan",
        debug: "blue",
        silly: "magenta",
    });

    logger.remove(logger.transports.Console);
    logger.add(logger.transports.Console,
        {
            timestamp: () => { return new Date().toLocaleTimeString(); },
            colorize: (process.env.MONOCHROME_CONSOLE) ? false : true,
            prettyPrint: true,
            level: "debug",
        },
    );
}
