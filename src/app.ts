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
import * as bodyParser from "body-parser";
import * as path from "path";
import * as winston from "winston";
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as storage from "./storage";
import { TeamsBot } from "./TeamsBot";
import uploadRouter from './routes/upload-image';

// Environment variables
const PORT = parseInt(process.env.PORT || "3978", 10);
const BASE_URI = process.env.BASE_URI || `http://localhost:${PORT}`;
const MONGODB_CONNECTION_STRING = process.env.MONGODB_CONNECTION_STRING;
const MONGODB_BOT_STATE_COLLECTION = process.env.MONGODB_BOT_STATE_COLLECTION || "botstate";
const BOT_STORAGE = process.env.BOT_STORAGE;
const MICROSOFT_APP_ID = process.env.MICROSOFT_APP_ID;
const MICROSOFT_APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD;

// Initialize logger
initLogger();

// Create Express app
const app = express();
app.set("port", PORT);

// Configure middleware
app.use(express.static(path.join(__dirname, "../../../public")));
app.use(bodyParser.json());
app.use(express.static('public'));

// Add CORS support for Teams
app.use((req, res, next) => {
    res.setHeader("Access-Control-Allow-Origin", "*");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
    next();
});

// Add request logging
app.use((req, res, next) => {
    winston.info(`Incoming request: ${req.method} ${req.url}`, {
        headers: req.headers,
        query: req.query,
        body: req.body
    });
    next();
});

// Configure bot storage
let botStorage: builder.IBotStorage;
if (BOT_STORAGE === "mongodb") {
    if (!MONGODB_CONNECTION_STRING) {
        throw new Error("MONGODB_CONNECTION_STRING environment variable is required when using mongodb storage");
    }
    botStorage = new storage.MongoDbBotStorage(
        MONGODB_BOT_STATE_COLLECTION,
        MONGODB_CONNECTION_STRING
    );
    winston.info("Using MongoDB for bot storage");
} else {
    botStorage = new builder.MemoryBotStorage();
    winston.info("Using in-memory bot storage");
}

// Create bot
const connector = new msteams.TeamsChatConnector({
    appId: MICROSOFT_APP_ID,
    appPassword: MICROSOFT_APP_PASSWORD,
});

const botSettings = {
    storage: botStorage,
};

const bot = new TeamsBot(connector as unknown as builder.ChatConnector, botSettings);

// Set up bot endpoint
app.post("/api/messages", connector.listen());

// Log bot errors
bot.on("error", (error: Error) => {
    winston.error("Bot error:", error);
});

// Configure routes
app.use('/api/upload-image', uploadRouter);

// Configure tabs and views
let tabs = require("./tabs");
tabs.setup(app);

// Health check endpoint
app.get("/ping", (req: express.Request, res: express.Response) => {
    res.status(200).send("OK");
});

// Error handling middleware
app.use((err: any, req: express.Request, res: express.Response, next: express.NextFunction) => {
    winston.error('Error handling request:', err);
    res.status(500).send('Internal Server Error');
});

// Start server
app.listen(PORT, "0.0.0.0", () => {
    winston.info(`Server Configuration:`, {
        port: PORT,
        baseUri: BASE_URI,
        environment: process.env.NODE_ENV || 'development'
    });

    winston.info(`Server is running on port ${PORT}`);
    winston.info(`Bot messaging endpoint: ${BASE_URI}/api/messages`);
});

// Logger initialization
function initLogger(): void {
    winston.configure({
        transports: [
            new winston.transports.Console({
                level: process.env.LOG_LEVEL || 'info',
                handleExceptions: true
            })
        ]
    });
}
