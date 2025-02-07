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
import { Logger, createLogger, format, transports } from "winston";
import {
    BotFrameworkAdapter,
    BotFrameworkAdapterSettings,
    MemoryStorage,
    ConversationState,
    UserState,
    Activity
} from 'botbuilder';
import * as storage from "./storage";
import { TeamsBot } from "./TeamsBot";
import uploadRouter from './routes/upload-image';
import { QABot } from './bot/QABot';
import { QuestionDialog } from './dialogs/QuestionDialog';

// Environment variables
const PORT = parseInt(process.env.PORT || "3978", 10);
const BASE_URI = process.env.BASE_URI || `http://localhost:${PORT}`;
const MONGODB_CONNECTION_STRING = process.env.MONGODB_CONNECTION_STRING;
const MONGODB_BOT_STATE_COLLECTION = process.env.MONGODB_BOT_STATE_COLLECTION || "botstate";
const BOT_STORAGE = process.env.BOT_STORAGE;
const MICROSOFT_APP_ID = process.env.MICROSOFT_APP_ID;
const MICROSOFT_APP_PASSWORD = process.env.MICROSOFT_APP_PASSWORD;

// Initialize logger
const logger = initLogger();

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
    // Extract Teams-specific information if available
    const teamsInfo = {
        tenantId: req.headers['x-ms-tenant-id'],
        conversationId: req.headers['x-ms-conversation-id'],
        clientInfo: req.headers['user-agent']?.includes('Microsoft-SkypeBotApi') ? 'Teams' : 'Other'
    };

    // For /api/messages endpoint, only log essential bot information
    if (req.path === '/api/messages' && req.method === 'POST') {
        const body = req.body || {};
        const logData = {
            type: 'Bot Message',
            messageType: body.type,
            activity: {
                type: body.type,
                text: body.text?.substring(0, 50), // Truncate long messages
                from: body.from?.name,
                locale: body.locale,
                conversationType: body.conversation?.conversationType
            },
            teams: teamsInfo
        };
        logger.info('Teams Bot Activity', logData);
    } else {
        // For other endpoints, log standard request info
        const logData = {
            type: 'HTTP Request',
            method: req.method,
            path: req.url,
            query: Object.keys(req.query).length ? req.query : undefined,
            contentType: req.headers['content-type'],
            userAgent: req.headers['user-agent']
        };
        logger.info('API Request', logData);
    }
    next();
});

// Configure bot storage
let botStorage: storage.IBotExtendedStorage;
if (BOT_STORAGE === "mongodb") {
    if (!MONGODB_CONNECTION_STRING) {
        throw new Error("MONGODB_CONNECTION_STRING environment variable is required when using mongodb storage");
    }
    botStorage = new storage.MongoDbBotStorage(
        MONGODB_BOT_STATE_COLLECTION,
        MONGODB_CONNECTION_STRING
    );
    logger.info("Storage configuration", { type: "MongoDB", collection: MONGODB_BOT_STATE_COLLECTION });
} else {
    botStorage = new storage.NullBotStorage();
    logger.info("Storage configuration", { type: "In-Memory" });
}

// Create adapter settings
const adapterSettings: Partial<BotFrameworkAdapterSettings> = {
    appId: MICROSOFT_APP_ID,
    appPassword: MICROSOFT_APP_PASSWORD
};

// Create adapter
const adapter = new BotFrameworkAdapter(adapterSettings);

// Add error handling
adapter.onTurnError = async (context, error) => {
    logger.error("Bot Framework error", {
        error: {
            name: error.name,
            message: error.message,
            stack: error.stack
        }
    });

    await context.sendActivity("Sorry, something went wrong!");
};

// Create storage and state
const memoryStorage = new MemoryStorage();
const conversationState = new ConversationState(memoryStorage);
const userState = new UserState(memoryStorage);

// Create dialog
const dialog = new QuestionDialog(logger);

// Create bot
const bot = new QABot(conversationState, userState, dialog, logger);

// Set up bot endpoint
app.post("/api/messages", async (req, res) => {
    await adapter.process(req, res, (context) => bot.run(context));
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
    logger.error("Request error", {
        path: req.url,
        method: req.method,
        error: {
            name: err.name,
            message: err.message,
            stack: err.stack
        }
    });
    res.status(500).send('Internal Server Error');
});

// Start server
app.listen(PORT, "0.0.0.0", () => {
    logger.info("Server started", {
        config: {
            port: PORT,
            baseUri: BASE_URI,
            environment: process.env.NODE_ENV || 'development'
        }
    });
});

// Logger initialization
function initLogger(): Logger {
    return createLogger({
        level: process.env.LOG_LEVEL || 'info',
        format: format.combine(
            format.timestamp(),
            format.colorize(),
            format.printf(({ timestamp, level, message, ...meta }) => {
                let logMessage = `${timestamp} ${level} ${message}`;
                
                // Format meta data
                if (Object.keys(meta).length) {
                    const metaObj = meta as any;
                    
                    // Special handling for Teams Bot Activity
                    if (metaObj.type === 'Bot Message' && metaObj.activity) {
                        const activity = metaObj.activity as Activity;
                        logMessage += ` [${activity.type || 'unknown'}] from: ${activity.from?.name || 'unknown'} text: "${activity.text || ''}"`;
                    }
                    // Special handling for API Request
                    else if (metaObj.type === 'HTTP Request') {
                        logMessage += ` [${metaObj.method}] ${metaObj.path}`;
                        if (metaObj.query) {
                            logMessage += ` query: ${JSON.stringify(metaObj.query)}`;
                        }
                    }
                    // For other types of logs, just stringify important fields
                    else {
                        const importantFields = ['type', 'message', 'error', 'config'];
                        const relevantData = {};
                        importantFields.forEach(field => {
                            if (metaObj[field]) {
                                relevantData[field] = metaObj[field];
                            }
                        });
                        if (Object.keys(relevantData).length) {
                            logMessage += ` ${JSON.stringify(relevantData)}`;
                        }
                    }
                }

                return logMessage.trim();
            })
        ),
        transports: [
            new transports.Console({
                handleExceptions: true
            })
        ]
    });
}
