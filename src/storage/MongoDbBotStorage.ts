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

import assert from "assert";
import * as mongodb from "mongodb";
import { Storage, StoreItems } from "botbuilder";
import { IBotExtendedStorage } from "./BotExtendedStorage";
import * as logger from "winston";

// tslint:disable-next-line:variable-name
const Fields = {
    userData: "userData",
    conversationData: "conversationData",
    privateConversationData: "privateConversationData",
};

/** MongoDB storage system for Bot Framework. */
export class MongoDbBotStorage implements IBotExtendedStorage {
    private initializePromise: Promise<void>;
    private mongoDb: mongodb.Db;
    private botStateCollection: mongodb.Collection;

    constructor(
        private collectionName: string,
        private connectionString: string
    ) {}

    // Read from storage
    public async read(keys: string[]): Promise<StoreItems> {
        await this.initialize();
        const data: StoreItems = {};
        
        try {
            for (const key of keys) {
                const filter = { key: key };
                const entity = await this.botStateCollection.findOne(filter);
                if (entity) {
                    data[key] = entity.data || {};
                }
            }
            return data;
        } catch (err) {
            logger.error("Error reading from storage", err);
            throw err;
        }
    }

    // Write to storage
    public async write(changes: StoreItems): Promise<void> {
        await this.initialize();
        
        try {
            const promises = Object.entries(changes).map(async ([key, change]) => {
                const filter = { key: key };
                const document = {
                    key: key,
                    data: change,
                    lastUpdate: new Date().valueOf()
                };
                await this.botStateCollection.updateOne(filter, { $set: document }, { upsert: true });
            });
            
            await Promise.all(promises);
        } catch (err) {
            logger.error("Error writing to storage", err);
            throw err;
        }
    }

    // Delete from storage
    public async delete(keys: string[]): Promise<void> {
        await this.initialize();
        
        try {
            const filter = { key: { $in: keys } };
            await this.botStateCollection.deleteMany(filter);
        } catch (err) {
            logger.error("Error deleting from storage", err);
            throw err;
        }
    }

    // Lookup user data by AAD object id
    public async getUserDataByAadObjectIdAsync(aadObjectId: string): Promise<any> {
        await this.initialize();

        const filter = {
            field: Fields.userData,
            "data.aadObjectId": aadObjectId,
        };
        const entity = await this.botStateCollection.findOne(filter);
        return entity ? entity.data || {} : null;
    }

    public getAAdObjectId(userData: any): string {
        return userData.aadObjectId;
    }

    public setAAdObjectId(userData: any, aadObjectId: string): void {
        userData.aadObjectId = aadObjectId;
    }

    // Initialize the storage
    private async initialize(): Promise<void> {
        if (!this.initializePromise) {
            this.initializePromise = this.initializeWorker();
        }
        return this.initializePromise;
    }

    private async initializeWorker(): Promise<void> {
        if (!this.mongoDb) {
            try {
                this.mongoDb = await mongodb.MongoClient.connect(this.connectionString);
                this.botStateCollection = await this.mongoDb.collection(this.collectionName);

                // Set up indexes
                await this.botStateCollection.createIndex({ key: 1 });
                await this.botStateCollection.createIndex({ lastUpdate: 1 });
            } catch (e) {
                logger.error(`Error initializing MongoDB: ${e.message}`, e);
                this.close();
                this.initializePromise = null;
                throw e;
            }
        }
    }

    private close(): void {
        this.botStateCollection = null;
        if (this.mongoDb) {
            this.mongoDb.close();
            this.mongoDb = null;
        }
    }
}
