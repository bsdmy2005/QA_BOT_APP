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

import {
    TeamsActivityHandler,
    TurnContext,
    MessagingExtensionQuery,
    MessagingExtensionResponse,
    CardFactory,
    MessagingExtensionResult
} from 'botbuilder';
import { TeamsBot } from "./TeamsBot";
import * as faker from "faker";
import { Logger } from 'winston';

export class MessagingExtension extends TeamsActivityHandler {
    private logger: Logger;

    constructor(
        private bot: TeamsBot,
        logger: Logger
    ) {
        super();
        this.logger = logger;
    }

    public async handleTeamsMessagingExtensionQuery(
        context: TurnContext,
        query: MessagingExtensionQuery
    ): Promise<MessagingExtensionResponse> {
        this.logger.info('Processing messaging extension query', { 
            commandId: query.commandId,
            parameters: query.parameters 
        });

        switch (query.commandId) {
            case 'getRandomText':
                return this.generateRandomResponse(query);
            default:
                this.logger.warn('Unknown command in messaging extension', { commandId: query.commandId });
                throw new Error('Not implemented');
        }
    }

    private generateRandomResponse(query: MessagingExtensionQuery): MessagingExtensionResponse {
        // If the user supplied a title via the cardTitle parameter then use it or use a fake title
        const titleParam = query.parameters?.find(p => p.name === "cardTitle");
        const title = titleParam ? titleParam.value : faker.lorem.sentence();

        const randomImageUrl = "https://loremflickr.com/200/200"; // Faker's random images uses lorempixel.com, which has been down a lot

        // Generate 5 results to send with fake text and fake images
        const attachments = Array.from({ length: 5 }, (_, i) => {
            return CardFactory.thumbnailCard(
                title,
                faker.lorem.paragraph(),
                [{ url: `${randomImageUrl}?random=${i}` }]
            );
        });

        const response: MessagingExtensionResult = {
            type: 'result',
            attachmentLayout: 'list',
            attachments: attachments
        };

        this.logger.info('Generated random response', { 
            title,
            attachmentCount: attachments.length 
        });

        return { composeExtension: response };
    }
}
