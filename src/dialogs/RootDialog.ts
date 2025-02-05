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

import * as builder from "botbuilder";
import * as constants from "../constants";
import * as utils from "../utils";
import * as logger from "winston";
import { BotFrameworkCard } from "./BotFrameworkCard";

export class RootDialog extends builder.IntentDialog {
    constructor() {
        super();
    }

    // Register the dialogs with the bot
    public register(bot: builder.UniversalBot): void {
        bot.dialog(constants.DialogId.Root, this);

        this.onBegin((session, args, next) => { logger.verbose("onDialogBegin called"); this.onDialogBegin(session, args, next); });
        this.onDefault((session) => { logger.verbose("onDefault called"); this.onMessageReceived(session); });
        new BotFrameworkCard(constants.DialogId.BFCard).register(bot, this);
        this.matches(/bfcard/i, constants.DialogId.BFCard);
        this.matches(/^ask/i, (session) => {
            // Create a card with a button to launch the TipTap form
            const card = {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: {
                    type: "AdaptiveCard",
                    version: "1.2",
                    body: [
                        {
                            type: "TextBlock",
                            text: "Ask a Question",
                            size: "Large",
                            weight: "Bolder"
                        },
                        {
                            type: "TextBlock",
                            text: "Click the button below to ask your question using our rich text editor.",
                            wrap: true
                        }
                    ],
                    actions: [
                        {
                            type: "Action.Submit",
                            title: "Ask Question",
                            data: {
                                msteams: {
                                    type: "task/fetch"
                                },
                                taskModule: "askquestion"
                            }
                        }
                    ]
                }
            };
            session.send(new builder.Message(session).addAttachment(card));
        });
    }

    // Handle start of dialog
    private async onDialogBegin(session: builder.Session, args: any, next: () => void): Promise<void> {
        next();
    }

    // Handle message
    private async onMessageReceived(session: builder.Session): Promise<void> {
        console.log("Context: " + JSON.stringify(utils.getContext(null, session)));
        if (session.message.text === "") {
            console.log("Empty message received");
            // This is a response from a generated AC card
            if (session.message.value !== undefined) {
                session.send("**Action.Submit results:**\n```" + JSON.stringify(session.message.value) + "```");
            }
        }
    }
}
