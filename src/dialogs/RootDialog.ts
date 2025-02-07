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
    ComponentDialog,
    DialogSet,
    WaterfallDialog,
    WaterfallStepContext
} from 'botbuilder-dialogs';
import { CardFactory, TurnContext } from 'botbuilder';
import * as constants from "../constants";
import * as utils from "../utils";
import { Logger } from 'winston';
import { BotFrameworkCard } from "./BotFrameworkCard";
import { questionsService } from "../services/questions-service";
import { Question, Profile } from "../db/schema";

const ROOT_DIALOG = 'ROOT_DIALOG';
const ROOT_WATERFALL_DIALOG = 'ROOT_WATERFALL_DIALOG';

export class RootDialog extends ComponentDialog {
    private logger: Logger;

    constructor(logger: Logger) {
        super(ROOT_DIALOG);
        
        this.logger = logger;

        // Add waterfall dialog
        this.addDialog(new WaterfallDialog(ROOT_WATERFALL_DIALOG, [
            this.initialStep.bind(this),
            this.finalStep.bind(this)
        ]));

        // Add other dialogs
        this.addDialog(new BotFrameworkCard(constants.DialogId.BFCard, logger));

        // Set the initial dialog to run
        this.initialDialogId = ROOT_WATERFALL_DIALOG;
    }

    private async initialStep(stepContext: WaterfallStepContext) {
        this.logger.info('Root dialog - Initial step');
        
        const text = stepContext.context.activity.text?.toLowerCase();
        
        // Handle ask command
        if (text === 'ask') {
            // Create a card with a button to launch the TipTap form
            const card = CardFactory.adaptiveCard({
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
            });

            await stepContext.context.sendActivity({ attachments: [card] });
        }
        
        // Handle questions command
        else if (text === 'questions') {
            try {
                const questions = await questionsService.getQuestions() as (Question & { profile: Profile })[];
                if (questions.length === 0) {
                    await stepContext.context.sendActivity("No questions have been asked yet. Be the first to ask one!");
                    return;
                }

                // Create a card showing recent questions
                const card = CardFactory.adaptiveCard({
                    type: "AdaptiveCard",
                    version: "1.2",
                    body: [
                        {
                            type: "TextBlock",
                            text: "Recent Questions",
                            size: "Large",
                            weight: "Bolder"
                        },
                        ...questions.slice(0, 5).map(question => ({
                            type: "Container",
                            items: [
                                {
                                    type: "TextBlock",
                                    text: question.title,
                                    wrap: true,
                                    weight: "Bolder"
                                },
                                {
                                    type: "TextBlock",
                                    text: `Posted by ${question.profile.firstName} ${question.profile.lastName} • ${new Date(question.createdAt).toLocaleString()}`,
                                    wrap: true,
                                    size: "Small",
                                    isSubtle: true
                                }
                            ],
                            separator: true,
                            selectAction: {
                                type: "Action.Submit",
                                data: {
                                    msteams: {
                                        type: "task/fetch"
                                    },
                                    taskModule: "viewQuestion",
                                    questionId: question.id
                                }
                            }
                        }))
                    ]
                });

                await stepContext.context.sendActivity({ attachments: [card] });
            } catch (error) {
                this.logger.error('Error fetching questions:', error);
                await stepContext.context.sendActivity("Sorry, I couldn't fetch the questions at this time.");
            }
        }
        
        // Handle help command
        else if (text === 'help') {
            const helpCard = CardFactory.adaptiveCard({
                type: "AdaptiveCard",
                version: "1.2",
                body: [
                    {
                        type: "TextBlock",
                        text: "Q&A Bot Help",
                        size: "Large",
                        weight: "Bolder"
                    },
                    {
                        type: "TextBlock",
                        text: "Available commands:",
                        weight: "Bolder"
                    },
                    {
                        type: "FactSet",
                        facts: [
                            {
                                title: "ask",
                                value: "Create a new question using the rich text editor"
                            },
                            {
                                title: "questions",
                                value: "View recent questions in the current context"
                            },
                            {
                                title: "help",
                                value: "Show this help message"
                            }
                        ]
                    },
                    {
                        type: "TextBlock",
                        text: "Tips:",
                        weight: "Bolder",
                        spacing: "Medium"
                    },
                    {
                        type: "TextBlock",
                        text: "• You can format your questions with rich text and add images\n• Question owners can mark answers as accepted\n• Click on any question to view its full details and answers",
                        wrap: true
                    }
                ]
            });

            await stepContext.context.sendActivity({ attachments: [helpCard] });
        }

        return await stepContext.next();
    }

    private async finalStep(stepContext: WaterfallStepContext) {
        this.logger.info('Root dialog - Final step');
        
        const message = stepContext.context.activity;
        
        if (message.text === "" && message.value) {
            this.logger.info("Processing Action.Submit", { value: message.value });
            await stepContext.context.sendActivity(`**Action.Submit results:**\n\`\`\`${JSON.stringify(message.value)}\`\`\``);
        }
        
        return await stepContext.endDialog();
    }
}
