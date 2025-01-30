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
import * as msteams from "botbuilder-teams";
import * as utils from "./utils";
import * as logger from "winston";

import { RootDialog } from "./dialogs/RootDialog";
import { fetchTemplates, cardTemplates } from "./dialogs/CardTemplates";
import { renderACAttachment } from "./utils";
import { questionsService } from "./services/questions-service";

export class TeamsBot extends builder.UniversalBot {

    constructor(
        public _connector: builder.IConnector,
        public _botSettings: any,
    )
    {
        super(_connector, _botSettings);
        this.set("persistentConversationData", true);

        // Handle generic invokes
        let teamsConnector = this._connector as msteams.TeamsChatConnector;
        teamsConnector.onInvoke(async (event, cb) => {
            try {
                logger.info("Received invoke event", {
                    name: (event as any).name,
                    value: (event as any).value
                });
                await this.onInvoke(event, cb);
            } catch (e) {
                logger.error("Invoke handler failed", {
                    error: e,
                    event: event
                });
                cb(e, null, 500);
            }
        });

        // Register dialogs
        new RootDialog().register(this);
    }

    // Handle incoming invoke
    private async onInvoke(event: builder.IEvent, cb: (err: Error, body: any, status?: number) => void): Promise<void> {
        logger.info("Processing invoke", {
            context: utils.getContext(event)
        });

        let session = await utils.loadSessionAsync(this, event);
        if (!session) {
            logger.error("Failed to load session for invoke");
            cb(new Error("Failed to load session"), null, 500);
            return;
        }

        let invokeType = (event as any).name;
        let invokeValue = (event as any).value;

        logger.info("Invoke details", {
            type: invokeType,
            value: invokeValue
        });

        if (invokeType === undefined) {
            invokeType = null;
        }

        switch (invokeType) {
            case "task/fetch": {
                try {
                    let taskModule = invokeValue.data.taskModule.toLowerCase();
                    logger.info("Processing task/fetch", {
                        taskModule: taskModule,
                        questionId: invokeValue.data.questionId
                    });

                    if (taskModule === "viewquestion") {
                        // Create a template for viewing question details
                        const questionId = invokeValue.data.questionId;
                        const response = {
                            task: {
                                type: "continue",
                                value: {
                                    title: "Question Details",
                                    height: 720,
                                    width: 1000,
                                    url: `${process.env.BASE_URI}/question/${questionId}`,
                                    fallbackUrl: `${process.env.BASE_URI}/question/${questionId}`
                                }
                            }
                        };
                        logger.info("Sending view question response", { response });
                        cb(null, response, 200);
                        return;
                    }

                    if (fetchTemplates[taskModule] !== undefined) {
                        const response = {
                            task: {
                                type: "continue",
                                value: {
                                    title: fetchTemplates[taskModule].title,
                                    height: fetchTemplates[taskModule].height,
                                    width: fetchTemplates[taskModule].width,
                                    url: fetchTemplates[taskModule].url,
                                    fallbackUrl: fetchTemplates[taskModule].fallbackUrl,
                                    completionBotId: process.env.MICROSOFT_APP_ID
                                }
                            }
                        };
                        logger.info("Sending task/fetch response", { response });
                        cb(null, response, 200);
                    } else {
                        let error = `Error: task module template for ${(invokeValue.taskModule === undefined ? "<undefined>" : invokeValue.taskModule)} not found.`;
                        logger.error(error);
                        cb(new Error(error), null, 404);
                    }
                } catch (error) {
                    logger.error("Error in task/fetch", {
                        error: error,
                        invokeValue: invokeValue
                    });
                    cb(error, null, 500);
                }
                break;
            }
            case "task/submit": {
                try {
                    logger.info("Processing task/submit", {
                        data: invokeValue.data
                    });

                    if (invokeValue.data !== undefined) {
                        // Save the question if it's from our custom form
                        if (invokeValue.data.title && invokeValue.data.text && invokeValue.data.userId) {
                            try {
                                // Format the text with markdown
                                const formattedQuestion = {
                                    ...invokeValue.data,
                                    text: invokeValue.data.text, // Remove markdown formatting since we're using rich text
                                    rawHtml: invokeValue.data.text // Store the raw HTML for future reference
                                };
                                
                                // Save the question
                                const savedQuestion = await questionsService.addQuestion(formattedQuestion);
                                
                                // Function to convert HTML to markdown-style formatting
                                function convertHtmlToMarkdown(html: string): { text: string; images: string[] } {
                                    if (!html) return { text: '', images: [] };
                                    
                                    // Extract images and store their information
                                    const images: string[] = [];
                                    const textWithoutImages = html.replace(/<img[^>]*src=["']([^"']*)["'][^>]*>/gi, (match, src) => {
                                        images.push(src);
                                        return ''; // Remove image placeholder completely
                                    });
                                    
                                    // Convert other HTML to markdown
                                    let text = textWithoutImages
                                        // Handle unordered lists
                                        .replace(/<ul>([\s\S]*?)<\/ul>/gi, (match, content) => {
                                            return content.split('</li>')
                                                .filter(item => item.trim())
                                                .map(item => {
                                                    item = item.replace(/<li>/gi, '').trim();
                                                    return `• ${item}\r`;
                                                })
                                                .join('');
                                        })
                                        // Handle ordered lists
                                        .replace(/<ol>([\s\S]*?)<\/ol>/gi, (match, content) => {
                                            let index = 1;
                                            return content.split('</li>')
                                                .filter(item => item.trim())
                                                .map(item => {
                                                    item = item.replace(/<li>/gi, '').trim();
                                                    return `${index++}. ${item}\r`;
                                                })
                                                .join('');
                                        })
                                        // Handle paragraphs
                                        .replace(/<p>([\s\S]*?)<\/p>/gi, '$1\r\r')
                                        // Handle line breaks
                                        .replace(/<br\s*\/?>/gi, '\r')
                                        // Handle bold text
                                        .replace(/<(strong|b)>([\s\S]*?)<\/(strong|b)>/gi, '**$2**')
                                        // Handle italic text
                                        .replace(/<(em|i)>([\s\S]*?)<\/(em|i)>/gi, '_$2_')
                                        // Handle links
                                        .replace(/<a[^>]*href=["']([^"']*)["'][^>]*>([\s\S]*?)<\/a>/gi, '[$2]($1)')
                                        // Remove any remaining HTML tags
                                        .replace(/<[^>]+>/g, '')
                                        // Fix spacing issues
                                        .replace(/\r{3,}/g, '\r\r')
                                        .replace(/\s+$/gm, '')
                                        .trim();

                                    return { text, images };
                                }

                                // Create and send the confirmation card
                                const { text: convertedText, images } = convertHtmlToMarkdown(savedQuestion.text);

                                const card = {
                                    contentType: "application/vnd.microsoft.card.adaptive",
                                    content: {
                                        type: "AdaptiveCard",
                                        version: "1.2",
                                        $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
                                        msteams: {
                                            width: "Full"
                                        },
                                        style: "default",
                                        body: [
                                            {
                                                type: "Container",
                                                style: "emphasis",
                                                items: [
                                                    {
                                                        type: "TextBlock",
                                                        text: savedQuestion.title,
                                                        wrap: true,
                                                        weight: "Bolder",
                                                        size: "Large",
                                                        color: "Default"
                                                    }
                                                ],
                                                spacing: "Medium",
                                                bleed: true
                                            },
                                            {
                                                type: "Container",
                                                spacing: "Medium",
                                                items: [
                                                    {
                                                        type: "TextBlock",
                                                        text: convertedText.length > 150 ? convertedText.substring(0, 150) + "..." : convertedText,
                                                        wrap: true,
                                                        size: "Medium"
                                                    }
                                                ]
                                            },
                                            {
                                                type: "Container",
                                                spacing: "Small",
                                                items: [
                                                    {
                                                        type: "TextBlock",
                                                        text: `Asked by ${savedQuestion.userName} • ${new Date(savedQuestion.timestamp).toLocaleString()}`,
                                                        wrap: true,
                                                        size: "Small",
                                                        isSubtle: true
                                                    }
                                                ]
                                            }
                                        ],
                                        actions: [
                                            {
                                                type: "Action.Submit",
                                                title: "View Full Question",
                                                style: "positive",
                                                data: {
                                                    msteams: {
                                                        type: "task/fetch"
                                                    },
                                                    taskModule: "viewQuestion",
                                                    questionId: savedQuestion.id
                                                }
                                            }
                                        ]
                                    }
                                };
                                
                                // Send the card to the conversation
                                const message = new builder.Message(session)
                                    .addAttachment(card);
                                session.send(message);
                                
                                // Close the task module
                                cb(null, null, 200);

                                // Add view question template
                                fetchTemplates.viewQuestion = {
                                    title: "Question Details",
                                    height: 600,
                                    width: 800,
                                    url: `${process.env.BASE_URI}/question/${savedQuestion.id}`,
                                    fallbackUrl: `${process.env.BASE_URI}/question/${savedQuestion.id}`
                                };
                                return;
                            } catch (error) {
                                logger.error("Error saving question", { error });
                                cb(error, null, 500);
                                return;
                            }
                        }

                        // Handle other task module responses
                        switch (invokeValue.data.taskResponse) {
                            case "message":
                                session.send("**task/submit results from the Adaptive card:**\n```" + JSON.stringify(invokeValue) + "```");
                                cb(null, { type: "message" }, 200);
                                break;
                            case "continue":
                                let fetchResponse = fetchTemplates.submitResponse;
                                fetchResponse.task.value.card = renderACAttachment(cardTemplates.acSubmitResponse, { results: JSON.stringify(invokeValue.data) });
                                logger.info("Sending continue response", {
                                    response: fetchResponse
                                });
                                cb(null, fetchResponse, 200);
                                break;
                            case "final":
                                cb(null, { type: "final" }, 200);
                                break;
                            default:
                                logger.info("Processing HTML task module response", {
                                    data: invokeValue.data
                                });
                                cb(null, fetchTemplates.submitMessageResponse, 200);
                                session.send("**task/submit results from HTML or deep link:**\n\n```" + JSON.stringify(invokeValue.data) + "```");
                        }
                    } else {
                        logger.warn("Received task/submit with undefined data");
                        cb(null, null, 200);
                    }
                } catch (error) {
                    logger.error("Error in task/submit", {
                        error: error,
                        invokeValue: invokeValue
                    });
                    cb(error, null, 500);
                }
                break;
            }
            // Invokes don't participate in middleware
            // If the message is not task/*, simulate a normal message and route it, but remember the original invoke message
            case null: {
                try {
                    let fakeMessage: any = {
                        ...event,
                        text: invokeValue.command + " " + JSON.stringify(invokeValue),
                        originalInvoke: event,
                    };

                    logger.info("Processing non-task invoke as message", {
                        message: fakeMessage
                    });

                    session.message = fakeMessage;
                    session.dispatch(session.sessionState, session.message, () => {
                        session.routeToActiveDialog();
                    });
                } catch (error) {
                    logger.error("Error processing non-task invoke", {
                        error: error,
                        event: event
                    });
                    cb(error, null, 500);
                }
                break;
            }
            default: {
                logger.warn("Unknown invoke type", {
                    type: invokeType
                });
                cb(null, "", 200);
            }
        }
    }
}
