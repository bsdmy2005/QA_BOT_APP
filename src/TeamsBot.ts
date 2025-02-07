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
    CardFactory,
    TeamsInfo,
    BotFrameworkAdapter,
    BotHandler,
    ConversationState,
    UserState
} from "botbuilder";
import logger from './utils/Logger';

import { RootDialog } from "./dialogs/RootDialog";
import { fetchTemplates, cardTemplates } from "./dialogs/CardTemplates";
import { renderACAttachment } from "./utils";
import { questionsService } from "./services/questions-service";
import { TaskModuleSizes } from "./constants";
import { QuestionsService } from './services/questions';
import { QABot } from './bot/QABot';
import { QuestionDialog } from './dialogs/QuestionDialog';

export class TeamsBot extends TeamsActivityHandler {
    private qaBot: QABot;
    private questionsService: QuestionsService;

    constructor(
        private adapter: BotFrameworkAdapter,
        private conversationState: ConversationState,
        private userState: UserState
    ) {
        super();
        const questionDialog = new QuestionDialog(logger);
        this.qaBot = new QABot(this.conversationState, this.userState, questionDialog, logger);
        this.questionsService = new QuestionsService();
    }

    onTurn(handler: BotHandler): this {
        return super.onTurn(handler);
    }

    protected async handleTeamsTaskModuleFetch(context: TurnContext, taskModuleRequest: any): Promise<any> {
        try {
            logger.info("Processing task/fetch", {
                taskModule: taskModuleRequest.data.taskModule,
                questionId: taskModuleRequest.data.questionId
            });

            const taskModule = taskModuleRequest.data.taskModule.toLowerCase();

            if (taskModule === "viewquestion") {
                const questionId = taskModuleRequest.data.questionId;
                return {
                    task: {
                        type: "continue",
                        value: {
                            title: "Question Details",
                            height: 1020,
                            width: 1632,
                            url: `${process.env.BASE_URI}/question/${questionId}`,
                            fallbackUrl: `${process.env.BASE_URI}/question/${questionId}`
                        }
                    }
                };
            }

            if (taskModule === "askquestion") {
                return {
                    task: {
                        type: "continue",
                        value: {
                            title: "Ask a Question",
                            height: 1020,
                            width: 1632,
                            url: `${process.env.BASE_URI}/customform-tiptap`,
                            fallbackUrl: `${process.env.BASE_URI}/customform-tiptap`
                        }
                    }
                };
            }

            if (fetchTemplates[taskModule]) {
                return {
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
            }

            throw new Error(`Task module template for ${taskModule} not found.`);
        } catch (error) {
            logger.error("Error in task/fetch", { error });
            throw error;
        }
    }

    protected async handleTeamsTaskModuleSubmit(context: TurnContext, taskModuleRequest: any): Promise<any> {
        try {
            logger.info("Processing task/submit", {
                data: taskModuleRequest.data
            });

            if (!taskModuleRequest.data) {
                return null;
            }

            if (taskModuleRequest.data.type === "answer_accepted" || 
                taskModuleRequest.data.type === "answer_submitted") {
                await this.handleAnswerResponse(context, taskModuleRequest);
                return null;
            }

            if (taskModuleRequest.data.title && taskModuleRequest.data.text && taskModuleRequest.data.userId) {
                await this.handleQuestionSubmission(context, taskModuleRequest);
                return null;
            }

            // Handle other task module responses
            switch (taskModuleRequest.data.taskResponse) {
                case "message":
                    await context.sendActivity("**task/submit results from the Adaptive card:**\n```" + JSON.stringify(taskModuleRequest) + "```");
                    return { type: "message" };
                case "continue":
                    const fetchResponse = fetchTemplates.submitResponse;
                    fetchResponse.task.value.card = renderACAttachment(cardTemplates.acSubmitResponse, { results: JSON.stringify(taskModuleRequest.data) });
                    return fetchResponse;
                case "final":
                    return { type: "final" };
                default:
                    logger.info("Processing HTML task module response", {
                        data: taskModuleRequest.data
                    });
                    await context.sendActivity("**task/submit results from HTML or deep link:**\n\n```" + JSON.stringify(taskModuleRequest.data) + "```");
                    return fetchTemplates.submitMessageResponse;
            }
        } catch (error) {
            logger.error("Error in task/submit", { error });
            throw error;
        }
    }

    private async handleAnswerResponse(context: TurnContext, taskModuleRequest: any): Promise<void> {
        try {
            // Get fresh data from the database
            const questionId = taskModuleRequest.data.data.question.id;
            const updatedQuestion = await this.questionsService.getQuestion(questionId);
            
            if (!updatedQuestion) {
                throw new Error('Question not found');
            }

            const card = CardFactory.adaptiveCard({
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
                                text: "✨ Question",
                                size: "Small",
                                weight: "Bolder",
                                color: "Accent",
                                spacing: "Small"
                            },
                            {
                                type: "TextBlock",
                                text: updatedQuestion.title,
                                wrap: true,
                                weight: "Bolder",
                                size: "ExtraLarge",
                                color: "Default",
                                spacing: "Small"
                            }
                        ],
                        spacing: "Large",
                        bleed: true
                    },
                    {
                        type: "Container",
                        spacing: "Medium",
                        items: [
                            {
                                type: "TextBlock",
                                text: `Posted by ${updatedQuestion.profile.firstName} ${updatedQuestion.profile.lastName} • ${new Date(updatedQuestion.createdAt).toLocaleString()}`,
                                wrap: true,
                                size: "Small",
                                isSubtle: true,
                                spacing: "Small"
                            }
                        ]
                    },
                    {
                        type: "Container",
                        spacing: "Large",
                        items: [
                            {
                                type: "TextBlock",
                                text: updatedQuestion.body.replace(/<[^>]*>/g, ''),
                                wrap: true,
                                size: "Medium",
                                spacing: "Medium"
                            }
                        ]
                    },
                    {
                        type: "Container",
                        spacing: "ExtraLarge",
                        separator: true,
                        items: [
                            {
                                type: "TextBlock",
                                text: `💬 Answers (${updatedQuestion.answers.length})`,
                                wrap: true,
                                weight: "Bolder",
                                size: "Medium",
                                color: "Accent",
                                spacing: "Medium"
                            }
                        ]
                    },
                    ...updatedQuestion.answers.map((answer, index) => ({
                        type: "Container",
                        spacing: index === 0 ? "Medium" : "Large",
                        style: answer.accepted ? "emphasis" : "default",
                        items: [
                            {
                                type: "Container",
                                spacing: "None",
                                items: [
                                    {
                                        type: "TextBlock",
                                        text: answer.accepted ? "✅ Accepted Answer" : `Answer ${index + 1}`,
                                        wrap: true,
                                        size: "Small",
                                        weight: "Bolder",
                                        color: answer.accepted ? "Good" : "Accent",
                                        spacing: "Small"
                                    },
                                    {
                                        type: "TextBlock",
                                        text: `${answer.profile.firstName} ${answer.profile.lastName} • ${new Date(answer.createdAt).toLocaleString()}`,
                                        wrap: true,
                                        size: "Small",
                                        isSubtle: true,
                                        spacing: "Small"
                                    }
                                ]
                            },
                            {
                                type: "TextBlock",
                                text: answer.body.replace(/<[^>]*>/g, ''),
                                wrap: true,
                                size: "Medium",
                                spacing: "Medium"
                            }
                        ],
                        separator: true
                    }))
                ],
                actions: [
                    {
                        type: "Action.Submit",
                        title: "View Full Question & Answers",
                        style: "positive",
                        data: {
                            msteams: {
                                type: "task/fetch"
                            },
                            taskModule: "viewQuestion",
                            questionId: updatedQuestion.id
                        }
                    }
                ]
            });

            await context.sendActivity({ attachments: [card] });
        } catch (error) {
            logger.error('Error handling answer response:', error);
            throw error;
        }
    }

    private async handleQuestionSubmission(context: TurnContext, taskModuleRequest: any): Promise<void> {
        const formattedQuestion = {
            ...taskModuleRequest.data,
            text: taskModuleRequest.data.text,
            rawHtml: taskModuleRequest.data.text
        };
        
        const savedQuestion = await questionsService.addQuestion(
            formattedQuestion.title,
            formattedQuestion.text,
            formattedQuestion.userName
        );
        
        function convertHtmlToMarkdown(html: string): { text: string; images: string[] } {
            if (!html) return { text: '', images: [] };
            
            const images: string[] = [];
            const textWithoutImages = html.replace(/<img[^>]*src=["']([^"']*)["'][^>]*>/gi, (match, src) => {
                images.push(src);
                return '';
            });
            
            let text = textWithoutImages
                .replace(/<ul>([\s\S]*?)<\/ul>/gi, (match, content) => {
                    return content.split('</li>')
                        .filter(item => item.trim())
                        .map(item => {
                            item = item.replace(/<li>/gi, '').trim();
                            return `• ${item}\r`;
                        })
                        .join('');
                })
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
                .replace(/<p>([\s\S]*?)<\/p>/gi, '$1\r\r')
                .replace(/<br\s*\/?>/gi, '\r')
                .replace(/<(strong|b)>([\s\S]*?)<\/(strong|b)>/gi, '**$2**')
                .replace(/<(em|i)>([\s\S]*?)<\/(em|i)>/gi, '_$2_')
                .replace(/<a[^>]*href=["']([^"']*)["'][^>]*>([\s\S]*?)<\/a>/gi, '[$2]($1)')
                .replace(/<[^>]+>/g, '')
                .replace(/\r{3,}/g, '\r\r')
                .replace(/\s+$/gm, '')
                .trim();

            return { text, images };
        }

        const { text: convertedText, images } = convertHtmlToMarkdown(savedQuestion.body);

        const userName = `${savedQuestion.profile.firstName} ${savedQuestion.profile.lastName}`;

        const card = CardFactory.adaptiveCard({
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
                            text: `Posted by ${userName} • ${new Date(savedQuestion.createdAt).toLocaleString()}`,
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
        });
        
        await context.sendActivity({ attachments: [card] });
    }
}
