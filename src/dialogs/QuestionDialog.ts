import {
    ComponentDialog,
    WaterfallDialog,
    WaterfallStepContext
} from 'botbuilder-dialogs';
import { CardFactory } from 'botbuilder';
import { Logger } from 'winston';
import { questionsService } from '../services/questions-service';
import { Question, Profile } from '../db/schema';

const QUESTION_DIALOG = 'QUESTION_DIALOG';
const QUESTION_WATERFALL_DIALOG = 'QUESTION_WATERFALL_DIALOG';

export class QuestionDialog extends ComponentDialog {
    private logger: Logger;
    private static lastAskCardId: string;

    constructor(logger: Logger) {
        super(QUESTION_DIALOG);
        
        this.logger = logger;

        // Add waterfall dialog
        this.addDialog(new WaterfallDialog(QUESTION_WATERFALL_DIALOG, [
            this.initialStep.bind(this),
            this.finalStep.bind(this)
        ]));

        // Set the initial dialog to run
        this.initialDialogId = QUESTION_WATERFALL_DIALOG;
    }

    private async initialStep(stepContext: WaterfallStepContext) {
        this.logger.info('Question dialog - Initial step');
        
        const text = stepContext.context.activity.text?.toLowerCase().trim();
        
        switch (text) {
            case 'ask':
                // Create an adaptive card for asking questions
                const askCard = CardFactory.adaptiveCard({
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
                const response = await stepContext.context.sendActivity({ attachments: [askCard] });
                QuestionDialog.lastAskCardId = response.id; // Store the activity ID
                this.logger.info('Ask card sent and ID stored:', { activityId: response.id });
                break;

            case 'questions':
                try {
                    const questions = await questionsService.getQuestions() as (Question & { profile: Profile })[];
                    if (questions.length === 0) {
                        await stepContext.context.sendActivity("No questions have been asked yet. Be the first to ask one!");
                        break;
                    }

                    // Create a card showing recent questions
                    const questionsCard = CardFactory.adaptiveCard({
                        type: "AdaptiveCard",
                        version: "1.2",
                        body: [
                            {
                                type: "TextBlock",
                                text: "Recent Questions",
                                size: "Large",
                                weight: "Bolder"
                            },
                            ...questions.slice(0, 5).map(question => {
                                // Safely handle profile data
                                const authorName = question.profile ? 
                                    `${question.profile.firstName || ''} ${question.profile.lastName || ''}`.trim() : 
                                    'Anonymous';
                                
                                const qaAppUrl = `${process.env.QA_APP_URI}/qna/${question.id}`;
                                
                                return {
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
                                            text: `Posted by ${authorName} • ${new Date(question.createdAt).toLocaleString()}`,
                                            wrap: true,
                                            size: "Small",
                                            isSubtle: true
                                        },
                                        {
                                            type: "ActionSet",
                                            actions: [
                                                {
                                                    type: "Action.OpenUrl",
                                                    title: "Open in Q&A App",
                                                    url: qaAppUrl
                                                },
                                                {
                                                    type: "Action.Submit",
                                                    title: "View in Teams",
                                                    data: {
                                                        msteams: {
                                                            type: "task/fetch"
                                                        },
                                                        taskModule: "viewQuestion",
                                                        questionId: question.id
                                                    }
                                                }
                                            ]
                                        }
                                    ],
                                    separator: true
                                };
                            })
                        ],
                        actions: [
                            {
                                type: "Action.Submit",
                                title: "Ask a Question",
                                data: {
                                    msteams: {
                                        type: "task/fetch"
                                    },
                                    taskModule: "askquestion"
                                },
                                style: "positive"
                            }
                        ]
                    });
                    await stepContext.context.sendActivity({ attachments: [questionsCard] });
                } catch (error) {
                    this.logger.error('Error fetching questions:', error);
                    await stepContext.context.sendActivity("Sorry, I couldn't fetch the questions at this time.");
                }
                break;

            case 'help':
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
                break;

            default:
                // Show the ask card as default action
                const defaultCard = CardFactory.adaptiveCard({
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
                            text: "Click the button below to ask your question, or type 'help' to see all available commands.",
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
                const defaultResponse = await stepContext.context.sendActivity({ attachments: [defaultCard] });
                QuestionDialog.lastAskCardId = defaultResponse.id; // Store the activity ID
                this.logger.info('Default ask card sent and ID stored:', { activityId: defaultResponse.id });
                break;
        }
        
        return await stepContext.endDialog();
    }

    private async finalStep(stepContext: WaterfallStepContext) {
        this.logger.info('Question dialog - Final step');
        return await stepContext.endDialog();
    }

    public static getLastAskCardId(): string {
        return this.lastAskCardId;
    }
} 