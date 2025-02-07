import {
    TeamsActivityHandler,
    TurnContext,
    CardFactory,
    TeamsInfo,
    ConversationState,
    UserState,
    ActivityTypes
} from 'botbuilder';
import { Dialog, DialogSet, DialogState } from 'botbuilder-dialogs';
import { Logger } from 'winston';
import { questionsService } from '../services/questions-service';
import { QuestionDialog } from '../dialogs/QuestionDialog';

const QUESTION_DIALOG = 'QUESTION_DIALOG';

export class QABot extends TeamsActivityHandler {
    private dialogSet: DialogSet;
    private logger: Logger;
    private conversationState: ConversationState;
    private static questionCardIds: Map<string, string> = new Map(); // Store questionId -> activityId

    constructor(conversationState: ConversationState, userState: UserState, dialog: Dialog, logger: Logger) {
        super();
        
        this.logger = logger;
        this.conversationState = conversationState;
        
        // Create DialogSet and add dialogs
        const dialogState = this.conversationState.createProperty<DialogState>("DialogState");
        this.dialogSet = new DialogSet(dialogState);
        this.dialogSet.add(dialog);
    }

    async onMessageActivity(context: TurnContext): Promise<void> {
        try {
            this.logger.info('Received message activity:', {
                text: context.activity.text,
                from: context.activity.from.name,
                value: context.activity.value
            });

            // Handle both text commands and command clicks
            let text = '';
            if (context.activity.value && context.activity.value.command) {
                // Handle command click
                text = context.activity.value.command.toLowerCase();
            } else {
                // Handle text command
                text = context.activity.text?.toLowerCase().trim() || '';
            }

            // Create dialog context
            const dialogContext = await this.dialogSet.createContext(context);
            
            // Continue any existing dialog
            const dialogResult = await dialogContext.continueDialog();

            // If no dialog is active and no response was sent
            if (!context.responded) {
                switch (text) {
                    case 'ask':
                    case 'questions':
                    case 'help':
                        await dialogContext.beginDialog(QUESTION_DIALOG);
                        break;
                    default:
                        // For any other text, also start the question dialog
                        // as it handles showing the default card
                        await dialogContext.beginDialog(QUESTION_DIALOG);
                }
            }
        } catch (error) {
            this.logger.error('Error in onMessageActivity:', error);
            await context.sendActivity('Sorry, I encountered an error processing your request.');
        }
    }

    protected async handleTeamsTaskModuleFetch(context: TurnContext, taskModuleRequest: any): Promise<any> {
        this.logger.info('Processing task/fetch invoke', {
            type: context.activity.type,
            taskModule: taskModuleRequest.data.taskModule
        });

        if (taskModuleRequest.data.taskModule === 'askquestion') {
            return {
                task: {
                    type: 'continue',
                    value: {
                        title: 'Ask a Question',
                        height: 1020,
                        width: 1632,
                        url: `${process.env.BASE_URI}/customform-tiptap`,
                        fallbackUrl: `${process.env.BASE_URI}/customform-tiptap`
                    }
                }
            };
        } else if (taskModuleRequest.data.taskModule === 'viewQuestion') {
            const questionId = taskModuleRequest.data.questionId;
            return {
                task: {
                    type: 'continue',
                    value: {
                        title: 'Question Details',
                        height: 1020,
                        width: 1632,
                        url: `${process.env.BASE_URI}/question/${questionId}`,
                        fallbackUrl: `${process.env.BASE_URI}/question/${questionId}`
                    }
                }
            };
        }
        return undefined;
    }

    protected async handleTeamsTaskModuleSubmit(context: TurnContext, taskModuleRequest: any): Promise<any> {
        this.logger.info('Processing task/submit invoke', {
            type: context.activity.type,
            data: taskModuleRequest.data
        });

        try {
            // Handle answer acceptance
            if (taskModuleRequest.data?.type === 'answer_accepted') {
                // Process synchronously to maintain context
                await this.handleAnswerAcceptance(context, taskModuleRequest);
                return {
                    task: {
                        type: 'message',
                        value: "Answer accepted successfully!"
                    }
                };
            }

            // Handle answer submission
            if (taskModuleRequest.data?.type === 'answer_submitted') {
                // Process synchronously to maintain context
                await this.handleAnswerSubmission(context, taskModuleRequest);
                return {
                    task: {
                        type: 'message',
                        value: "Answer submitted successfully!"
                    }
                };
            }

            // Handle question submission
            if (taskModuleRequest.data?.title && taskModuleRequest.data?.text) {
                // Process synchronously to maintain context
                await this.handleQuestionSubmission(context, taskModuleRequest);
                return {
                    task: {
                        type: 'message',
                        value: "Question submitted successfully!"
                    }
                };
            }

            // No matching submission type
            return {
                task: {
                    type: 'message',
                    value: "No question data received."
                }
            };
        } catch (error) {
            this.logger.error('Error in task module submission', {
                error: {
                    message: error.message,
                    stack: error.stack,
                    name: error.name
                }
            });
            return {
                task: {
                    type: 'message',
                    value: "An error occurred. Please try again."
                }
            };
        }
    }

    private async handleAnswerAcceptance(context: TurnContext, taskModuleRequest: any): Promise<void> {
        try {
            const questionData = taskModuleRequest.data.data.question;
            const answers = taskModuleRequest.data.data.answers;

            this.logger.info('Handling answer acceptance with data:', {
                questionData,
                answers
            });

            // Try to delete the previous card if it exists
            const previousCardId = QABot.questionCardIds.get(questionData.id);
            if (previousCardId) {
                try {
                    await context.deleteActivity(previousCardId);
                    this.logger.info('Successfully deleted previous question card', { 
                        questionId: questionData.id,
                        activityId: previousCardId 
                    });
                } catch (error) {
                    this.logger.warn('Failed to delete previous question card', { 
                        error, 
                        questionId: questionData.id,
                        activityId: previousCardId 
                    });
                }
            }

            // Get fresh data from the database to ensure we have the latest state
            const updatedQuestion = await questionsService.getQuestion(questionData.id);
            if (!updatedQuestion) {
                throw new Error('Question not found');
            }

            this.logger.info('Retrieved updated question data:', {
                updatedQuestion
            });

            // Create and send the updated card
            const card = this.createQuestionCard(updatedQuestion, updatedQuestion.answers);
            const messageActivity = {
                type: ActivityTypes.Message,
                attachments: [card],
                channelData: {
                    ...context.activity.channelData,
                    notification: {
                        alert: true,
                        alertInMeeting: true
                    }
                }
            };

            const cardResponse = await context.sendActivity(messageActivity);
            if (cardResponse?.id) {
                QABot.questionCardIds.set(questionData.id, cardResponse.id);
                this.logger.info('Stored new question card activity ID', { 
                    questionId: questionData.id,
                    activityId: cardResponse.id 
                });
            }
        } catch (error) {
            this.logger.error('Error handling answer acceptance', { error });
        }
    }

    private async handleAnswerSubmission(context: TurnContext, taskModuleRequest: any): Promise<void> {
        try {
            const questionId = taskModuleRequest.data.data.questionId;
            
            // Fetch the updated question data
            const questionData = await questionsService.getQuestion(questionId);
            if (!questionData) {
                throw new Error('Question not found');
            }

            // Try to delete the previous card if it exists
            const previousCardId = QABot.questionCardIds.get(questionId);
            if (previousCardId) {
                try {
                    await context.deleteActivity(previousCardId);
                    this.logger.info('Successfully deleted previous question card', { 
                        questionId: questionId,
                        activityId: previousCardId 
                    });
                } catch (error) {
                    this.logger.warn('Failed to delete previous question card', { 
                        error, 
                        questionId: questionId,
                        activityId: previousCardId 
                    });
                }
            }

            // Create and send the updated card
            const card = this.createQuestionCard(questionData, questionData.answers);
            const messageActivity = {
                type: ActivityTypes.Message,
                attachments: [card],
                channelData: {
                    ...context.activity.channelData,
                    notification: {
                        alert: true,
                        alertInMeeting: true
                    }
                }
            };

            const cardResponse = await context.sendActivity(messageActivity);
            if (cardResponse?.id) {
                QABot.questionCardIds.set(questionId, cardResponse.id);
                this.logger.info('Stored new question card activity ID', { 
                    questionId: questionId,
                    activityId: cardResponse.id 
                });
            }
        } catch (error) {
            this.logger.error('Error handling answer submission', { error });
            throw error;
        }
    }

    private async handleQuestionSubmission(context: TurnContext, taskModuleRequest: any): Promise<void> {
        if (!taskModuleRequest?.data) {
            throw new Error('No data received in task module request');
        }

        // Create a new question object
        const question = {
            id: Math.random().toString(36).substring(7),
            title: taskModuleRequest.data.title,
            text: taskModuleRequest.data.text,
            userName: taskModuleRequest.data.userName || context.activity.from.name,
            answers: [],
            createdAt: new Date(),
            updatedAt: new Date()
        };

        // Validate required fields
        if (!question.title || !question.text) {
            throw new Error('Question title and text are required');
        }

        try {
            // Save the question locally and post to external API
            const savedQuestion = await questionsService.addQuestion(
                question.title,
                question.text,
                question.userName
            );
            
            // Try to delete the ask card
            const lastAskCardId = QuestionDialog.getLastAskCardId();
            if (lastAskCardId) {
                try {
                    await context.deleteActivity(lastAskCardId);
                    this.logger.info('Successfully deleted ask card', { 
                        activityId: lastAskCardId,
                        questionId: savedQuestion.id 
                    });
                } catch (error) {
                    this.logger.warn('Failed to delete ask card', { 
                        error: error.message, 
                        activityId: lastAskCardId,
                        questionId: savedQuestion.id
                    });
                }
            }

            function convertHtmlToMarkdown(html: string): { text: string; images: string[] } {
                if (!html) return { text: '', images: [] };
                
                const images: string[] = [];
                const textWithoutImages = html.replace(/<img[^>]*src=["']([^"']*)["'][^>]*>/gi, (match, src) => {
                    images.push(src);
                    return '';
                });
                
                let text = textWithoutImages
                    .replace(/<[^>]+>/g, '')
                    .replace(/\r{3,}/g, '\r\r')
                    .replace(/\s+$/gm, '')
                    .trim();

                return { text, images };
            }

            const { text: convertedText } = convertHtmlToMarkdown(savedQuestion.body);

            // Create the card with the URL if available
            const qaAppUrl = `${process.env.QA_APP_URI}/qna/${savedQuestion.id}`;
            
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
                                text: "✨ New Question Posted",
                                size: "Small",
                                weight: "Bolder",
                                color: "Accent",
                                spacing: "Small"
                            },
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
                                text: `Posted by ${savedQuestion.profile.firstName} ${savedQuestion.profile.lastName} • ${new Date(savedQuestion.createdAt).toLocaleString()}`,
                                wrap: true,
                                size: "Small",
                                isSubtle: true
                            }
                        ]
                    }
                ].filter(Boolean),
                actions: [
                    {
                        type: "Action.OpenUrl",
                        title: "Open in Q&A App",
                        url: qaAppUrl,
                        style: "positive"
                    },
                    {
                        type: "Action.Submit",
                        title: "View Full Question",
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
        } catch (error) {
            this.logger.error('Error in handleQuestionSubmission', {
                error: {
                    message: error.message,
                    stack: error.stack
                }
            });
            throw error;
        }
    }

    private createQuestionCard(questionData: any, answers: any[]): any {
        const qaAppUrl = `${process.env.QA_APP_URI}/qna/${questionData.id}`;
        
        return CardFactory.adaptiveCard({
            type: "AdaptiveCard",
            version: "1.2",
            $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
            msteams: {
                width: "Full"
            },
            body: [
                {
                    type: "Container",
                    style: "emphasis",
                    items: [
                        {
                            type: "TextBlock",
                            text: answers.length > 0 ? "✨ Question" : "✨ New Question Posted",
                            size: "Small",
                            weight: "Bolder",
                            color: "Accent",
                            spacing: "Small"
                        },
                        {
                            type: "TextBlock",
                            text: questionData.title,
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
                            text: `Posted by ${questionData.profile.firstName} ${questionData.profile.lastName} • ${new Date(questionData.createdAt).toLocaleString()}`,
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
                            text: questionData.body.replace(/<[^>]*>/g, ''),
                            wrap: true,
                            size: "Medium",
                            spacing: "Medium"
                        }
                    ]
                },
                ...(answers.length > 0 ? [
                    {
                        type: "Container",
                        spacing: "ExtraLarge",
                        separator: true,
                        items: [
                            {
                                type: "TextBlock",
                                text: `💬 Answers (${answers.length})`,
                                wrap: true,
                                weight: "Bolder",
                                size: "Medium",
                                color: "Accent",
                                spacing: "Medium"
                            }
                        ]
                    },
                    ...answers.map((answer: any, index: number) => ({
                        type: "Container",
                        spacing: index === 0 ? "Medium" : "Large",
                        style: answer.accepted ? "emphasis" : "default",
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
                ] : [])
            ],
            actions: [
                {
                    type: "Action.OpenUrl",
                    title: "Open in Q&A App",
                    url: qaAppUrl,
                    style: "positive"
                },
                {
                    type: "Action.Submit",
                    title: answers.length > 0 ? "View Full Question & Answers" : "View Full Question",
                    data: {
                        msteams: {
                            type: "task/fetch"
                        },
                        taskModule: "viewQuestion",
                        questionId: questionData.id
                    }
                }
            ]
        });
    }

    public async run(context: TurnContext): Promise<void> {
        await super.run(context);
        // Save state changes
        await this.conversationState.saveChanges(context);
    }
} 