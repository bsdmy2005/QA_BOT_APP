import { eq, desc } from 'drizzle-orm';
import { db } from '../db/db';
import { questionsTable, profilesTable, answersTable } from '../db/schema';
import logger from '../utils/Logger';

export class QuestionsService {
    async getQuestion(id: string) {
        try {
            const [question] = await db
                .select({
                    questions: questionsTable,
                    profiles: profilesTable
                })
                .from(questionsTable)
                .where(eq(questionsTable.id, id))
                .leftJoin(profilesTable, eq(questionsTable.userId, profilesTable.userId))
                .limit(1);

            if (!question) {
                return null;
            }

            const answers = await db
                .select({
                    answers: answersTable,
                    profiles: profilesTable
                })
                .from(answersTable)
                .where(eq(answersTable.questionId, id))
                .leftJoin(profilesTable, eq(answersTable.userId, profilesTable.userId))
                .orderBy(desc(answersTable.createdAt));

            return {
                ...question.questions,
                profile: question.profiles,
                answers: answers.map(answer => ({
                    ...answer.answers,
                    profile: answer.profiles
                }))
            };
        } catch (error) {
            logger.error('Error getting question:', error);
            throw error;
        }
    }
} 