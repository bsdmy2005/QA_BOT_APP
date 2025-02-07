import { and, eq } from 'drizzle-orm';
import { db } from '../../db/db';
import { 
  questionsTable,
  answersTable,
  profilesTable,
  type Question,
  type Answer,
  type Profile
} from '../../db/schema';
import logger from '../../utils/Logger';
import { ActionState } from '../../types/types';

// Helper function to get user ID from email
async function getUserIdFromEmail(email: string): Promise<string | null> {
  try {
    const profile = await db.query.profilesTable.findFirst({
      where: eq(profilesTable.email, email)
    });
    return profile?.userId || null;
  } catch (error) {
    logger.error('Error getting user ID from email:', error);
    return null;
  }
}

export async function createQuestionAction(
  title: string,
  body: string,
  userEmail: string
): Promise<ActionState<Question>> {
  try {
    const userId = await getUserIdFromEmail(userEmail);
    if (!userId) {
      logger.error('Failed to find user:', { userEmail });
      return {
        isSuccess: false,
        message: 'User not found'
      };
    }

    logger.info('Attempting to insert question:', { title, userId });
    const [question] = await db.insert(questionsTable)
      .values({
        userId,
        title,
        body
      })
      .returning();

    logger.info('Question inserted successfully:', {
      questionId: question.id,
      userId,
      title
    });

    return {
      isSuccess: true,
      message: 'Question created successfully',
      data: question
    };
  } catch (error) {
    logger.error('Error creating question:', {
      error: error.message,
      userEmail,
      title
    });
    return {
      isSuccess: false,
      message: 'Failed to create question'
    };
  }
}

export async function getQuestionAction(
  id: string
): Promise<ActionState<Question & { answers: (Answer & { profile: Profile })[] } & { profile: Profile }>> {
  try {
    const question = await db.query.questionsTable.findFirst({
      where: eq(questionsTable.id, id),
      with: {
        profile: true,
        answers: {
          with: {
            profile: true
          }
        }
      }
    });

    if (!question) {
      return {
        isSuccess: false,
        message: 'Question not found'
      };
    }

    return {
      isSuccess: true,
      message: 'Question retrieved successfully',
      data: question
    };
  } catch (error) {
    logger.error('Error getting question:', error);
    return {
      isSuccess: false,
      message: 'Failed to get question'
    };
  }
}

export async function createAnswerAction(
  questionId: string,
  body: string,
  userEmail: string
): Promise<ActionState<Answer>> {
  try {
    const userId = await getUserIdFromEmail(userEmail);
    if (!userId) {
      return {
        isSuccess: false,
        message: 'User not found'
      };
    }

    const [answer] = await db.insert(answersTable)
      .values({
        questionId,
        userId,
        body
      })
      .returning();

    logger.info('Answer created successfully:', {
      answerId: answer.id,
      questionId,
      userId
    });

    return {
      isSuccess: true,
      message: 'Answer created successfully',
      data: answer
    };
  } catch (error) {
    logger.error('Error creating answer:', error);
    return {
      isSuccess: false,
      message: 'Failed to create answer'
    };
  }
}

export async function acceptAnswerAction(
  questionId: string,
  answerId: string
): Promise<ActionState<void>> {
  try {
    // First reset all answers for this question to not accepted
    await db.update(answersTable)
      .set({ [answersTable.accepted.name]: false })
      .where(eq(answersTable.questionId, questionId));

    // Then set the specific answer as accepted
    await db.update(answersTable)
      .set({ [answersTable.accepted.name]: true })
      .where(and(
        eq(answersTable.id, answerId),
        eq(answersTable.questionId, questionId)
      ));

    logger.info('Answer accepted successfully:', {
      questionId,
      answerId
    });

    return {
      isSuccess: true,
      message: 'Answer accepted successfully',
      data: undefined
    };
  } catch (error) {
    logger.error('Error accepting answer:', error);
    return {
      isSuccess: false,
      message: 'Failed to accept answer'
    };
  }
}

export async function getAllQuestionsAction(): Promise<ActionState<(Question & { profile: Profile })[]>> {
  try {
    const questions = await db.query.questionsTable.findMany({
      orderBy: (questions, { desc }) => [desc(questions.createdAt)],
      with: {
        profile: true
      }
    });

    return {
      isSuccess: true,
      message: 'Questions retrieved successfully',
      data: questions
    };
  } catch (error) {
    logger.error('Error getting questions:', error);
    return {
      isSuccess: false,
      message: 'Failed to get questions'
    };
  }
} 