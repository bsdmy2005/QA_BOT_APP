import logger from '../utils/Logger';
import { createQuestionAction, getQuestionAction, createAnswerAction, acceptAnswerAction, getAllQuestionsAction } from '../actions/db/questions-actions';
import { Question, Answer, Profile } from '../db/schema';

export class QuestionsService {
  async addQuestion(
    title: string, 
    text: string, 
    userName: string
  ): Promise<Question & { profile: Profile }> {
    try {
      logger.info('Creating question in database:', { 
        title,
        userName
      });

      const result = await createQuestionAction(title, text, userName);
      
      if (!result.isSuccess) {
        logger.error('Failed to create question:', { 
          error: result.message,
          title,
          userName
        });
        throw new Error(result.message);
      }

      // Get the full question with profile after creation
      const questionWithProfile = await this.getQuestion(result.data.id);
      return questionWithProfile;

    } catch (error) {
      logger.error('Error in addQuestion:', error);
      throw error;
    }
  }

  async getQuestion(id: string): Promise<Question & { answers: (Answer & { profile: Profile })[] } & { profile: Profile }> {
    const result = await getQuestionAction(id);
    if (!result.isSuccess) {
      throw new Error(result.message);
    }
    return result.data;
  }

  async addAnswer(questionId: string, text: string, userName: string): Promise<Answer> {
    const result = await createAnswerAction(questionId, text, userName);
    if (!result.isSuccess) {
      throw new Error(result.message);
    }
    return result.data;
  }

  async acceptAnswer(questionId: string, answerId: string): Promise<boolean> {
    const result = await acceptAnswerAction(questionId, answerId);
    return result.isSuccess;
  }

  async getQuestions(): Promise<Question[]> {
    const result = await getAllQuestionsAction();
    if (!result.isSuccess) {
      throw new Error(result.message);
    }
    return result.data;
  }
}

export const questionsService = new QuestionsService(); 