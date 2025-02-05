import * as fs from 'fs';
import * as path from 'path';
import * as logger from 'winston';

interface Answer {
  id: string;
  text: string;
  userId: string;
  userName: string;
  timestamp: string;
  isAccepted: boolean;
}

interface Question {
  id: string;
  title: string;
  text: string;
  userId: string;
  userName: string;
  answers: Answer[];
  timestamp: string;
  url?: string;
}

class QuestionsService {
  private questionsFile = path.join(__dirname, '../../data/questions.json');

  constructor() {
    // Ensure questions.json exists
    if (!fs.existsSync(this.questionsFile)) {
      fs.writeFileSync(this.questionsFile, '[]');
    }
  }

  private readQuestions(): Question[] {
    try {
      const data = fs.readFileSync(this.questionsFile, 'utf8');
      return JSON.parse(data);
    } catch (error) {
      logger.error('Error reading questions file:', error);
      return [];
    }
  }

  private writeQuestions(questions: Question[]): void {
    try {
      fs.writeFileSync(this.questionsFile, JSON.stringify(questions, null, 2));
    } catch (error) {
      logger.error('Error writing questions file:', error);
    }
  }

  public addQuestion(question: Question): Question {
    const questions = this.readQuestions();
    questions.push(question);
    this.writeQuestions(questions);
    return question;
  }

  public getQuestions(): Question[] {
    return this.readQuestions();
  }

  public getQuestion(id: string): Question | undefined {
    const questions = this.readQuestions();
    return questions.find(q => q.id === id);
  }

  public addAnswer(questionId: string, answer: Answer): Answer | null {
    const questions = this.readQuestions();
    const question = questions.find(q => q.id === questionId);
    
    if (!question) {
      return null;
    }

    // Initialize answer with isAccepted property
    const newAnswer = {
      ...answer,
      id: Math.random().toString(36).substring(7), // Generate a simple ID
      isAccepted: false
    };

    question.answers.push(newAnswer);
    this.writeQuestions(questions);
    return newAnswer;
  }

  public acceptAnswer(questionId: string, answerId: string): boolean {
    const questions = this.readQuestions();
    const question = questions.find(q => q.id === questionId);
    
    if (!question) {
      logger.warn('Question not found for accept answer', { questionId });
      return false;
    }

    logger.info('Found question for accept answer', { 
      questionId,
      answerId,
      currentAnswers: question.answers.map(a => ({ id: a.id, isAccepted: a.isAccepted }))
    });

    // First, set all answers to not accepted
    question.answers.forEach(answer => {
      answer.isAccepted = false;
    });

    // Then find and accept the specific answer
    const answer = question.answers.find(a => a.id === answerId);
    if (!answer) {
      logger.warn('Answer not found in question', { 
        questionId, 
        answerId,
        availableAnswerIds: question.answers.map(a => a.id)
      });
      return false;
    }

    answer.isAccepted = true;
    this.writeQuestions(questions);
    logger.info('Successfully accepted answer', {
      questionId,
      answerId,
      updatedAnswers: question.answers.map(a => ({ id: a.id, isAccepted: a.isAccepted }))
    });
    return true;
  }
}

export const questionsService = new QuestionsService(); 