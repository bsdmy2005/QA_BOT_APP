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

    question.answers.push(answer);
    this.writeQuestions(questions);
    return answer;
  }

  public updateAnswerStatus(questionId: string, answerId: string, isAccepted: boolean): boolean {
    const questions = this.readQuestions();
    const question = questions.find(q => q.id === questionId);
    
    if (!question) {
      return false;
    }

    const answer = question.answers.find(a => a.id === answerId);
    if (!answer) {
      return false;
    }

    answer.isAccepted = isAccepted;
    this.writeQuestions(questions);
    return true;
  }
}

export const questionsService = new QuestionsService(); 