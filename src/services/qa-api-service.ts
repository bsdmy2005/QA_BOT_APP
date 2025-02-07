import { Logger } from 'winston';
import logger from '../utils/Logger';

interface User {
  id: string;
  name: string;
}

interface Question {
  id: string;
  title: string;
  body: string;
  user: User;
}

interface QuestionRequest {
  data: {
    question: Question;
  };
}

interface QuestionResponse {
  success: boolean;
  data: {
    question: {
      id: string;
      title: string;
      body: string;
      createdAt: string;
      url: string;
    };
  };
}

class QAApiService {
  private apiUrl: string;
  private apiKey: string;
  private logger: Logger;

  constructor() {
    this.apiUrl = process.env.QA_APP_URI;
    this.apiKey = process.env.QA_APP_KEY;
    this.logger = logger;

    if (!this.apiUrl || !this.apiKey) {
      this.logger.warn('QA API configuration missing', {
        hasUrl: !!this.apiUrl,
        hasKey: !!this.apiKey
      });
    }
  }

  async createQuestion(questionData: any): Promise<QuestionResponse | null> {
    try {
      if (!this.apiUrl || !this.apiKey) {
        this.logger.warn('QA API not configured, skipping external API call');
        return null;
      }

      // Format the request payload according to the expected structure
      const payload: QuestionRequest = {
        data: {
          question: {
            id: questionData.id,
            title: questionData.title,
            body: questionData.text, // The HTML content including images
            user: {
              id: questionData.userName,
              name: questionData.userName
            }
          }
        }
      };

      this.logger.info('Sending question to external API', { 
        url: this.apiUrl,
        payload: JSON.stringify(payload)
      });

      const response = await fetch(`${this.apiUrl}/api/questions`, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${this.apiKey}`,
          'Content-Type': 'application/json',
          'Accept': 'application/json'
        },
        body: JSON.stringify(payload)
      });

      const responseText = await response.text();
      
      if (!response.ok) {
        this.logger.error('API error response:', { 
          status: response.status,
          response: responseText 
        });
        throw new Error(`API error: ${response.statusText}`);
      }

      try {
        const responseData: QuestionResponse = JSON.parse(responseText);
        
        this.logger.info('Successfully posted question to external API', {
          questionId: responseData.data.question.id,
          url: responseData.data.question.url,
          status: response.status
        });

        return responseData;
      } catch (parseError) {
        this.logger.error('Error parsing API response:', { 
          error: parseError,
          responseText 
        });
        throw new Error('Invalid JSON response from API');
      }
    } catch (error) {
      this.logger.error('Error posting question to external API', {
        error: {
          message: error.message,
          stack: error.stack
        }
      });
      return null;
    }
  }
}

export const qaApiService = new QAApiService(); 