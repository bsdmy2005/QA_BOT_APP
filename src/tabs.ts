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

import { Request, Response } from "express";
import * as bodyParser from "body-parser";
import { questionsService } from './services/questions-service';
import * as builder from "botbuilder";

module.exports.setup = function(app: any): void {
    let path = require("path");
    let express = require("express");
    let logger = require("winston");

    // Configure the view engine, views folder and the statics path
    app.set("view engine", "pug");
    app.set("views", path.join(__dirname, "../src/views"));

    app.use(bodyParser.json());
    app.use(bodyParser.urlencoded({
        extended: true,
    }));

    // Add request logging middleware
    app.use((req: Request, res: Response, next: any) => {
        logger.info(`Incoming ${req.method} request to ${req.url}`, {
            headers: req.headers,
            query: req.query,
            body: req.body
        });
        next();
    });

    // Setup home page
    app.get("/", function(req: Request, res: Response): void {
        res.render("hello");
    });

    // Setup the static tab
    app.get("/hello", function(req: Request, res: Response): void {
        res.render("hello");
    });

    // Setup the configure tab, with first and second as content tabs
    app.get("/configure", function(req: Request, res: Response): void {
        res.render("configure");
    });

    app.get("/first", function(req: Request, res: Response): void {
        res.render("first");
    });

    app.get("/second", function(req: Request, res: Response): void {
        res.render("second");
    });

    app.get("/taskmodule", function(req: Request, res: Response): void {
        logger.info("Rendering taskmodule view", {
            appId: process.env.MICROSOFT_APP_ID,
            baseUri: process.env.BASE_URI
        });
        // Render the template, passing the appId so it's included in the rendered HTML
        res.render("taskmodule", { 
            appId: process.env.MICROSOFT_APP_ID,
            baseUri: process.env.BASE_URI
        });
    });

    app.get("/youtube", function(req: Request, res: Response): void {
        logger.info("Rendering youtube view");
        try {
            res.render("youtube");
        } catch (error) {
            logger.error("Error rendering youtube view", { error });
            res.status(500).send("Error rendering youtube view");
        }
    });

    app.get("/powerapp", function(req: Request, res: Response): void {
        logger.info("Rendering powerapp view");
        try {
            res.render("powerapp");
        } catch (error) {
            logger.error("Error rendering powerapp view", { error });
            res.status(500).send("Error rendering powerapp view");
        }
    });

    app.get("/customform", function(req: Request, res: Response): void {
        logger.info("Rendering customform view", {
            appId: process.env.MICROSOFT_APP_ID,
            baseUri: process.env.BASE_URI
        });
        try {
            // Get Teams context from the request headers
            const userId = req.headers['x-ms-client-principal-id'] as string || 'anonymous';
            const userName = req.headers['x-ms-client-principal-name'] as string || 'Anonymous User';

            // Render the template with user info
            res.render("customform", { 
                appId: process.env.MICROSOFT_APP_ID,
                baseUri: process.env.BASE_URI,
                userId: userId,
                userName: userName
            });
        } catch (error) {
            logger.error("Error rendering customform view", { error });
            res.status(500).send("Error rendering customform view");
        }
    });

    app.get("/customform-tiptap", function(req: Request, res: Response): void {
        logger.info("Rendering customform-tiptap view", {
            appId: process.env.MICROSOFT_APP_ID,
            baseUri: process.env.BASE_URI
        });
        try {
            // Get Teams context from the request headers
            const userId = req.headers['x-ms-client-principal-id'] as string || 'anonymous';
            const userName = req.headers['x-ms-client-principal-name'] as string || 'Anonymous User';

            // Render the template with user info
            res.render("customform-tiptap", { 
                appId: process.env.MICROSOFT_APP_ID,
                baseUri: process.env.BASE_URI,
                userId: userId,
                userName: userName
            });
        } catch (error) {
            logger.error("Error rendering customform-tiptap view", { error });
            res.status(500).send("Error rendering customform-tiptap view");
        }
    });

    // API endpoint to handle question submission
    app.post("/api/questions", function(req: Request, res: Response): void {
        try {
            const question = req.body;
            logger.info("Received question submission", { question });

            // Generate URL for the question
            question.url = `${process.env.BASE_URI}/qna/${question.id}`;
            
            const savedQuestion = questionsService.addQuestion(question);
            res.status(200).json(savedQuestion);
        } catch (error) {
            logger.error("Error saving question", { error });
            res.status(500).json({ error: "Failed to save question" });
        }
    });

    // API endpoint to get all questions
    app.get("/api/questions", function(req: Request, res: Response): void {
        try {
            const questions = questionsService.getQuestions();
            res.status(200).json(questions);
        } catch (error) {
            logger.error("Error getting questions", { error });
            res.status(500).json({ error: "Failed to get questions" });
        }
    });

    // API endpoint to get a specific question
    app.get("/api/questions/:id", function(req: Request, res: Response): void {
        try {
            const question = questionsService.getQuestion(req.params.id);
            if (!question) {
                res.status(404).json({ error: "Question not found" });
                return;
            }
            res.status(200).json(question);
        } catch (error) {
            logger.error("Error getting question", { error });
            res.status(500).json({ error: "Failed to get question" });
        }
    });

    // API endpoint to add an answer to a question
    app.post("/api/questions/:id/answers", function(req: Request, res: Response): void {
        try {
            // Get Teams context from the request headers with proper fallback
            const userId = (
                req.headers['x-ms-client-principal-id'] || 
                req.headers['userobjectid'] || 
                'anonymous'
            ) as string;
            
            const userName = (
                req.headers['x-ms-client-principal-name'] || 
                req.headers['userprincipalname'] || 
                'Anonymous User'
            ) as string;

            logger.info("Adding answer with user context", { 
                userId, 
                userName,
                headers: req.headers 
            });

            // Create answer with user info from headers
            const answer = {
                ...req.body,
                userId: userId,
                userName: userName
            };

            const savedAnswer = questionsService.addAnswer(req.params.id, answer);
            if (!savedAnswer) {
                res.status(404).json({ error: "Question not found" });
                return;
            }
            res.status(200).json(savedAnswer);
        } catch (error) {
            logger.error("Error saving answer", { error });
            res.status(500).json({ error: "Failed to save answer" });
        }
    });

    // API endpoint to accept an answer
    app.put("/api/questions/:questionId/answers/:answerId/accept", function(req: Request, res: Response): void {
        try {
            // Get Teams context from the request headers with proper fallback
            const userId = (
                req.headers['x-ms-client-principal-id'] || 
                req.headers['userobjectid'] || 
                'anonymous'
            ) as string;

            logger.info("Accepting answer with user context", { 
                userId,
                questionId: req.params.questionId,
                answerId: req.params.answerId,
                headers: req.headers
            });
            
            // Get the question to verify ownership
            const question = questionsService.getQuestion(req.params.questionId);
            if (!question) {
                logger.warn("Question not found for accept answer", {
                    questionId: req.params.questionId
                });
                res.status(404).json({ error: "Question not found" });
                return;
            }

            // Verify the user is the question owner
            if (question.userId !== userId) {
                logger.warn("Unauthorized accept answer attempt", {
                    questionUserId: question.userId,
                    requestUserId: userId
                });
                res.status(403).json({ error: "Only the question owner can accept answers" });
                return;
            }

            const success = questionsService.acceptAnswer(
                req.params.questionId,
                req.params.answerId
            );

            if (!success) {
                logger.warn("Answer not found for accept", {
                    questionId: req.params.questionId,
                    answerId: req.params.answerId
                });
                res.status(404).json({ error: "Answer not found" });
                return;
            }

            // Get the updated question to return
            const updatedQuestion = questionsService.getQuestion(req.params.questionId);
            logger.info("Successfully accepted answer", {
                questionId: req.params.questionId,
                answerId: req.params.answerId,
                updatedQuestion
            });
            res.status(200).json(updatedQuestion);
        } catch (error) {
            logger.error("Error accepting answer", { error });
            res.status(500).json({ error: "Failed to accept answer" });
        }
    });

    app.post("/register", function(req: Request, res: Response): void {
        logger.info("Received form submission", {
            name: req.body.name,
            email: req.body.email,
            favoriteBook: req.body.favoriteBook
        });
        res.status(200).send("Registration successful");
    });

    // Route to display questions
    app.get("/questions", function(req: Request, res: Response): void {
        logger.info("Rendering questions view");
        try {
            // Get Teams context from the request headers
            const userId = req.headers['x-ms-client-principal-id'] as string || 'anonymous';
            const userName = req.headers['x-ms-client-principal-name'] as string || 'Anonymous User';

            // Get all questions
            const questions = questionsService.getQuestions();

            // Render the template with questions and user info
            res.render("questions", { 
                questions,
                userId: userId,
                userName: userName
            });
        } catch (error) {
            logger.error("Error rendering questions view", { error });
            res.status(500).send("Error rendering questions view");
        }
    });

    // Route to display a specific question
    app.get("/qna/:id", function(req: Request, res: Response): void {
        logger.info("Rendering single question view");
        try {
            // Get Teams context from the request headers
            const userId = req.headers['x-ms-client-principal-id'] as string || 'anonymous';
            const userName = req.headers['x-ms-client-principal-name'] as string || 'Anonymous User';

            // Get the question
            const question = questionsService.getQuestion(req.params.id);
            if (!question) {
                res.status(404).send("Question not found");
                return;
            }

            // Render the template with the question and user info
            res.render("questions", { 
                questions: [question],
                userId: userId,
                userName: userName
            });
        } catch (error) {
            logger.error("Error rendering single question view", { error });
            res.status(500).send("Error rendering single question view");
        }
    });

    // Route to display a specific question in a task module
    app.get("/question/:id", function(req: Request, res: Response): void {
        logger.info("Rendering question details view");
        try {
            // Get Teams context from the request headers with proper fallback
            const userId = (
                req.headers['x-ms-client-principal-id'] || 
                req.headers['userobjectid'] || 
                'anonymous'
            ) as string;
            
            const userName = (
                req.headers['x-ms-client-principal-name'] || 
                req.headers['userprincipalname'] || 
                'Anonymous User'
            ) as string;

            logger.info("User context for question view", { 
                userId, 
                userName,
                headers: req.headers,
                allHeaders: JSON.stringify(req.headers)
            });

            // Get the question
            const question = questionsService.getQuestion(req.params.id);
            if (!question) {
                res.status(404).send("Question not found");
                return;
            }

            // Check if current user is the question owner
            const isQuestionOwner = question.userId === userId;
            logger.info("Question ownership check", { 
                isQuestionOwner, 
                questionUserId: question.userId, 
                currentUserId: userId,
                question: JSON.stringify(question),
                userIdType: typeof userId,
                questionUserIdType: typeof question.userId,
                exactMatch: `${userId} === ${question.userId}`
            });

            // Render the template with the question and user info
            res.render("question", { 
                question,
                baseUri: process.env.BASE_URI,
                userId: userId,
                userName: userName,
                isQuestionOwner: isQuestionOwner
            });
        } catch (error) {
            logger.error("Error rendering question details view", { error });
            res.status(500).send("Error rendering question details view");
        }
    });
};
