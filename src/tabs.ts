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

    // Setup home page to show questions list
    app.get("/", function(req: Request, res: Response): void {
        res.render("questions");
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
    app.post("/api/questions", async function(req: Request, res: Response): Promise<void> {
        try {
            const { title, text, userName } = req.body;
            logger.info("Received question submission", { title, userName });

            // Generate URL for the question
            const savedQuestion = await questionsService.addQuestion(title, text, userName);
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
    app.post("/api/questions/:id/answers", async function(req: Request, res: Response): Promise<void> {
        try {
            const userName = req.headers['x-ms-client-principal-name'] as string || 'Anonymous User';
            const text = req.body.text;

            logger.info("Adding answer", { 
                questionId: req.params.id,
                userName
            });

            const savedAnswer = await questionsService.addAnswer(req.params.id, text, userName);
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
    app.put("/api/questions/:questionId/answers/:answerId/accept", async function(req: Request, res: Response): Promise<void> {
        try {
            const userName = req.headers['x-ms-client-principal-name'] as string || 'Anonymous User';

            logger.info("Accepting answer", { 
                questionId: req.params.questionId,
                answerId: req.params.answerId,
                userName
            });
            
            // Get the question
            const question = await questionsService.getQuestion(req.params.questionId);
            if (!question) {
                logger.warn("Question not found for accept answer", {
                    questionId: req.params.questionId
                });
                res.status(404).json({ error: "Question not found" });
                return;
            }

            const success = await questionsService.acceptAnswer(
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
            const updatedQuestion = await questionsService.getQuestion(req.params.questionId);
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
    app.get("/question/:id", async function(req: Request, res: Response): Promise<void> {
        logger.info("Rendering question details view");
        try {
            const userName = req.headers['x-ms-client-principal-name'] as string || 'Anonymous User';

            logger.info("User context for question view", { 
                userName,
                headers: req.headers
            });

            // Get the question
            const question = await questionsService.getQuestion(req.params.id);
            if (!question) {
                res.status(404).send("Question not found");
                return;
            }

            // Render the template with the question and user info
            res.render("question", { 
                question,
                baseUri: process.env.BASE_URI,
                userName: userName
            });
        } catch (error) {
            logger.error("Error rendering question details view", { error });
            res.status(500).send("Error rendering question details view");
        }
    });
};
