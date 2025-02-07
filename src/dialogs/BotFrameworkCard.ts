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

import {
    ComponentDialog,
    WaterfallDialog,
    WaterfallStepContext
} from 'botbuilder-dialogs';
import { CardFactory } from 'botbuilder';
import * as constants from "../constants";
import * as utils from "../utils";
import { cardTemplates, fetchTemplates, appRoot } from "./CardTemplates";
import { taskModuleLink } from "../utils/DeepLinks";
import { renderCard } from "../utils/CardUtils";
import { Logger } from 'winston';

const BFCARD_DIALOG = 'BFCARD_DIALOG';
const BFCARD_WATERFALL_DIALOG = 'BFCARD_WATERFALL_DIALOG';

export class BotFrameworkCard extends ComponentDialog {
    private logger: Logger;

    constructor(dialogId: string, logger: Logger) {
        super(dialogId);
        
        this.logger = logger;

        // Add waterfall dialog
        this.addDialog(new WaterfallDialog(BFCARD_WATERFALL_DIALOG, [
            this.initialStep.bind(this),
            this.finalStep.bind(this)
        ]));

        // Set the initial dialog to run
        this.initialDialogId = BFCARD_WATERFALL_DIALOG;
    }

    private async initialStep(stepContext: WaterfallStepContext) {
        this.logger.info('BotFrameworkCard dialog - Initial step');
        
        // Message might contain @mentions which we would like to strip off in the response
        let text = utils.getTextWithoutMentions(stepContext.context.activity);

        let appInfo = {
            appId: process.env.MICROSOFT_APP_ID,
        };

        let taskModuleUrls = {
            url1: taskModuleLink(
                appInfo.appId, 
                constants.TaskModuleStrings.CustomFormTipTapTitle, 
                constants.TaskModuleSizes.customformtiptap.height, 
                constants.TaskModuleSizes.customformtiptap.width, 
                `${appRoot()}/${constants.TaskModuleIds.CustomFormTipTap}`
            )
        };

        let cardData: any = {
            title: "Q&A Bot",
            subTitle: "Ask a Question",
            instructions: "Click the button below to ask a question using our rich text editor.",
            linkbutton1: constants.TaskModuleStrings.CustomFormTipTapName,
            url1: taskModuleUrls.url1,
            fetchButtonId1: constants.TaskModuleIds.CustomFormTipTap,
            fetchButtonTitle1: constants.TaskModuleStrings.CustomFormTipTapName
        };

        if (text === constants.DialogId.BFCard) {
            await stepContext.context.sendActivity({ 
                attachments: [renderCard(cardTemplates.questionSubmitted, cardData)]
            });
        }

        return await stepContext.next();
    }

    private async finalStep(stepContext: WaterfallStepContext) {
        this.logger.info('BotFrameworkCard dialog - Final step');
        return await stepContext.endDialog();
    }
}
