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

import * as constants from "../constants";
import { renderACAttachment } from "../utils";

// Function that works both in Node (where window === undefined) or the browser
export function appRoot(): string {
    if (typeof window === "undefined") {
        return process.env.BASE_URI;
    } else {
        return window.location.protocol + "//" + window.location.host;
    }
}

// tslint:disable:trailing-comma
export const cardTemplates: any = {
    questionSubmitted: {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "TextBlock",
                "size": "Medium",
                "weight": "Bolder",
                "text": "New Question Submitted"
            },
            {
                "type": "TextBlock",
                "text": "Title: {{title}}",
                "wrap": true,
                "weight": "Bolder"
            },
            {
                "type": "TextBlock",
                "text": "Asked by: {{userName}}",
                "wrap": true,
                "isSubtle": true
            },
            {
                "type": "TextBlock",
                "text": "{{timestamp}}",
                "wrap": true,
                "isSubtle": true,
                "spacing": "None"
            }
        ],
        "version": "1.0"
    }
};

export const fetchTemplates: any = {
    customformtiptap: {
        title: "Ask a Question",
        height: 1020,
        width: 1632,
        url: `${appRoot()}/customform-tiptap`,
        fallbackUrl: `${appRoot()}/customform-tiptap`,
        completionBotId: process.env.MICROSOFT_APP_ID
    },
    submitMessageResponse: {
        type: "message",
        text: "Thanks for your submission!"
    },
    // currently required until null response supported
    submitNullResponse: {
        "task": {
            "type": "message",
            "value": "",
        },
    },
};
