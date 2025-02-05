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
    taskModule: {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "size": "Medium",
                        "weight": "Bolder",
                        "text": "{{title}}"
                    },
                    {
                        "type": "ColumnSet",
                        "columns": [
                            {
                                "type": "Column",
                                "items": [
                                    {
                                        "type": "TextBlock",
                                        "weight": "Bolder",
                                        "text": "{{subTitle}}",
                                        "wrap": true
                                    }
                                ],
                                "width": "stretch"
                            }
                        ]
                    }
                ]
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "{{instructions}}",
                        "wrap": true
                    }
                ]
            }
        ],
        "actions": [
            {
                "type": "Action.ShowCard",
                "title": "Deep Links",
                "card": {
                    "type": "AdaptiveCard",
                    "style": "emphasis",
                    "body": [
                        {
                            "type": "FactSet",
                            "facts": [
                                {
                                    "title": "Button 1 URL",
                                    "value": "{{markdown1}}"
                                },
                                {
                                    "title": "Button 2 URL",
                                    "value": "{{markdown2}}"
                                },
                                {
                                    "title": "Button 3 URL",
                                    "value": "{{markdown3}}"
                                },
                                {
                                    "title": "Button 4 URL",
                                    "value": "{{markdown4}}"
                                },
                                {
                                    "title": "Button 5 URL",
                                    "value": "{{markdown5}}"
                                }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "text": "Click on the buttons below below to open a task module via deep link.",
                            "wrap": true
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "{{linkbutton1}}",
                            "url": "{{url1}}"
                        },
                        {
                            "type": "Action.OpenUrl",
                            "title": "{{linkbutton2}}",
                            "url": "{{url2}}"
                        },
                        {
                            "type": "Action.OpenUrl",
                            "title": "{{linkbutton3}}",
                            "url": "{{url3}}"
                        },
                        {
                            "type": "Action.OpenUrl",
                            "title": "{{linkbutton4}}",
                            "url": "{{url4}}"
                        },
                        {
                            "type": "Action.OpenUrl",
                            "title": "{{linkbutton5}}",
                            "url": "{{url5}}"
                        }
                    ]
                }
            },
            {
                "type": "Action.ShowCard",
                "title": "task/fetch",
                "card": {
                    "type": "AdaptiveCard",
                    "style": "emphasis",
                    "body": [
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "{{tfJsonTitle1}}"
                        },
                        {
                            "type": "TextBlock",
                            "spacing": "None",
                            "height": "stretch",
                            "text": "{{tfJson1}}",
                            "isSubtle": true,
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "{{tfJsonTitle2}}"
                        },
                        {
                            "type": "TextBlock",
                            "spacing": "None",
                            "height": "stretch",
                            "text": "{{tfJson2}}",
                            "isSubtle": true,
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "weight": "Bolder",
                            "text": "{{tfJsonTitle3}}"
                        },
                        {
                            "type": "TextBlock",
                            "spacing": "None",
                            "height": "stretch",
                            "text": "{{tfJson3}}",
                            "isSubtle": true,
                            "wrap": true
                        },
                        {
                            "type": "TextBlock",
                            "text": "Click on the buttons below below to open a task module via task/fetch.",
                            "wrap": true
                        }
                    ],
                    "actions": [
                        {
                            "type": "Action.Submit",
                            "id": "{{fetchButtonId1}}",
                            "title": "{{fetchButtonTitle1}}",
                            "data": {
                                "msteams": {
                                    "type": "task/fetch"
                                },
                                "taskModule": "{{fetchButtonId1}}"
                            }
                        },
                        {
                            "type": "Action.Submit",
                            "id": "{{fetchButtonId2}}",
                            "title": "{{fetchButtonTitle2}}",
                            "data": {
                                "msteams": {
                                    "type": "task/fetch"
                                },
                                "taskModule": "{{fetchButtonId2}}"
                            }
                        },
                        {
                            "type": "Action.Submit",
                            "id": "{{fetchButtonId3}}",
                            "title": "{{fetchButtonTitle3}}",
                            "data": {
                                "msteams": {
                                    "type": "task/fetch"
                                },
                                "taskModule": "{{fetchButtonId3}}"
                            }
                        },
                        {
                            "type": "Action.Submit",
                            "id": "{{fetchButtonId4}}",
                            "title": "{{fetchButtonTitle4}}",
                            "data": {
                                "msteams": {
                                    "type": "task/fetch"
                                },
                                "taskModule": "{{fetchButtonId4}}"
                            }
                        },
                        {
                            "type": "Action.Submit",
                            "id": "{{fetchButtonId5}}",
                            "title": "{{fetchButtonTitle5}}",
                            "data": {
                                "msteams": {
                                    "type": "task/fetch"
                                },
                                "taskModule": "{{fetchButtonId5}}"
                            }
                        }
                    ]
                }
            }
        ],
        "version": "1.0"
    },
    acTester: {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "size": "Medium",
                        "weight": "Bolder",
                        "text": "Adaptive Card Tester"
                    }
                ]
            },
            {
                "type": "Container",
                "items": [
                    {
                        "type": "TextBlock",
                        "text": "Copy/paste Adaptive Card JSON from [Adaptive Cards Designer](http://acdesignerbeta.azurewebsites.net/) or [Adaptive Cards Samples](http://adaptivecards.io/samples/) into the text box below and press Submit. You can also copy/paste O365Connector/MessageCard JSON from [Message Card Playground](https://messagecardplayground.azurewebsites.net/) or Bot Framework card JSON.",
                        "wrap": true
                    }
                ]
            },
            {
                "type": "Input.Text",
                "id": "acBody",
                "title": "New Input.Toggle",
                "placeholder": "Adaptive Card JSON",
                "isMultiline": true
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "Submit",
                "data": {
                    "actester": true
                }
            }
        ],
        "version": "1.0"
    },
    adaptiveCard: {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "TextBlock",
                "separator": true,
                "size": "Large",
                "weight": "Bolder",
                "text": "Enter basic information for this position:",
                "isSubtle": true,
                "wrap": true
            },
            {
                "type": "TextBlock",
                "separator": true,
                "text": "Title",
                "wrap": true
            },
            {
                "type": "Input.Text",
                "id": "jobTitle",
                "placeholder": "E.g. Senior PM"
            },
            {
                "type": "ColumnSet",
                "columns": [
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "Level",
                                "wrap": true
                            },
                            {
                                "type": "Input.Number",
                                "id": "jobLevel",
                                "value": "7",
                                "placeholder": "Level",
                                "min": 7,
                                "max": 10
                            }
                        ],
                        "width": 2
                    },
                    {
                        "type": "Column",
                        "items": [
                            {
                                "type": "TextBlock",
                                "text": "Location"
                            },
                            {
                                "type": "Input.ChoiceSet",
                                "id": "jobLocation",
                                "value": "1",
                                "choices": [
                                    {
                                        "title": "San Francisco",
                                        "value": "1"
                                    },
                                    {
                                        "title": "London",
                                        "value": "2"
                                    },
                                    {
                                        "title": "Singapore",
                                        "value": "3"
                                    },
                                    {
                                        "title": "Dubai",
                                        "value": "3"
                                    },
                                    {
                                        "title": "Frankfurt",
                                        "value": "3"
                                    }
                                ],
                                "isCompact": true
                            }
                        ],
                        "width": 2
                    }
                ]
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "id": "createPosting",
                "title": "Create posting",
                "data": {
                    "command": "createPosting",
                    "taskResponse": "{{responseType}}",
                }
            },
            {
                "type": "Action.Submit",
                "id": "cancel",
                "title": "Cancel"
            }
        ],
        "version": "1.0"
    },
    adaptiveCardKitchenSink: {
        "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
        "type": "AdaptiveCard",
        "version": "1.0",
        "body": [
          {
            "type": "TextBlock",
            "size": "medium",
            "weight": "bolder",
            "text": "Input.Text elements",
            "horizontalAlignment": "center"
          },
          {
            "type": "Input.Text",
            "placeholder": "Name",
            "style": "text",
            "maxLength": 0,
            "id": "SimpleVal"
          },
          {
            "type": "Input.Text",
            "placeholder": "Homepage",
            "style": "url",
            "maxLength": 0,
            "id": "UrlVal"
          },
          {
            "type": "Input.Text",
            "placeholder": "Email",
            "style": "email",
            "maxLength": 0,
            "id": "EmailVal"
          },
          {
            "type": "Input.Text",
            "placeholder": "Phone",
            "style": "tel",
            "maxLength": 0,
            "id": "TelVal"
          },
          {
            "type": "Input.Text",
            "placeholder": "Comments",
            "style": "text",
            "isMultiline": true,
            "maxLength": 0,
            "id": "MultiLineVal"
          },
          {
            "type": "Input.Number",
            "placeholder": "Quantity",
            "min": -5,
            "max": 5,
            "value": 1,
            "id": "NumVal"
          },
          {
            "type": "Input.Date",
            "placeholder": "Due Date",
            "id": "DateVal",
            "value": "2017-09-20"
          },
          {
            "type": "Input.Time",
            "placeholder": "Start time",
            "id": "TimeVal",
            "value": "16:59"
          },
          {
            "type": "TextBlock",
            "size": "medium",
            "weight": "bolder",
            "text": "Input.ChoiceSet",
            "horizontalAlignment": "center"
          },
          {
            "type": "TextBlock",
            "text": "What color do you want? (compact)"
          },
          {
            "type": "Input.ChoiceSet",
            "id": "CompactSelectVal",
            "style": "compact",
            "value": "1",
            "choices": [
              {
                "title": "Red",
                "value": "1"
              },
              {
                "title": "Green",
                "value": "2"
              },
              {
                "title": "Blue",
                "value": "3"
              }
            ]
          },
          {
            "type": "TextBlock",
            "text": "What color do you want? (expanded)"
          },
          {
            "type": "Input.ChoiceSet",
            "id": "SingleSelectVal",
            "style": "expanded",
            "value": "1",
            "choices": [
              {
                "title": "Red",
                "value": "1"
              },
              {
                "title": "Green",
                "value": "2"
              },
              {
                "title": "Blue",
                "value": "3"
              }
            ]
          },
          {
            "type": "TextBlock",
            "text": "What colors do you want? (multiselect)"
          },
          {
            "type": "Input.ChoiceSet",
            "id": "MultiSelectVal",
            "isMultiSelect": true,
            "value": "1,3",
            "choices": [
              {
                "title": "Red",
                "value": "1"
              },
              {
                "title": "Green",
                "value": "2"
              },
              {
                "title": "Blue",
                "value": "3"
              }
            ]
          },
          {
            "type": "TextBlock",
            "size": "medium",
            "weight": "bolder",
            "text": "Input.Toggle",
            "horizontalAlignment": "center"
          },
          {
            "type": "Input.Toggle",
            "title": "I accept the terms and conditions (True/False)",
            "valueOn": "true",
            "valueOff": "false",
            "id": "AcceptsTerms"
          },
          {
            "type": "Input.Toggle",
            "title": "Red cars are better than other cars",
            "valueOn": "RedCars",
            "valueOff": "NotRedCars",
            "id": "ColorPreference"
          }
        ],
        "actions": [
          {
            "type": "Action.Submit",
            "title": "Submit",
            "data": {
              "id": "1234567890",
              "taskResponse": "{{responseType}}",
            }
          },
          {
            "type": "Action.ShowCard",
            "title": "Show Card",
            "card": {
              "type": "AdaptiveCard",
              "body": [
                {
                  "type": "Input.Text",
                  "placeholder": "enter comment",
                  "style": "text",
                  "maxLength": 0,
                  "id": "CommentVal"
                }
              ],
              "actions": [
                {
                  "type": "Action.Submit",
                  "title": "OK",
                  "data": {
                    "taskResponse": "{{responseType}}",
                  }
                }
              ]
            }
          }
        ]
    },
    acSubmitResponse: {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "TextBlock",
                "weight": "Bolder",
                "text": "Action.Submit Results"
            },
            {
                "type": "TextBlock",
                "separator": true,
                "size": "Medium",
                "text": "{{results}}",
                "wrap": true
            }
        ],
        "actions": [
            {
                "type": "Action.Submit",
                "title": "OK",
                "data": {
                    "taskResponse": "final",
                    "taskModule": "acResponse"
                }
            }
        ],
        "version": "1.0"
    },
    bfThumbnailDeepLink: {
        "contentType": "application/vnd.microsoft.card.thumbnail",
        "content": {
          "title": "{{title}}",
          "subtitle": "{{subTitleDL}}",
          "text": "{{instructionsDL}}",
          "images": [
            {
              "url": "https://upload.wikimedia.org/wikipedia/commons/thumb/8/81/Wiki_puzzle_piece_blank.svg/951px-Wiki_puzzle_piece_blank.svg.png",
              "alt": "Task Module Puzzle Piece"
            }
          ],
          "buttons": [
            {
                "type": "openUrl",
                "title": "{{linkbutton1}}",
                "value": "{{url1}}"
            },
            {
                "type": "openUrl",
                "title": "{{linkbutton2}}",
                "value": "{{url2}}"
            },
            {
                "type": "openUrl",
                "title": "{{linkbutton3}}",
                "value": "{{url3}}"
            },
            {
                "type": "openUrl",
                "title": "{{linkbutton4}}",
                "value": "{{url4}}"
            },
            {
                "type": "openUrl",
                "title": "{{linkbutton5}}",
                "value": "{{url5}}"
            },
          ],
        }
    },
    bfThumbnailTaskFetch: {
        "contentType": "application/vnd.microsoft.card.thumbnail",
        "content": {
          "title": "{{title}}",
          "subtitle": "{{subTitleTF}}",
          "text": "{{instructionsTF}}",
          "images": [
            {
              "url": "https://upload.wikimedia.org/wikipedia/commons/thumb/8/81/Wiki_puzzle_piece_blank.svg/951px-Wiki_puzzle_piece_blank.svg.png",
              "alt": "Task Module Puzzle Piece"
            }
          ],
          "buttons": [
            {
                "type": "invoke",
                "title": "{{fetchButtonTitle1}}",
                "value": {
                    "msteams": {
                        "type": "task/fetch"
                    },
                    "data": {
                        "taskModule": "{{fetchButtonId1}}"
                    }
                }
            },
            {
                "type": "invoke",
                "title": "{{fetchButtonTitle2}}",
                "value": {
                    "msteams": {
                        "type": "task/fetch"
                    },
                    "data": {
                        "taskModule": "{{fetchButtonId2}}"
                    }
                }
            },
            {
                "type": "invoke",
                "title": "{{fetchButtonTitle3}}",
                "value": {
                    "msteams": {
                        "type": "task/fetch"
                    },
                    "data": {
                        "taskModule": "{{fetchButtonId3}}"
                    }
                }
            },
            {
                "type": "invoke",
                "title": "{{fetchButtonTitle4}}",
                "value": {
                    "msteams": {
                        "type": "task/fetch"
                    },
                    "data": {
                        "taskModule": "{{fetchButtonId4}}"
                    }
                }
            },
            {
                "type": "invoke",
                "title": "{{fetchButtonTitle5}}",
                "value": {
                    "msteams": {
                        "type": "task/fetch"
                    },
                    "data": {
                        "taskModule": "{{fetchButtonId5}}"
                    }
                }
            }
          ],
        }
    },
    ninjaCatAdaptiveCard: {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "TextBlock",
                "text": "Here is a ninja cat:"
            },
            {
                "type": "Image",
                "url": "http://adaptivecards.io/content/cats/1.png",
                "size": "Medium"
            }
        ],
        "version": "1.0"
    },
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
    },
    questionFormOptions: {
        "type": "AdaptiveCard",
        "body": [
            {
                "type": "TextBlock",
                "size": "Medium",
                "weight": "Bolder",
                "text": "Choose a Form Type"
            },
            {
                "type": "TextBlock",
                "text": "Select which form you would like to use to ask your question:",
                "wrap": true
            }
        ],
        "actions": [
            {
                "type": "Action.OpenUrl",
                "title": "Rich Text Editor (TipTap)",
                "url": `${appRoot()}/customformtiptap`
            },
            {
                "type": "Action.OpenUrl",
                "title": "Simple Form",
                "url": `${appRoot()}/customform`
            }
        ],
        "version": "1.0"
    },
};

export const fetchTemplates: any = {
    questionformoptions: {
        title: "Choose Question Form Type",
        card: cardTemplates.questionFormOptions,
        height: 300,
        width: 500
    },
    customform: {
        title: constants.TaskModuleStrings.CustomFormTitle,
        height: 580,
        width: 920,
        url: `${appRoot()}/customform`,
        fallbackUrl: `${appRoot()}/customform`,
        completionBotId: process.env.MICROSOFT_APP_ID
    },
    customformtiptap: {
        title: "Ask a Question",
        height: 1020,
        width: 1632,
        url: `${appRoot()}/customform-tiptap`,
        fallbackUrl: `${appRoot()}/customform-tiptap`,
        completionBotId: process.env.MICROSOFT_APP_ID
    },
    youtube: {
        title: constants.TaskModuleStrings.YouTubeTitle,
        height: constants.TaskModuleSizes.youtube.height,
        width: constants.TaskModuleSizes.youtube.width,
        url: `${appRoot()}/youtube`,
        fallbackUrl: `${appRoot()}/youtube`,
        completionBotId: process.env.MICROSOFT_APP_ID
    },
    powerapp: {
        title: constants.TaskModuleStrings.PowerAppTitle,
        height: constants.TaskModuleSizes.powerapp.height,
        width: constants.TaskModuleSizes.powerapp.width,
        url: `${appRoot()}/powerapp`,
        fallbackUrl: `${appRoot()}/powerapp`,
        completionBotId: process.env.MICROSOFT_APP_ID
    },
    submitResponse: {
        title: constants.TaskModuleStrings.ActionSubmitResponseTitle,
        card: null,
        height: constants.TaskModuleSizes.adaptivecard.height,
        width: constants.TaskModuleSizes.adaptivecard.width,
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
