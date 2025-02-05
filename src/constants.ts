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

// Activity types
export const messageType = "message";
export const invokeType = "invoke";

// Dialog ids
// tslint:disable-next-line:variable-name
export const DialogId = {
    Root: "/",
    ACTester: "actester",
    BFCard: "bfcard",
};

// Telemetry events
// tslint:disable-next-line:variable-name
export const TelemetryEvent = {
    UserActivity: "UserActivity",
    BotActivity: "BotActivity",
};

// URL Placeholders - not currently supported
// tslint:disable-next-line:variable-name
export const UrlPlaceholders = "loginHint={loginHint}&upn={userPrincipalName}&aadId={userObjectId}&theme={theme}&groupId={groupId}&tenantId={tid}&locale={locale}";

// Task Module Strings
// tslint:disable-next-line:variable-name
export const TaskModuleStrings = {
    CustomFormTipTapTitle: "Ask a Question (TipTap Editor)",
    CustomFormTipTapName: "Ask Question (TipTap)",
};

// Task Module Ids
// tslint:disable-next-line:variable-name
export const TaskModuleIds = {
    CustomFormTipTap: "customformtiptap",
};

// Task Module Sizes
// tslint:disable-next-line:variable-name
export const TaskModuleSizes = {
    customformtiptap: {
        width: 1632,
        height: 1020
    }
};
