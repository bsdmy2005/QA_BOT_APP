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
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
//

import * as microsoftTeams from "@microsoft/teams-js";
import * as constants from "./constants";
import { appRoot } from "./dialogs/CardTemplates";
import { taskModuleLink } from "./utils/DeepLinks";

declare var appId: any; // Injected at template render time

// Set the desired theme
function setTheme(theme: string): void {
    if (theme) {
        // Possible values for theme: 'default', 'light', 'dark' and 'contrast'
        document.body.className = "theme-" + (theme === "default" ? "light" : theme);
    }
}

// Call the initialize API first
microsoftTeams.initialize();

// Check the initial theme user chose and respect it
microsoftTeams.getContext(function(context: microsoftTeams.Context): void {
    if (context && context.theme) {
        setTheme(context.theme);
    }
});

// Handle theme changes
microsoftTeams.registerOnThemeChangeHandler(function(theme: string): void {
    setTheme(theme);
});

// Logic to initialize task module functionality
document.addEventListener("DOMContentLoaded", function(): void {
    // If we are on the Task Module page, initialize the buttons and deep links
    let taskModuleButtons = document.getElementsByClassName("taskModuleButton");
    if (taskModuleButtons.length > 0) {
        // Initialize deep link for TipTap form
        let deepLink = document.getElementById("dlCustomFormTipTap") as HTMLAnchorElement;
        deepLink.href = taskModuleLink(appId, constants.TaskModuleStrings.CustomFormTipTapTitle, 1200, 1920, `${appRoot()}/${constants.TaskModuleIds.CustomFormTipTap}`, null, `${appRoot()}/${constants.TaskModuleIds.CustomFormTipTap}`);

        // Initialize task module functionality
        let taskInfo = {
            title: null,
            height: null,
            width: null,
            url: null,
            fallbackUrl: null,
        };

        for (let btn of taskModuleButtons) {
            btn.addEventListener("click", function (): void {
                if (this.id.toLowerCase() === constants.TaskModuleIds.CustomFormTipTap) {
                    taskInfo.title = constants.TaskModuleStrings.CustomFormTipTapTitle;
                    taskInfo.height = 1020;
                    taskInfo.width = 1632;
                    taskInfo.url = `${appRoot()}/${this.id.toLowerCase()}?theme={theme}`;
                    
                    const submitHandler = (err: string, result: any): void => {
                        // Handle the result if needed
                        if (err) {
                            console.error("Error:", err);
                        }
                        if (result) {
                            console.log("Result:", result);
                        }
                    };
                    
                    microsoftTeams.tasks.startTask(taskInfo, submitHandler);
                }
            });
        }
    }
});
