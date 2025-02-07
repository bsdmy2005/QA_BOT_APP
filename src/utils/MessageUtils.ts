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

import urlJoin from "url-join";
import {
    Activity,
    TurnContext,
    TeamsInfo,
    CardFactory,
    MessageFactory,
    ConversationReference,
    ChannelInfo,
    TeamDetails,
    TeamsChannelData,
    Entity,
    ConversationAccount
} from 'botbuilder';
import axios, { AxiosRequestConfig } from 'axios';
import { Logger } from 'winston';

// Creates a new Message
export function createMessage(context: TurnContext, text = "", textFormat = "xml"): Partial<Activity> {
    return MessageFactory.text(text, text, textFormat);
}

// Get the channel id in the event
export function getChannelId(activity: Partial<Activity>): string {
    const channelData = activity.channelData as TeamsChannelData;
    if (channelData && channelData.channel) {
        return channelData.channel.id;
    }
    return "";
}

// Get the team id in the event
export function getTeamId(activity: Partial<Activity>): string {
    const channelData = activity.channelData as TeamsChannelData;
    if (channelData && channelData.team) {
        return channelData.team.id;
    }
    return "";
}

// Get the tenant id in the event
export function getTenantId(activity: Partial<Activity>): string {
    const channelData = activity.channelData as TeamsChannelData;
    if (channelData && channelData.tenant) {
        return channelData.tenant.id;
    }
    return "";
}

// Returns true if this is message sent to a channel
export function isChannelMessage(activity: Partial<Activity>): boolean {
    return !!getChannelId(activity);
}

// Returns true if this is message sent to a group (group chat or channel)
export function isGroupMessage(activity: Partial<Activity>): boolean {
    const conversation = activity.conversation;
    return conversation?.conversationType === 'channel' || isChannelMessage(activity);
}

// Strip all mentions from text
export function getTextWithoutMentions(activity: Partial<Activity>): string {
    let text = activity.text || '';
    if (activity.entities) {
        activity.entities
            .filter(entity => entity.type === "mention")
            .forEach(entity => {
                text = text.replace(entity.text, "");
            });
        text = text.trim();
    }
    return text;
}

// Get all user mentions
export function getUserMentions(activity: Partial<Activity>): Entity[] {
    const entities = activity.entities || [];
    const botId = activity.recipient?.id?.toLowerCase();
    return entities.filter(entity => 
        entity.type === "mention" && 
        entity.mentioned?.id?.toLowerCase() !== botId
    );
}

// Create a mention entity for the user that sent this message
export function createUserMention(activity: Partial<Activity>): Entity {
    const user = activity.from;
    const text = `<at>${user?.name}</at>`;
    return {
        type: "mention",
        mentioned: {
            id: user?.id || '',
            name: user?.name || '',
            aadObjectId: (user as any)?.aadObjectId
        },
        text: text
    };
}

// Gets the members of the given conversation
export async function getConversationMembers(context: TurnContext, conversationId?: string): Promise<any[]> {
    try {
        const actualConversationId = conversationId || context.activity.conversation?.id;
        if (!actualConversationId) {
            throw new Error("No conversation ID provided");
        }
        return await TeamsInfo.getMembers(context);
    } catch (error) {
        throw new Error(`Failed to get conversation members: ${error.message}`);
    }
}

// Starts a 1:1 chat with the given user
export async function startPersonalConversation(context: TurnContext, userId: string): Promise<ConversationReference> {
    try {
        const ref = await TeamsInfo.getMember(context, userId);
        const conversationParameters = {
            isGroup: false,
            channelData: {
                tenant: getTenantId(context.activity)
            },
            bot: context.activity.recipient,
            members: [ref],
            tenantId: getTenantId(context.activity)
        };
        
        const connectorClient = context.turnState.get(context.adapter.ConnectorClientKey);
        const conversationResource = await connectorClient.conversations.createConversation(conversationParameters);
        
        return {
            channelId: context.activity.channelId,
            serviceUrl: context.activity.serviceUrl,
            conversation: {
                id: conversationResource.id,
                name: '',
                isGroup: false,
                conversationType: 'personal'
            } as ConversationAccount,
            bot: context.activity.recipient,
            user: ref
        };
    } catch (error) {
        throw new Error(`Failed to start personal conversation: ${error.message}`);
    }
}

// Get locale from client info in activity
export function getLocale(activity: Partial<Activity>): string {
    if (activity.entities) {
        const clientInfo = activity.entities.find(e => e.type === "clientInfo");
        return (clientInfo as any)?.locale;
    }
    return activity.locale || null;
}

// Get Teams context from activity
export function getTeamsContext(activity: Partial<Activity>): any {
    const channelData = activity.channelData as TeamsChannelData;
    const meetingInfo = channelData?.meeting || {};
    
    return {
        locale: getLocale(activity),
        country: null, // No longer available in v4
        platform: 'teams',
        timezone: null, // No longer available in v4
        tenant: getTenantId(activity),
        teamsChannelId: getChannelId(activity),
        teamsTeamId: getTeamId(activity),
        userObjectId: (activity.from as any)?.aadObjectId,
        messageId: activity.id,
        conversationId: activity.conversation?.id,
        conversationType: activity.conversation?.conversationType,
        theme: (meetingInfo as any)?.theme || null
    };
}
