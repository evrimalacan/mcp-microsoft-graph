import type { Client } from '@microsoft/microsoft-graph-client';
import { markdownToHtml } from '../utils/markdown.js';
import { escapeODataString } from '../utils/odata.js';
import { processMentionsInHtml } from '../utils/users.js';
import type {
  CallTranscript,
  Chat,
  ChatMessage,
  CreateCalendarEventParams,
  CreateChatParams,
  Event,
  GetCalendarEventsParams,
  GetChatMessagesParams,
  GetChatParams,
  GetMeetingByJoinUrlParams,
  GetTranscriptContentParams,
  GetUserParams,
  GraphApiResponse,
  ListMailsParams,
  ListTranscriptsParams,
  Message,
  OnlineMeeting,
  SearchChatsParams,
  SearchChatsResult,
  SearchMessagesParams,
  SearchRequest,
  SearchResponse,
  SearchUsersParams,
  SendChatMessageParams,
  SendMailParams,
  SetMessageReactionParams,
  UnsetMessageReactionParams,
  User,
} from './graph.types.js';

export class GraphClient {
  constructor(private client: Client) {}

  // ===== User Operations =====

  async getCurrentUser(): Promise<User> {
    return await this.client.api('/me').get();
  }

  async getUser(params: GetUserParams): Promise<User> {
    return await this.client.api(`/users/${params.userId}`).get();
  }

  async searchUsers(params: SearchUsersParams): Promise<User[]> {
    const response = (await this.client
      .api('/users?ConsistencyLevel=eventual')
      .search(`"mail:${params.query}" OR "displayName:${params.query}"`)
      .get()) as GraphApiResponse<User>;

    return response.value ?? [];
  }

  // ===== Chat Operations =====

  async getChat(params: GetChatParams): Promise<Chat> {
    return await this.client.api(`/me/chats/${params.chatId}`).expand('members').get();
  }

  async searchChats(params: SearchChatsParams): Promise<SearchChatsResult> {
    let query = this.client.api('/me/chats').expand('members');

    // Apply limit (default 20 if not specified)
    const limit = params.limit || 20;
    query = query.top(limit);

    // Apply skipToken if provided (for pagination)
    if (params.skipToken) {
      query = query.query({ $skiptoken: params.skipToken });
    }

    const filters: string[] = [];

    if (params.searchTerm) {
      // Escape special characters (e.g., apostrophes) to prevent OData injection
      const escapedTerm = escapeODataString(params.searchTerm);
      filters.push(`contains(topic, '${escapedTerm}')`);
    }

    if (params.memberName) {
      // Escape special characters (e.g., apostrophes) to prevent OData injection
      const escapedName = escapeODataString(params.memberName);
      filters.push(`members/any(c:contains(c/displayName, '${escapedName}'))`);
    }

    if (params.chatTypes && params.chatTypes.length > 0) {
      // Build filter for chat types using 'or' operator
      const chatTypeFilters = params.chatTypes.map((type) => `chatType eq '${type}'`).join(' or ');
      filters.push(`(${chatTypeFilters})`);
    }

    if (filters.length > 0) {
      query = query.filter(filters.join(' and '));
    }

    const response = (await query.get()) as GraphApiResponse<Chat>;

    // Extract nextToken from @odata.nextLink if present
    let nextToken: string | undefined;
    if (response['@odata.nextLink']) {
      const url = new URL(response['@odata.nextLink']);
      const token = url.searchParams.get('$skiptoken');
      if (token) {
        nextToken = token;
      }
    }

    return {
      chats: response.value || [],
      nextToken,
    };
  }

  async createChat(params: CreateChatParams): Promise<Chat> {
    const payload = {
      chatType: params.chatType,
      members: params.members,
      ...(params.topic && { topic: params.topic }),
    };

    return await this.client.api('/chats').post(payload);
  }

  // ===== Message Operations =====

  async getChatMessages(params: GetChatMessagesParams): Promise<ChatMessage[]> {
    let query = this.client
      .api(`/me/chats/${params.chatId}/messages`)
      .top(params.limit || 20)
      .orderby('createdDateTime desc');

    if (params.fromUser) {
      query = query.filter(`from/user/id eq '${params.fromUser}'`);
    }

    const response = (await query.get()) as GraphApiResponse<ChatMessage>;
    let messages = response?.value || [];

    // Client-side filtering for date ranges
    if ((params.from || params.to) && messages.length > 0) {
      messages = messages.filter((message: ChatMessage) => {
        if (!message.createdDateTime) return true;

        const messageDate = new Date(message.createdDateTime);
        if (params.from) {
          const fromDate = new Date(params.from);
          if (messageDate <= fromDate) return false;
        }
        if (params.to) {
          const toDate = new Date(params.to);
          if (messageDate >= toDate) return false;
        }
        return true;
      });
    }

    return messages;
  }

  async sendChatMessage(params: SendChatMessageParams): Promise<ChatMessage> {
    let content: string;
    let contentType: 'text' | 'html';

    // Handle markdown conversion
    if (params.format === 'markdown') {
      content = await markdownToHtml(params.message);
      contentType = 'html';
    } else {
      content = params.message;
      contentType = 'text';
    }

    // Process mentions if provided
    const mentionMappings: Array<{ mention: string; userId: string; displayName: string }> = [];
    if (params.mentions && params.mentions.length > 0) {
      for (const mention of params.mentions) {
        try {
          const userResponse = await this.client.api(`/users/${mention.userId}`).select('displayName').get();
          mentionMappings.push({
            mention: mention.mention,
            userId: mention.userId,
            displayName: userResponse.displayName || mention.mention,
          });
        } catch (_error) {
          mentionMappings.push({
            mention: mention.mention,
            userId: mention.userId,
            displayName: mention.mention,
          });
        }
      }
    }

    let finalMentions: Array<{
      id: number;
      mentionText: string;
      mentioned: { user: { id: string } };
    }> = [];

    if (mentionMappings.length > 0) {
      const result = processMentionsInHtml(content, mentionMappings);
      content = result.content;
      finalMentions = result.mentions;
      contentType = 'html';
    }

    const messagePayload: any = {
      body: {
        content,
        contentType,
      },
      importance: params.importance || 'normal',
    };

    if (finalMentions.length > 0) {
      messagePayload.mentions = finalMentions;
    }

    return await this.client.api(`/me/chats/${params.chatId}/messages`).post(messagePayload);
  }

  async searchMessages(params: SearchMessagesParams): Promise<SearchResponse> {
    const queryParts: string[] = [];

    if (params.query) {
      queryParts.push(params.query);
    }

    if (params.mentions) {
      queryParts.push(`mentions:${params.mentions}`);
    }

    if (params.from) {
      const dateOnly = params.from.split('T')[0];
      queryParts.push(`sent>=${dateOnly}`);
    }

    if (params.fromUser) {
      queryParts.push(`from:${params.fromUser}`);
    }

    let enhancedQuery = queryParts.length > 0 ? queryParts.join(' AND ') : '*';

    const scope = params.scope || 'all';
    if (scope === 'channels') {
      enhancedQuery = `${enhancedQuery} AND (channelIdentity/channelId:*)`;
    } else if (scope === 'chats') {
      enhancedQuery = `${enhancedQuery} AND (chatId:* AND NOT channelIdentity/channelId:*)`;
    }

    const searchRequest: SearchRequest = {
      entityTypes: ['chatMessage'],
      query: {
        queryString: enhancedQuery,
      },
      from: 0,
      size: params.limit || 25,
      enableTopResults: params.enableTopResults !== undefined ? params.enableTopResults : true,
    };

    return await this.client.api('/search/query').post({ requests: [searchRequest] });
  }

  async setMessageReaction(params: SetMessageReactionParams): Promise<void> {
    await this.client
      .api(`/chats/${params.chatId}/messages/${params.messageId}/setReaction`)
      .post({ reactionType: params.reactionType });
  }

  async unsetMessageReaction(params: UnsetMessageReactionParams): Promise<void> {
    await this.client
      .api(`/chats/${params.chatId}/messages/${params.messageId}/unsetReaction`)
      .post({ reactionType: params.reactionType });
  }

  // ===== Mail Operations =====

  async listMails(params: ListMailsParams): Promise<GraphApiResponse<Message>> {
    const limit = params.limit || 20;
    const includeBody = params.includeBody || false;
    const bodyFormat = params.bodyFormat || 'html';

    // Build select fields based on whether full body is requested
    const selectFields = includeBody
      ? 'id,subject,from,receivedDateTime,body,bodyPreview,isRead,hasAttachments'
      : 'id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments';

    let query = this.client.api('/me/messages').top(limit).orderby('receivedDateTime desc').select(selectFields);

    // Set body content format preference (only applies when body is included in select)
    if (includeBody && bodyFormat === 'text') {
      query = query.header('Prefer', 'outlook.body-content-type="text"');
    }

    // Build filter conditions - receivedDateTime must be first since it's in orderby
    const filters: string[] = [];

    if (params.since) {
      filters.push(`receivedDateTime ge ${params.since}`);
    } else {
      // Use a broad date range if not specified to satisfy API requirement
      filters.push(`receivedDateTime ge 1900-01-01T00:00:00Z`);
    }

    if (params.until) {
      filters.push(`receivedDateTime le ${params.until}`);
    }

    // Other filters must come after receivedDateTime (the orderby field)
    if (params.from) {
      filters.push(`from/emailAddress/address eq '${params.from}'`);
    }

    if (params.unreadOnly) {
      filters.push('isRead eq false');
    }

    if (params.hasAttachments !== undefined) {
      filters.push(`hasAttachments eq ${params.hasAttachments}`);
    }

    query = query.filter(filters.join(' and '));

    return await query.get();
  }

  async sendMail(params: SendMailParams): Promise<void> {
    const message: any = {
      subject: params.subject,
      body: {
        contentType: params.bodyType || 'text',
        content: params.body,
      },
      toRecipients: params.to.map((email) => ({
        emailAddress: { address: email },
      })),
    };

    if (params.cc && params.cc.length > 0) {
      message.ccRecipients = params.cc.map((email) => ({
        emailAddress: { address: email },
      }));
    }

    if (params.bcc && params.bcc.length > 0) {
      message.bccRecipients = params.bcc.map((email) => ({
        emailAddress: { address: email },
      }));
    }

    await this.client.api('/me/sendMail').post({ message });
  }

  // ===== Calendar Operations =====

  async getCalendarEvents(params: GetCalendarEventsParams): Promise<Event[]> {
    const limit = params.limit || 50;
    const now = new Date().toISOString();
    const thirtyDaysFromNow = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();

    const startDateTime = params.from || now;
    const endDateTime = params.to || thirtyDaysFromNow;

    let query = this.client
      .api('/me/calendarView')
      .query({
        startDateTime,
        endDateTime,
      })
      .top(limit)
      .orderby('start/dateTime');

    if (params.subject) {
      query = query.filter(`contains(subject, '${params.subject}')`);
    }

    const response = (await query.get()) as GraphApiResponse<Event>;
    return response.value || [];
  }

  async createCalendarEvent(params: CreateCalendarEventParams): Promise<Event> {
    const event: any = {
      subject: params.subject,
      start: {
        dateTime: params.startDateTime,
        timeZone: 'UTC',
      },
      end: {
        dateTime: params.endDateTime,
        timeZone: 'UTC',
      },
      body: {
        contentType: 'text',
        content: params.body,
      },
    };

    if (params.attendees && params.attendees.length > 0) {
      event.attendees = params.attendees.map((attendee) => {
        if (typeof attendee === 'string') {
          return {
            emailAddress: { address: attendee },
            type: 'required',
          };
        } else {
          return {
            emailAddress: {
              address: attendee.email,
              ...(attendee.name && { name: attendee.name }),
            },
            type: attendee.type || 'required',
          };
        }
      });
    }

    if (params.location) {
      event.location = {
        displayName: params.location,
      };
    }

    if (params.isOnlineMeeting) {
      event.isOnlineMeeting = true;
      event.onlineMeetingProvider = 'teamsForBusiness';
    }

    return await this.client.api('/me/events').post(event);
  }

  // ===== Transcript Operations =====

  async getMeetingByJoinUrl(params: GetMeetingByJoinUrlParams): Promise<OnlineMeeting | null> {
    const response = (await this.client
      .api('/me/onlineMeetings')
      .filter(`JoinWebUrl eq '${params.joinWebUrl}'`)
      .get()) as GraphApiResponse<OnlineMeeting>;

    return response.value?.[0] || null;
  }

  async listTranscripts(params: ListTranscriptsParams): Promise<CallTranscript[]> {
    const response = (await this.client
      .api(`/me/onlineMeetings/${params.meetingId}/transcripts`)
      .get()) as GraphApiResponse<CallTranscript>;

    return response.value || [];
  }

  async getTranscriptContent(params: GetTranscriptContentParams): Promise<string> {
    const stream = await this.client
      .api(`/me/onlineMeetings/${params.meetingId}/transcripts/${params.transcriptId}/content`)
      .query({ $format: 'text/vtt' })
      .get();

    // Convert stream to text
    const reader = stream.getReader();
    const chunks: Uint8Array[] = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
    }

    const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const combined = new Uint8Array(totalLength);
    let offset = 0;
    for (const chunk of chunks) {
      combined.set(chunk, offset);
      offset += chunk.length;
    }

    return new TextDecoder('utf-8').decode(combined);
  }
}
