// Re-export official Microsoft Graph types
export type {
  Attendee,
  CallTranscript,
  Channel,
  ChannelMembershipType,
  Chat,
  ChatMessage,
  ChatMessageAttachment,
  ChatMessageImportance,
  ChatMessageInfo,
  ChatType,
  ConversationMember,
  DateTimeTimeZone,
  EmailAddress,
  Event,
  ItemBody,
  Location,
  MeetingTimeSuggestion,
  MeetingTimeSuggestionsResult,
  Message,
  NullableOption,
  OnlineMeeting,
  Team,
  TeamSpecialization,
  TeamsAppInstallation,
  TeamVisibilityType,
  User,
} from '@microsoft/microsoft-graph-types';

// Also import for use in interface definitions
import type { Chat } from '@microsoft/microsoft-graph-types';

// ===== API Response Wrapper Types =====

export interface GraphApiResponse<T> {
  value?: T[];
  '@odata.count'?: number;
  '@odata.nextLink'?: string;
}

export interface GraphError {
  code: string;
  message: string;
  innerError?: {
    code?: string;
    message?: string;
    'request-id'?: string;
    date?: string;
  };
}

// ===== Search API Types =====

export interface SearchRequest {
  entityTypes: string[];
  query: {
    queryString: string;
  };
  from?: number;
  size?: number;
  enableTopResults?: boolean;
}

export interface SearchResponse {
  value: SearchResult[];
}

export interface SearchResult {
  searchTerms: string[];
  hitsContainers: SearchHitsContainer[];
}

export interface SearchHitsContainer {
  hits: SearchHit[];
  total: number;
  moreResultsAvailable: boolean;
}

export interface SearchHit {
  hitId: string;
  rank: number;
  summary: string;
  resource: {
    '@odata.type': string;
    id: string;
    createdDateTime?: string;
    lastModifiedDateTime?: string;
    from?: {
      emailAddress?: {
        name?: string;
        address?: string;
      };
    };
    subject?: string;
    chatId?: string;
    channelIdentity?: {
      teamId?: string;
      channelId?: string;
    };
  };
}

// ===== Method Parameter Types =====

// User operations
export interface GetUserParams {
  userId: string;
}

export interface SearchUsersParams {
  query: string;
  limit?: number;
}

// Chat operations
export interface GetChatParams {
  chatId: string;
}

export interface SearchChatsParams {
  searchTerm?: string;
  memberName?: string;
  chatTypes?: Array<'oneOnOne' | 'group' | 'meeting'>;
  limit?: number;
  skipToken?: string;
}

export interface SearchChatsResult {
  chats: Chat[];
  nextToken?: string;
}

export interface CreateChatParams {
  chatType: 'oneOnOne' | 'group';
  members: Array<any>; // Using any to avoid importing ConversationMember directly
  topic?: string;
}

// Message operations
export interface GetChatMessagesParams {
  chatId: string;
  limit?: number;
  from?: string; // ISO datetime
  to?: string; // ISO datetime
  fromUser?: string;
}

export interface SendChatMessageParams {
  chatId: string;
  message: string;
  importance?: 'normal' | 'high' | 'urgent';
  format?: 'text' | 'markdown';
  replyToId?: string;
  attachments?: string[];
}

export interface UpdateChatMessageParams {
  chatId: string;
  messageId: string;
  message: string;
  importance?: 'normal' | 'high' | 'urgent';
  format?: 'text' | 'markdown';
}

export interface SearchMessagesParams {
  query?: string;
  scope?: 'all' | 'channels' | 'chats';
  limit?: number;
  enableTopResults?: boolean;
  mentions?: string;
  from?: string;
  fromUser?: string;
}

export interface SetMessageReactionParams {
  chatId: string;
  messageId: string;
  reactionType: string;
}

export interface UnsetMessageReactionParams {
  chatId: string;
  messageId: string;
  reactionType: string;
}

// Mail operations
export interface ListMailsParams {
  limit?: number;
  since?: string;
  until?: string;
  unreadOnly?: boolean;
  hasAttachments?: boolean;
  from?: string;
  includeBody?: boolean;
  bodyFormat?: 'text' | 'html';
}

export interface SendMailParams {
  to: string[];
  subject: string;
  body: string;
  bodyType?: 'text' | 'html';
  cc?: string[];
  bcc?: string[];
}

// Presence operations
export interface SetPresenceParams {
  availability:
    | 'Available'
    | 'Away'
    | 'BeRightBack'
    | 'Busy'
    | 'DoNotDisturb'
    | 'InACall'
    | 'InAMeeting'
    | 'Offline'
    | 'PresenceUnknown'
    | 'Presenting';
  activity?:
    | 'Available'
    | 'Away'
    | 'BeRightBack'
    | 'Busy'
    | 'DoNotDisturb'
    | 'InACall'
    | 'InAMeeting'
    | 'Offline'
    | 'PresenceUnknown'
    | 'Presenting';
  expirationDuration?: string; // ISO 8601 duration (e.g., 'PT1H' for 1 hour)
}

export interface SetPreferredPresenceParams {
  availability: 'Available' | 'Away' | 'BeRightBack' | 'Busy' | 'DoNotDisturb' | 'Offline';
  activity?: 'Available' | 'Away' | 'BeRightBack' | 'Busy' | 'DoNotDisturb' | 'Offline';
  expirationDuration?: string; // ISO 8601 duration (e.g., 'PT8H' for 8 hours)
}

// Calendar operations
export interface GetCalendarEventsParams {
  from?: string;
  to?: string;
  limit?: number;
  subject?: string;
}

export interface CreateCalendarEventParams {
  subject: string;
  startDateTime: string;
  endDateTime: string;
  body: string;
  attendees?: Array<
    | string
    | {
        email: string;
        name?: string;
        type?: 'required' | 'optional' | 'resource';
      }
  >;
  location?: string;
  isOnlineMeeting?: boolean;
}

export interface CancelCalendarEventParams {
  eventId: string;
  comment?: string;
}

// Availability operations
export interface FindMeetingTimesParams {
  attendees: Array<{
    email: string;
    type: 'required' | 'optional';
  }>;
  startDateTime: string; // ISO 8601 UTC format
  endDateTime: string; // ISO 8601 UTC format
  meetingDuration: string; // ISO 8601 duration format (e.g., 'PT30M', 'PT1H')
}

// Transcript operations
export interface GetMeetingByJoinUrlParams {
  joinWebUrl: string;
}

export interface ListTranscriptsParams {
  meetingId: string;
}

export interface GetTranscriptContentParams {
  meetingId: string;
  transcriptId: string;
}

// ===== SharePoint/OneDrive Operations =====

export interface DownloadSharePointFileParams {
  contentUrl: string;
}

export interface DownloadSharePointFileResult {
  data: Buffer;
  filename: string;
  size: number;
  itemId: string;
  driveId: string;
}

export interface UploadToSharePointParams {
  filePath: string;
}

export interface UploadToSharePointResult {
  itemId: string;
  driveId: string;
  shareUrl: string;
}

export interface UpdateSharePointFileParams {
  itemId: string;
  driveId: string;
  filePath: string;
}

export interface UpdateSharePointFileResult {
  filename: string;
  size: number;
}

export interface GrantFilePermissionParams {
  itemId: string;
  driveId: string;
  emails: string[];
  role: 'read' | 'write';
}
