import type {
  Attendee,
  Channel,
  ChannelMembershipType,
  Chat,
  ChatMessage,
  ChatMessageImportance,
  ChatMessageInfo,
  ChatType,
  ConversationMember,
  DateTimeTimeZone,
  EmailAddress,
  Event,
  ItemBody,
  Location,
  NullableOption,
  Team,
  TeamSpecialization,
  TeamsAppInstallation,
  TeamVisibilityType,
  User,
} from '@microsoft/microsoft-graph-types';

// Re-export Microsoft Graph types we use
export type {
  Attendee,
  Channel,
  ChannelMembershipType,
  Chat,
  ChatMessage,
  ChatMessageImportance,
  ChatMessageInfo,
  ChatType,
  ConversationMember,
  DateTimeTimeZone,
  EmailAddress,
  Event,
  ItemBody,
  Location,
  NullableOption,
  Team,
  TeamSpecialization,
  TeamsAppInstallation,
  TeamVisibilityType,
  User,
};

// Custom types for our responses
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

// Simplified types for our API responses - all properties are optional to handle Graph API variability
export interface UserSummary {
  id?: string | undefined;
  displayName?: NullableOption<string> | undefined;
  userPrincipalName?: NullableOption<string> | undefined;
  mail?: NullableOption<string> | undefined;
  jobTitle?: NullableOption<string> | undefined;
  department?: NullableOption<string> | undefined;
  officeLocation?: NullableOption<string> | undefined;
  photo?: NullableOption<string> | undefined;
}

export interface TeamSummary {
  id?: string | undefined;
  displayName?: NullableOption<string> | undefined;
  description?: NullableOption<string> | undefined;
  isArchived?: NullableOption<boolean> | undefined;
}

export interface ChannelSummary {
  id?: string | undefined;
  displayName?: string | undefined;
  description?: NullableOption<string> | undefined;
  membershipType?: NullableOption<ChannelMembershipType> | undefined;
}

export interface ChatSummary {
  id?: string | undefined;
  topic?: NullableOption<string> | undefined;
  chatType?: ChatType | undefined;
  memberCount?: number | undefined;
}

export interface MessageSummary {
  id?: string | undefined;
  content?: NullableOption<string> | undefined;
  from?: NullableOption<string> | undefined;
  createdDateTime?: NullableOption<string> | undefined;
  importance?: ChatMessageImportance | undefined;
}

// Create chat payload
export interface CreateChatPayload {
  chatType: 'oneOnOne' | 'group';
  members: ConversationMember[];
  topic?: string;
}

// Send message payload
export interface SendMessagePayload {
  body: {
    content: string;
    contentType: 'text' | 'html';
  };
  importance?: ChatMessageImportance;
}

// New types for search functionality
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
