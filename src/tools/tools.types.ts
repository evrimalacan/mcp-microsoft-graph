import type { ChatType } from '../client/graph.types.js';

// Define Reaction type locally since it's not exported from @microsoft/microsoft-graph-types
export interface Reaction {
  reactionType?: string;
  createdDateTime?: string;
  user?: {
    user?: {
      id?: string;
      displayName?: string;
    };
  };
}

/**
 * Optimized types for MCP tools.
 * These types strip unnecessary fields from API responses to reduce token usage.
 */

// ===== User Types =====

export interface OptimizedUser {
  id?: string;
  displayName?: string | null;
  mail?: string | null;
  jobTitle?: string | null;
  department?: string | null;
}

// ===== Chat Types =====

export interface OptimizedChatMember {
  displayName?: string | null;
  email?: string | null;
}

export interface OptimizedChat {
  id?: string;
  topic?: string | null;
  chatType?: ChatType;
  members?: OptimizedChatMember[];
}

// ===== Message Types =====

export interface OptimizedAttachment {
  name?: string;
  contentUrl?: string; // For download_sharepoint_file
}

export interface OptimizedChatMessage {
  id?: string;
  content?: string;
  from?: string; // Just display name, not full User object
  createdDateTime?: string;
  reactions?: any[]; // Use any[] to avoid type conflicts with ChatMessageReaction
  attachments?: OptimizedAttachment[]; // SharePoint file attachments
}

export interface OptimizedChatMessageWithFilters {
  filters: {
    from?: string;
    to?: string;
    fromUser?: string;
  };
  filteringMethod: 'client-side' | 'server-side';
  totalReturned: number;
  hasMore: boolean;
  messages: OptimizedChatMessage[];
}

export interface OptimizedSearchHit {
  id: string;
  rank: number;
  content: string;
  from?: string;
  createdDateTime?: string;
  chatId?: string;
  teamId?: string;
  channelId?: string;
}

export interface OptimizedSearchResult {
  query?: string;
  scope: string;
  totalResults: number;
  results: OptimizedSearchHit[];
  moreResultsAvailable: boolean;
}

// ===== Mail Types =====

export interface OptimizedMail {
  id?: string;
  subject?: string;
  from?: {
    name?: string;
    address?: string;
  };
  receivedDateTime?: string;
  bodyPreview?: string;
  body?: {
    contentType: string;
    content: string;
  };
  isRead?: boolean;
  hasAttachments?: boolean;
}

// ===== Calendar Types =====

export interface OptimizedLocation {
  name?: string;
  type?: string; // "conferenceRoom", "default", etc.
}

export interface OptimizedAttendee {
  name?: string;
  email?: string;
  status?: string; // "accepted", "declined", "tentative", "none"
}

export interface OptimizedEvent {
  id?: string;
  subject?: string;
  start?: any; // Keep as any to avoid complex NullableOption type issues
  end?: any;
  location?: OptimizedLocation;
  attendees?: OptimizedAttendee[];
  isOnlineMeeting?: boolean;
  chatId?: string; // Meeting chat ID for online meetings (extracted from joinUrl)
}

export interface OptimizedMeetingTimeSuggestion {
  confidence: number;
  start: string;
  end: string;
}

// ===== Transcript Types =====

// SDK CallTranscript type is incomplete - API returns callId and endDateTime
export interface TranscriptWithExtras {
  id?: string;
  callId?: string;
  createdDateTime?: string;
  endDateTime?: string;
}
