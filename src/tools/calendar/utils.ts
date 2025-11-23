import type { Event } from '../../client/graph.types.js';
import { extractChatIdFromJoinUrl } from '../../utils/teams.js';
import type { OptimizedAttendee, OptimizedEvent, OptimizedLocation } from '../tools.types.js';

/**
 * Transforms a Graph API Event into an OptimizedEvent for MCP responses.
 * Extracts chatId from joinUrl, simplifies location and attendees.
 */
export function optimizeEvent(event: Event): OptimizedEvent {
  // Extract chatId from joinUrl for online meetings
  let chatId: string | undefined;
  if (event.isOnlineMeeting && event.onlineMeeting?.joinUrl) {
    try {
      chatId = extractChatIdFromJoinUrl(event.onlineMeeting.joinUrl);
    } catch {
      // If extraction fails, leave chatId undefined
    }
  }

  // Simplify location to just name and type
  let location: OptimizedLocation | undefined;
  if (event.location?.displayName) {
    location = {
      name: event.location.displayName,
      type: event.location.locationType || undefined,
    };
  }

  // Simplify attendees to just name, email, and status
  let attendees: OptimizedAttendee[] | undefined;
  if (event.attendees?.length) {
    attendees = event.attendees.map((a) => ({
      name: a.emailAddress?.name || undefined,
      email: a.emailAddress?.address || undefined,
      status: a.status?.response || undefined,
    }));
  }

  return {
    id: event.id,
    subject: event.subject || undefined,
    start: event.start || undefined,
    end: event.end || undefined,
    location,
    attendees,
    isOnlineMeeting: event.isOnlineMeeting || undefined,
    chatId,
  };
}
