/**
 * Utilities for working with Microsoft Teams URLs and IDs
 */

export interface MeetingInfo {
  meetingId: string;
  chatId: string;
}

/**
 * Extracts meetingId and chatId from a Teams meeting join URL.
 *
 * The meetingId is constructed by base64-encoding: "1*{organizerId}*0**{threadId}"
 * The chatId is the threadId extracted from the URL path.
 *
 * @param joinUrl - The Teams meeting join URL (from calendar event or chat)
 * @returns Object containing meetingId and chatId
 * @throws Error if URL cannot be parsed
 *
 * @example
 * const url = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc...%40thread.v2/0?context=...";
 * const { meetingId, chatId } = extractMeetingInfoFromJoinUrl(url);
 * // chatId = "19:meeting_abc...@thread.v2"
 * // meetingId = "MSo1NmNh..." (base64 encoded)
 */
export function extractMeetingInfoFromJoinUrl(joinUrl: string): MeetingInfo {
  // Decode the URL to handle %3a → : and %40 → @
  const decodedUrl = decodeURIComponent(joinUrl);

  // Extract threadId (chatId) from URL path
  // Pattern: /meetup-join/19:meeting_...@thread.v2/
  const threadIdMatch = decodedUrl.match(/meetup-join\/(19:[^/]+)/);
  if (!threadIdMatch) {
    throw new Error(`Could not extract threadId from join URL: ${joinUrl.substring(0, 100)}...`);
  }
  const chatId = threadIdMatch[1];

  // Extract Oid (organizer ID) from context JSON parameter
  // Pattern: context={"Tid":"...","Oid":"..."}
  const contextMatch = decodedUrl.match(/context=(\{[^}]+\})/);
  if (!contextMatch) {
    throw new Error(`Could not extract context from join URL: ${joinUrl.substring(0, 100)}...`);
  }

  let context: { Tid?: string; Oid?: string };
  try {
    context = JSON.parse(contextMatch[1]);
  } catch {
    throw new Error(`Could not parse context JSON from join URL`);
  }

  if (!context.Oid) {
    throw new Error(`Context missing Oid (organizer ID) in join URL`);
  }

  // Construct meetingId: base64("1*{Oid}*0**{threadId}")
  const meetingIdRaw = `1*${context.Oid}*0**${chatId}`;
  const meetingId = Buffer.from(meetingIdRaw).toString('base64');

  return { meetingId, chatId };
}

/**
 * Extracts just the chatId (threadId) from a Teams meeting join URL.
 * Use this when you only need the chatId and don't need to compute meetingId.
 *
 * @param joinUrl - The Teams meeting join URL
 * @returns The chatId (e.g., "19:meeting_...@thread.v2")
 * @throws Error if URL cannot be parsed
 */
export function extractChatIdFromJoinUrl(joinUrl: string): string {
  const decodedUrl = decodeURIComponent(joinUrl);
  const threadIdMatch = decodedUrl.match(/meetup-join\/(19:[^/]+)/);
  if (!threadIdMatch) {
    throw new Error(`Could not extract threadId from join URL: ${joinUrl.substring(0, 100)}...`);
  }
  return threadIdMatch[1];
}
