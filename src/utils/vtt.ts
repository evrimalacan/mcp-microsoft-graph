/**
 * Parse WebVTT transcript to clean text format.
 * Strips timestamps, WEBVTT header, and XML-like voice tags.
 * Consolidates consecutive lines from the same speaker.
 */
export function parseVttToText(vttContent: string): string {
  const lines = vttContent.split('\n');
  const entries: Array<{ speaker: string; text: string }> = [];

  let currentSpeaker = '';
  let currentText = '';

  for (const line of lines) {
    const trimmed = line.trim();

    // Skip empty lines, WEBVTT header, and timestamp lines
    if (!trimmed || trimmed === 'WEBVTT' || /^\d{2}:\d{2}/.test(trimmed)) {
      continue;
    }

    // Parse voice tag: <v Speaker Name>Text</v>
    const voiceMatch = trimmed.match(/^<v ([^>]+)>(.+)<\/v>$/);
    if (voiceMatch) {
      const speaker = voiceMatch[1];
      const text = voiceMatch[2];

      // If same speaker, append to current text
      if (speaker === currentSpeaker) {
        currentText += ` ${text}`;
      } else {
        // Save previous entry if exists
        if (currentSpeaker && currentText) {
          entries.push({ speaker: currentSpeaker, text: currentText.trim() });
        }
        currentSpeaker = speaker;
        currentText = text;
      }
    }
  }

  // Don't forget the last entry
  if (currentSpeaker && currentText) {
    entries.push({ speaker: currentSpeaker, text: currentText.trim() });
  }

  // Format as "Speaker: Text" lines
  return entries.map((e) => `${e.speaker}: ${e.text}`).join('\n\n');
}
