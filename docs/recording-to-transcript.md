# Getting Transcripts from Recording URLs

When a user shares a Teams meeting recording URL, we can extract the transcript by leveraging metadata embedded in the recording's driveItem.

## How It Works

1. **Get driveItem metadata** from the recording URL using `/shares/{encoded}/driveItem`

2. **Extract meeting identifiers** from the `source` property:
   - `source.threadId` → the meeting chat ID (e.g., `19:meeting_...@thread.v2`)
   - `source.externalId` → the call ID that matches the transcript

3. **List transcripts** using `list_meeting_transcripts(chatId)` which returns all transcripts for that recurring meeting

4. **Match the correct transcript** by comparing `source.externalId` from the recording with `callId` from each transcript

5. **Fetch transcript content** using `get_meeting_transcript(meetingId, transcriptId)`

## Example

Recording URL metadata:
```json
{
  "source": {
    "threadId": "19:meeting_NTdjMDA5ZTMt...@thread.v2",
    "externalId": "cdbf5f91-6af3-4b61-97d8-4c2131c821fc"
  }
}
```

Matching transcript:
```json
{
  "id": "ktVizInG...",
  "callId": "cdbf5f91-6af3-4b61-97d8-4c2131c821fc"
}
```

The `externalId` and `callId` match, confirming this is the transcript for that specific recording.
