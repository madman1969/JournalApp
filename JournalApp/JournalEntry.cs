using System;

namespace JournalApp
{
  public class JournalEntry
  {
    public string ConversationID;   // Unique ID for journal entry
    public string Subject;
    public string EntryType;
    public DateTime StartTime;
    public string Body;
  }
}
