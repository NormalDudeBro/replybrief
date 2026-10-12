namespace ReplyBrief.Core;

public sealed record ConversationCapture(
    string Source,
    string Title,
    string Participants,
    string Content,
    string ContextUsed)
{
    public const int MaxContentLength = 60_000;

    public void Validate()
    {
        if (Source is not ("outlook" or "gmail" or "teams" or "manual"))
            throw new ArgumentException("Unknown conversation source.");
        if (string.IsNullOrWhiteSpace(Content))
            throw new ArgumentException("The conversation is empty.");
        if (Content.Length > MaxContentLength)
            throw new ArgumentException($"The conversation exceeds {MaxContentLength:N0} characters. Select a smaller portion.");
        if (Title is null || Title.Length > 500 || Participants is null || Participants.Length > 2_000 ||
            ContextUsed is null || ContextUsed.Length > 1_000)
            throw new ArgumentException("Conversation metadata is missing or too long.");
    }
}
