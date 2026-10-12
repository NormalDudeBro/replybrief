namespace ReplyBrief.Core;

public sealed record AnalysisResult(
    string[] Summary,
    string[] NeedsAttention,
    string[] SuggestedReplyPoints,
    string[] MissingInformation);
