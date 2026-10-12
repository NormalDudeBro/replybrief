namespace ReplyBrief.Connectors.Tests;

public sealed class PairingTests
{
    [Fact]
    public void IssuesAnUnpaddedBase64UrlCodeWithThirtyTwoBytesOfEntropy()
    {
        var code = Pairing.Create();

        Assert.Equal(43, code.Length);
        Assert.DoesNotContain("=", code);
        Assert.DoesNotContain("+", code);
        Assert.DoesNotContain("/", code);
        Assert.All(code, character => Assert.True(char.IsAsciiLetterOrDigit(character) || character is '-' or '_', $"unexpected character '{character}' in a pairing code"));
    }

    [Fact]
    public void IssuesADifferentCodeEveryTime()
    {
        var codes = Enumerable.Range(0, 512).Select(_ => Pairing.Create()).ToArray();

        Assert.Equal(codes.Length, codes.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void AcceptsOnlyTheCodeItIssued()
    {
        var issued = Pairing.Create();

        Assert.True(Pairing.Matches(issued, issued));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("nope")]
    [InlineData("AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA")]
    public void RejectsAnythingThatIsNotTheIssuedCode(string? presented)
    {
        Assert.False(Pairing.Matches(Pairing.Create(), presented));
    }

    [Fact]
    public void RejectsACodeThatDiffersOnlyInItsLastCharacter()
    {
        var issued = Pairing.Create();
        var presented = issued[..^1] + (issued[^1] == 'A' ? 'B' : 'A');

        Assert.NotEqual(issued, presented);
        Assert.False(Pairing.Matches(issued, presented));
    }

    [Fact]
    public void RejectsACodeOfTheRightLengthButTheWrongBytes()
    {
        var issued = Pairing.Create();

        // Same length, every character different: the case the fixed time comparison exists for.
        var presented = new string(issued.Select(character => character == 'a' ? 'b' : 'a').ToArray());

        Assert.Equal(issued.Length, presented.Length);
        Assert.False(Pairing.Matches(issued, presented));
    }
}