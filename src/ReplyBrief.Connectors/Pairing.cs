using System.Security.Cryptography;
using System.Text;

namespace ReplyBrief.Connectors;

/// <summary>
/// Issues and verifies the single-use secret that lets an Outlook task pane prove it may post a
/// capture to <see cref="LocalCaptureHost"/>. The code is 32 bytes of CSPRNG entropy in unpadded
/// base64url (43 characters), which carries far more entropy than the loopback attacker model needs.
/// </summary>
internal static class Pairing
{
    private const int EntropyBytes = 32;

    internal static string Create()
    {
        Span<byte> entropy = stackalloc byte[EntropyBytes];
        RandomNumberGenerator.Fill(entropy);
        return Convert.ToBase64String(entropy).TrimEnd('=').Replace('+', '-').Replace('/', '_');
    }

    /// <summary>
    /// Compares a presented code against the issued one without leaking where they first differ.
    /// The length guard runs first because <see cref="CryptographicOperations.FixedTimeEquals"/>
    /// throws on mismatched lengths; the code length is not a secret.
    /// </summary>
    internal static bool Matches(string issued, string? presented)
    {
        if (presented is null) return false;

        var issuedBytes = Encoding.UTF8.GetBytes(issued);
        var presentedBytes = Encoding.UTF8.GetBytes(presented);

        if (presentedBytes.Length != issuedBytes.Length) return false;
        return CryptographicOperations.FixedTimeEquals(issuedBytes, presentedBytes);
    }
}