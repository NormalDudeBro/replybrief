using System.Xml.Linq;

namespace ReplyBrief.Connectors.Tests;

/// <summary>
/// Guards the Outlook manifest against schema mistakes that an XML parser cannot see: wrong element
/// names, invented resource elements, and resid references that resolve to nothing. Every name and
/// relationship checked here is taken from Microsoft's own mail manifest sample, not from this
/// repository's preferences.
/// </summary>
public sealed class OutlookManifestTests
{
    private const string AppForOffice = "http://schemas.microsoft.com/office/appforoffice/1.1";
    private const string BasicTypes = "http://schemas.microsoft.com/office/officeappbasictypes/1.0";
    private const string MailOverrides = "http://schemas.microsoft.com/office/mailappversionoverrides";
    private const string MailOverrides10 = "http://schemas.microsoft.com/office/mailappversionoverrides/1.0";
    private const string MailOverrides11 = "http://schemas.microsoft.com/office/mailappversionoverrides/1.1";
    private const string Xsi = "http://www.w3.org/2001/XMLSchema-instance";
    private const string AddInOrigin = "https://localhost:47831";
    private const string TaskPaneUrl = AddInOrigin + "/outlook/taskpane.html";

    private static readonly Lazy<XDocument> Manifest = new(() =>
        XDocument.Load(Path.Combine(AppContext.BaseDirectory, "outlook", "manifest.xml")));

    private static readonly XNamespace AppNs = XNamespace.Get(AppForOffice);
    private static readonly XNamespace BtNs = XNamespace.Get(BasicTypes);
    private static readonly XNamespace MailNs = XNamespace.Get(MailOverrides);
    private static readonly XNamespace Mail11Ns = XNamespace.Get(MailOverrides11);
    private static readonly XNamespace XsiNs = XNamespace.Get(Xsi);

    [Fact]
    public void RootIsAMailAppOnTheOfficeAddInNamespace()
    {
        var root = Document.Root ?? throw new InvalidOperationException("the manifest has no root element");

        Assert.Equal("OfficeApp", root.Name.LocalName);
        Assert.Equal(AppNs, root.Name.Namespace);
        Assert.Equal("MailApp", Type(root));

        Assert.Equal(BasicTypes, NamespaceDeclaration(root, "bt"));
        Assert.Equal(Xsi, NamespaceDeclaration(root, "xsi"));
        Assert.Equal(MailOverrides10, NamespaceDeclaration(root, "mailappor"));
    }

    [Theory]
    [InlineData("Id")]
    [InlineData("Version")]
    [InlineData("ProviderName")]
    [InlineData("DefaultLocale")]
    [InlineData("DisplayName")]
    [InlineData("Description")]
    [InlineData("IconUrl")]
    [InlineData("HighResolutionIconUrl")]
    [InlineData("SupportUrl")]
    [InlineData("AppDomains")]
    [InlineData("Hosts")]
    [InlineData("Requirements")]
    [InlineData("FormSettings")]
    [InlineData("Permissions")]
    [InlineData("Rule")]
    public void DeclaresEveryRequiredMailAppElement(string name)
    {
        Assert.NotNull(App(name));
    }

    [Fact]
    public void DeclaresExactlyOneMailVersionOverrides()
    {
        Assert.Single(Document.Root?.Elements(MailNs + "VersionOverrides") ?? []);
    }

    [Fact]
    public void UsesTheDefaultValueAttributeFormRatherThanTheLegacyElements()
    {
        // MailApp v1.1 takes <DisplayName DefaultValue="..."/>. The <DisplayNameDefaultValue>
        // spelling belongs to no element in this schema.
        Assert.Equal("ReplyBrief", App("DisplayName").Attribute("DefaultValue")?.Value);
        Assert.NotNull(App("Description").Attribute("DefaultValue"));
        Assert.Null(Document.Root?.Element(AppNs + "DisplayNameDefaultValue"));
        Assert.Null(Document.Root?.Element(AppNs + "DescriptionDefaultValue"));
    }

    [Fact]
    public void UsesAPlainGuidForTheProductId()
    {
        Assert.True(Guid.TryParse(App("Id").Value, out var id), "the product id must be a plain GUID");
        Assert.NotEqual(Guid.Empty, id);
    }

    [Fact]
    public void DeclaresTheMailboxHostRatherThanOutlook()
    {
        Assert.Equal(["Mailbox"], App("Hosts").Elements().Select(host => host.Attribute("Name")?.Value));
    }

    [Fact]
    public void ScopesEveryUrlToTheLoopbackOrigin()
    {
        Assert.Equal([AddInOrigin], App("AppDomains").Elements().Select(domain => domain.Value));

        foreach (var url in Urls())
        {
            Assert.True(
                url == AddInOrigin || url.StartsWith(AddInOrigin + "/", StringComparison.Ordinal),
                $"{url} is outside the add-in origin");
        }
    }

    [Fact]
    public void EveryUrlInTheDocumentIsHttps()
    {
        Assert.All(Urls(), url => Assert.StartsWith("https://", url));
    }

    [Fact]
    public void RequestsOnlyReadItem()
    {
        Assert.Equal("ReadItem", App("Permissions").Value.Trim());
    }

    [Fact]
    public void DeclaresTheFormSettingsItemReadSourceLocation()
    {
        var form = App("FormSettings").Element(AppNs + "Form")
            ?? throw new InvalidOperationException("<FormSettings> must contain a <Form>");

        Assert.Equal("ItemRead", Type(form));
        Assert.Equal(TaskPaneUrl, form.Element(AppNs + "DesktopSettings")?.Element(AppNs + "SourceLocation")?.Attribute("DefaultValue")?.Value);
    }

    [Fact]
    public void UsesTheStandardItemIsRuleAndNotAnInventedOne()
    {
        var rule = App("Rule").Elements().Single();

        Assert.Equal("RuleCollection", Type(App("Rule")));
        Assert.Equal("ItemIs", Type(rule));
        Assert.Equal("Message", rule.Attribute("ItemType")?.Value);
        Assert.Equal("Read", rule.Attribute("FormType")?.Value);
        Assert.DoesNotContain("RuleIsAmong", Text());
    }

    [Fact]
    public void NestsAVersionOverridesV1_1InsideV1_0ForPinnableTaskPanes()
    {
        var outer = OverridesRoot();
        var inner = Overrides();

        Assert.Equal(MailNs, outer.Name.Namespace);
        Assert.Equal("VersionOverridesV1_0", Type(outer));
        Assert.Equal(Mail11Ns, inner.Name.Namespace);
        Assert.Equal("VersionOverridesV1_1", Type(inner));
    }

    [Fact]
    public void RequiresMailboxOnePointFiveThroughTheBasicTypesSetElements()
    {
        var sets = Overrides().Element(Mail11Ns + "Requirements")?.Elements().SingleOrDefault()
            ?? throw new InvalidOperationException("<Requirements> must contain a <bt:Sets>");
        var set = sets.Elements().SingleOrDefault()
            ?? throw new InvalidOperationException("<bt:Sets> must contain a <bt:Set>");

        Assert.Equal("Sets", sets.Name.LocalName);
        Assert.Equal(BtNs, sets.Name.Namespace);
        Assert.Equal("Set", set.Name.LocalName);
        Assert.Equal(BtNs, set.Name.Namespace);
        Assert.Equal("Mailbox", set.Attribute("Name")?.Value);
        Assert.Equal("1.5", set.Attribute("MinVersion")?.Value ?? sets.Attribute("DefaultMinVersion")?.Value);
    }

    [Fact]
    public void UsesTheMessageReadCommandSurfaceExtensionPoint()
    {
        var host = Overrides().Element(Mail11Ns + "Hosts")?.Elements().SingleOrDefault()
            ?? throw new InvalidOperationException("<Hosts> must contain a single <Host>");
        var extensionPoint = host.Element(Mail11Ns + "DesktopFormFactor")?.Elements(Mail11Ns + "ExtensionPoint").SingleOrDefault()
            ?? throw new InvalidOperationException("<DesktopFormFactor> must contain an <ExtensionPoint>");

        Assert.Equal("MailHost", Type(host));
        Assert.Equal("MessageReadCommandSurface", Type(extensionPoint));
        Assert.DoesNotContain("PrimaryCommandSurface", Text());

        var action = extensionPoint.Descendants(Mail11Ns + "Action").SingleOrDefault()
            ?? throw new InvalidOperationException("the control must declare an <Action>");

        Assert.Equal("ShowTaskpane", Type(action));
        Assert.Equal(TaskPaneUrl, ResourceValue(action.Element(Mail11Ns + "SourceLocation")?.Attribute("resid")?.Value));
    }

    [Fact]
    public void DeclaresResourcesInsideTheVersionOverridesUsingTheStandardCollections()
    {
        var resources = Overrides().Element(Mail11Ns + "Resources")
            ?? throw new InvalidOperationException("<Resources> belongs inside the mail <VersionOverrides>");

        Assert.Null(Document.Root?.Element(AppNs + "Resources"));
        Assert.Equal(["Images", "Urls", "ShortStrings", "LongStrings"], resources.Elements().Select(child => child.Name.LocalName));
        Assert.All(resources.Elements(), child => Assert.Equal(BtNs, child.Name.Namespace));
    }

    [Fact]
    public void UsesBtImageForButtonIconsRatherThanBtUrl()
    {
        var control = Overrides().Descendants(Mail11Ns + "Control").SingleOrDefault()
            ?? throw new InvalidOperationException("the manifest must declare the capture button");
        var references = control.Descendants(BtNs + "Image").ToArray();

        Assert.Equal(["16", "32", "80"], references.Select(reference => reference.Attribute("size")?.Value));
        Assert.All(references, reference => Assert.StartsWith(AddInOrigin + "/outlook/", ResourceValue(reference.Attribute("resid")?.Value)));
        Assert.DoesNotContain("bt:Url size", Text());
    }

    [Fact]
    public void TheFunctionFileIsItsOwnPageRatherThanTheTaskPane()
    {
        var functionFile = Overrides().Descendants(Mail11Ns + "FunctionFile").SingleOrDefault()
            ?? throw new InvalidOperationException("<DesktopFormFactor> must declare a <FunctionFile>");

        var url = ResourceValue(functionFile.Attribute("resid")?.Value);

        Assert.EndsWith("/outlook/commands.html", url);
        Assert.NotEqual(TaskPaneUrl, url);
        Assert.True(File.Exists(Path.Combine(AddInAssets.DefaultDirectory, "commands.html")));
    }

    [Fact]
    public void TheSupportUrlPointsAtAServedSupportPage()
    {
        var url = App("SupportUrl").Attribute("DefaultValue")?.Value;

        Assert.Equal($"{AddInOrigin}/outlook/support.html", url);
        Assert.True(File.Exists(Path.Combine(AddInAssets.DefaultDirectory, "support.html")));
    }

    [Fact]
    public void DoesNotClaimPinningSupportItDoesNotDeclare()
    {
        // The nested v1.1 override is required by Mailbox 1.5, not by pinning. ReplyBrief ships no
        // TaskpaneId, so pinning is deliberately not declared and must not be implied anywhere.
        var localNames = Document.Descendants().Select(element => element.Name.LocalName).ToHashSet(StringComparer.Ordinal);

        Assert.DoesNotContain("TaskpaneId", localNames);
        Assert.DoesNotContain("SupportsPinning", localNames);
    }

    [Fact]
    public void EveryResidReferenceResolvesToADeclaredResource()
    {
        var referenced = ResidReferences();

        Assert.NotEmpty(referenced);
        Assert.All(referenced, resid => Assert.Contains(resid, ResourceIds()));
    }

    [Fact]
    public void EveryDeclaredResourceIsReferencedAndIdsAreUnique()
    {
        var declared = ResourceIds();
        var referenced = ResidReferences().ToHashSet(StringComparer.Ordinal);

        Assert.Equal(declared.Length, declared.Distinct(StringComparer.Ordinal).Count());
        Assert.All(declared, id => Assert.Contains(id, referenced));
    }

    [Fact]
    public void DoesNotUseInventedResourceElements()
    {
        Assert.DoesNotContain("resid ", Text());
        Assert.DoesNotContain(Document.Descendants().Select(element => element.Name.LocalName), name => name == "resid");
    }

    [Fact]
    public void TheManifestPointsAtTheTaskPaneTheHostActuallyServes()
    {
        var resid = Overrides().Descendants(Mail11Ns + "SourceLocation").SingleOrDefault()?.Attribute("resid")?.Value;

        Assert.Equal(TaskPaneUrl, ResourceValue(resid));
        Assert.True(File.Exists(Path.Combine(AddInAssets.DefaultDirectory, "taskpane.html")));
    }

    private static XDocument Document => Manifest.Value;

    private static string Text() => Document.ToString();

    private static XElement App(string name) =>
        Document.Root?.Element(AppNs + name)
        ?? throw new InvalidOperationException($"the manifest is missing the required <{name}> element");

    private static XElement OverridesRoot() =>
        Document.Root?.Elements(MailNs + "VersionOverrides").SingleOrDefault()
        ?? throw new InvalidOperationException("the manifest is missing the mail <VersionOverrides> element");

    private static XElement Overrides() =>
        OverridesRoot().Elements().SingleOrDefault()
        ?? throw new InvalidOperationException("<VersionOverrides> must contain a single nested <VersionOverrides>");

    private static string? Type(XElement element) => element.Attribute(XsiNs + "type")?.Value;

    private static string NamespaceDeclaration(XElement element, string prefix) =>
        element.Attributes().Single(attribute => attribute.IsNamespaceDeclaration && attribute.Name.LocalName == prefix).Value;

    private static IEnumerable<XElement> BasicTypeElements() =>
        Document.Descendants().Where(element => element.Name.Namespace == BtNs);

    private static string[] ResourceIds() =>
        BasicTypeElements().Select(element => element.Attribute("id")?.Value).OfType<string>().ToArray();

    private static string[] ResidReferences() =>
        Document.Descendants().Select(element => element.Attribute("resid")?.Value).OfType<string>().ToArray();

    private static string ResourceValue(string? id) =>
        BasicTypeElements()
            .Where(element => element.Attribute("id")?.Value == id)
            .Select(element => element.Attribute("DefaultValue")?.Value ?? string.Empty)
            .Single();

    private static IEnumerable<string> Urls() =>
        Document.Descendants()
            .Select(element => element.Attribute("DefaultValue")?.Value ?? element.Value)
            .Where(value => value.StartsWith("http", StringComparison.OrdinalIgnoreCase));
}