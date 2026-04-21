using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Markdig.Renderers.Html;
using Markdig.Syntax;
using md_to_docx_sync.Enums;
using md_to_docx_sync.Models;
using md_to_docx_sync.Services;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

// Get command line arguments

string? inputMDFilePath = args.ElementAtOrDefault(0);
if (string.IsNullOrEmpty(inputMDFilePath))
{
    Console.Error.WriteLine("[ERROR] Please provide the input MD file path as the first argument.");
    return;
}

string? inputDOCXTemplateFilePath = args.ElementAtOrDefault(1);
if (string.IsNullOrEmpty(inputDOCXTemplateFilePath))
{
    Console.Error.WriteLine("[ERROR] Please provide the input DOCX template file path as the second argument.");
    return;
}

string? outputDOCXFilePath = args.ElementAtOrDefault(2);
if (string.IsNullOrEmpty(outputDOCXFilePath))
{
    Console.Error.WriteLine("[ERROR] Please provide the output DOCX file path as the third argument.");
    return;
}

// Get the directory containing the markdown file for resolving relative paths
string? inputMDFileDir = Path.GetDirectoryName(Path.GetFullPath(inputMDFilePath));

// Read and parse Markdown File
var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().UseCustomContainers().UseSmartyPants().Build();

// Pre-process: replace [@key] / [@key1; @key2] with «key» / «key1;key2» placeholders
// so Markdig's link parser doesn't consume the brackets and break citation detection.
var rawMarkdown = File.ReadAllText(inputMDFilePath);
var citationPreProcess = new Regex(@"\[@([a-zA-Z0-9_:-]+(?:\s*;\s*@[a-zA-Z0-9_:-]+)*)\]");
rawMarkdown = citationPreProcess.Replace(rawMarkdown, match =>
{
    // Strip the leading @ from each key and join with ;
    var keys = match.Groups[1].Value
        .Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
        .Select(k => k.TrimStart('@'));
    return $"\u00AB{string.Join(";", keys)}\u00BB";
});

var markdownContent = Markdown.Parse(rawMarkdown, pipeline);

// Create DocxContent from markdown
DocxContent docxContent = new();
DocxContent mainDocxContent = docxContent;
DocxContent frontMatterDocxContent = new();
DocxContent appendixDocxContent = new();
List<Block> appendixBlocks = new();

const string targetIdPattern = "^[A-Za-z][A-Za-z0-9:_-]*$";
var definedCrossReferenceTargets = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
var referencedCrossReferenceTargets = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
var crossReferenceErrors = new List<string>();

bool IsValidTargetId(string targetId)
{
    return Regex.IsMatch(targetId, targetIdPattern);
}

void AddCrossReferenceError(string message)
{
    crossReferenceErrors.Add(message);
}

void RegisterCrossReferenceTarget(string targetId, string location)
{
    if (!IsValidTargetId(targetId))
    {
        AddCrossReferenceError($"Invalid target ID '{targetId}' at {location}. Allowed pattern: {targetIdPattern}");
        return;
    }

    if (definedCrossReferenceTargets.TryGetValue(targetId, out string? existingLocation))
    {
        AddCrossReferenceError($"Duplicate target ID '{targetId}' at {location}. Already defined at {existingLocation}.");
        return;
    }

    definedCrossReferenceTargets[targetId] = location;
}

void RegisterCrossReferenceReference(string targetId, string location)
{
    if (string.IsNullOrWhiteSpace(targetId))
    {
        AddCrossReferenceError($"Empty internal cross-reference target at {location}. Use [text](#target-id).");
        return;
    }

    if (!IsValidTargetId(targetId))
    {
        AddCrossReferenceError($"Invalid referenced target ID '{targetId}' at {location}. Allowed pattern: {targetIdPattern}");
        return;
    }

    if (!referencedCrossReferenceTargets.TryGetValue(targetId, out var locations))
    {
        locations = new List<string>();
        referencedCrossReferenceTargets[targetId] = locations;
    }

    locations.Add(location);
}

void ValidateCrossReferencesOrThrow()
{
    foreach (var (targetId, locations) in referencedCrossReferenceTargets)
    {
        if (definedCrossReferenceTargets.ContainsKey(targetId))
        {
            continue;
        }

        string joinedLocations = string.Join("; ", locations.Distinct(StringComparer.Ordinal));
        AddCrossReferenceError($"Unresolved cross-reference target '{targetId}' referenced at {joinedLocations}.");
    }

    if (crossReferenceErrors.Count == 0)
    {
        return;
    }

    foreach (string error in crossReferenceErrors)
    {
        Console.Error.WriteLine($"[ERROR] {error}");
    }

    throw new InvalidOperationException("Cross-reference validation failed.");
}

(bool markerFound, string textWithoutMarker, string? targetId) ParseTrailingTargetMarker(string text, string location)
{
    string working = text.TrimEnd();
    int markerStart = working.LastIndexOf("{#", StringComparison.Ordinal);

    if (markerStart < 0)
    {
        return (false, working, null);
    }

    if (!working.EndsWith("}", StringComparison.Ordinal))
    {
        AddCrossReferenceError($"Malformed target marker in {location}: '{working}'. Expected trailing '{{#id}}'.");
        return (true, working, null);
    }

    if (markerStart == 0 || !char.IsWhiteSpace(working[markerStart - 1]))
    {
        AddCrossReferenceError($"Malformed target marker in {location}: '{working}'. Marker must be preceded by whitespace.");
        return (true, working, null);
    }

    string targetIdCandidate = working.Substring(markerStart + 2, working.Length - markerStart - 3).Trim();
    string visibleText = working.Substring(0, markerStart).TrimEnd();

    if (string.IsNullOrWhiteSpace(targetIdCandidate))
    {
        AddCrossReferenceError($"Malformed target marker in {location}: '{working}'. Target ID cannot be empty.");
        return (true, visibleText, null);
    }

    if (!IsValidTargetId(targetIdCandidate))
    {
        AddCrossReferenceError($"Invalid target ID '{targetIdCandidate}' in {location}. Allowed pattern: {targetIdPattern}");
        return (true, visibleText, null);
    }

    return (true, visibleText, targetIdCandidate);
}

string? ConsumeTrailingTargetMarkerFromInline(Markdig.Syntax.Inlines.ContainerInline? container, string location)
{
    if (container == null)
    {
        return null;
    }

    var lastInline = container.LastChild;
    while (lastInline is Markdig.Syntax.Inlines.LineBreakInline)
    {
        lastInline = lastInline.PreviousSibling;
    }

    if (lastInline is Markdig.Syntax.Inlines.LiteralInline literalInline)
    {
        var parseResult = ParseTrailingTargetMarker(literalInline.Content.ToString(), location);
        if (!parseResult.markerFound)
        {
            return null;
        }

        literalInline.Content = new Markdig.Helpers.StringSlice(parseResult.textWithoutMarker);
        return parseResult.targetId;
    }

    // Keep behavior explicit: marker stripping is only supported when the marker is in trailing literal text.
    var plainText = GetPlainText(container);
    var fallbackParseResult = ParseTrailingTargetMarker(plainText, location);
    if (fallbackParseResult.markerFound && fallbackParseResult.targetId != null)
    {
        AddCrossReferenceError($"Target marker in {location} must be plain trailing text (for example: ... {{#{fallbackParseResult.targetId}}}).");
    }

    return null;
}

// Helper to parse caption with optional width and optional target marker (e.g., "Caption text {#fig-id} |> 400")
(string caption, int? width, string? targetId) ParseCaptionMetadata(string? rawCaption, string location)
{
    if (string.IsNullOrEmpty(rawCaption))
        return (string.Empty, null, null);

    string working = rawCaption.Trim();
    int? width = null;

    var widthMatch = Regex.Match(working, @"^(.+?)\s*\|>\s*(\d+)\s*$");
    if (widthMatch.Success)
    {
        working = widthMatch.Groups[1].Value.TrimEnd();
        width = int.Parse(widthMatch.Groups[2].Value);
    }

    var markerResult = ParseTrailingTargetMarker(working, location);
    if (markerResult.markerFound)
    {
        working = markerResult.textWithoutMarker;
    }

    return (working.Trim(), width, markerResult.targetId);
}

// Helper to get plain text from inline elements (for captions)
string GetPlainText(Markdig.Syntax.Inlines.ContainerInline? container)
{
    if (container == null) return string.Empty;
    var sb = new System.Text.StringBuilder();
    foreach (var inline in container)
    {
        if (inline is Markdig.Syntax.Inlines.LiteralInline literal)
            sb.Append(literal.Content);
        else if (inline is Markdig.Syntax.Inlines.ContainerInline nested)
            sb.Append(GetPlainText(nested));
    }
    return sb.ToString();
}

string? GetTargetIdFromAttributes(MarkdownObject markdownObject)
{
    var attributes = markdownObject.TryGetAttributes();
    if (attributes == null || string.IsNullOrWhiteSpace(attributes.Id))
    {
        return null;
    }

    return attributes.Id;
}

void WarnContentOutsideFigureBlock(string contentType, MarkdownObject markdownObject)
{
    string lineText = markdownObject.Line >= 0
        ? (markdownObject.Line + 1).ToString()
        : "unknown";

    Console.WriteLine($"[WARNING] {contentType} is used without figure block (^^^) at line {lineText}.");
}

// Helper to parse references YAML block into ReferenceSource list
List<ReferenceSource> ParseReferencesYaml(string yamlContent)
{
    var deserializer = new DeserializerBuilder()
        .WithNamingConvention(CamelCaseNamingConvention.Instance)
        .Build();

    var yamlData = deserializer.Deserialize<List<Dictionary<string, object>>>(yamlContent);
    if (yamlData == null) return new List<ReferenceSource>();

    var references = new List<ReferenceSource>();
    foreach (var entry in yamlData)
    {
        if (!entry.TryGetValue("key", out var keyObj) || !entry.TryGetValue("type", out var typeObj))
        {
            Console.WriteLine("[WARNING] Reference entry missing 'key' or 'type' field, skipping.");
            continue;
        }

        var reference = new ReferenceSource
        {
            Key = keyObj.ToString()!,
            Type = typeObj.ToString()!
        };

        foreach (var (fieldName, fieldValue) in entry)
        {
            if (fieldName == "key" || fieldName == "type") continue;
            reference.Fields[fieldName] = fieldValue?.ToString() ?? string.Empty;
        }

        references.Add(reference);
    }

    return references;
}

// Regex for detecting «key1;key2» citation placeholders (produced by pre-processing step)
var citationPattern = new Regex(@"\u00AB([a-zA-Z0-9_;:-]+)\u00BB");

// Helper to process literal text that may contain citation placeholders «key1;key2»
void AddTextWithCitations(string text, bool isBold, bool isItalic, bool isUnderline, bool isStrikethrough, bool isLiteral)
{
    int lastIndex = 0;
    foreach (Match match in citationPattern.Matches(text))
    {
        // Add text before the citation, stripping any trailing space
        // (Word's citation field rendering already provides the necessary spacing)
        if (match.Index > lastIndex)
        {
            var beforeText = text.Substring(lastIndex, match.Index - lastIndex).TrimEnd(' ');
            if (beforeText.Length > 0)
                docxContent.AddText(beforeText, isBold, isItalic, isUnderline, isStrikethrough, isLiteral);
        }

        // Parse citation keys from placeholder (keys are ;-separated)
        var citationKeys = match.Groups[1].Value
            .Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        docxContent.AddCitation(citationKeys);

        lastIndex = match.Index + match.Length;
    }

    // Add remaining text after last citation
    if (lastIndex < text.Length)
    {
        docxContent.AddText(text.Substring(lastIndex), isBold, isItalic, isUnderline, isStrikethrough, isLiteral);
    }
}

bool IsExternalHyperlinkTarget(string? linkTarget)
{
    if (string.IsNullOrWhiteSpace(linkTarget))
    {
        return false;
    }

    return Uri.TryCreate(linkTarget, UriKind.Absolute, out _);
}

// Process inline content with formatting
void ProcessInlines(Markdig.Syntax.Inlines.ContainerInline? container, bool isBold = false, bool isItalic = false, bool isUnderline = false, bool isStrikethrough = false,
    bool allowBold = true, bool allowItalic = true, bool allowUnderline = true, bool allowStrikethrough = true)
{
    if (container == null) return;

    foreach (var inline in container)
    {
        if (inline is Markdig.Syntax.Inlines.LiteralInline literal)
        {
            AddTextWithCitations(literal.Content.ToString(), isBold, isItalic, isUnderline, isStrikethrough, false);
        }
        else if (inline is Markdig.Syntax.Inlines.EmphasisInline emphasis)
        {
            bool newBold = isBold;
            bool newItalic = isItalic;
            bool newStrikethrough = isStrikethrough;
            bool fmtDisabled = false;
            string delim = new string(emphasis.DelimiterChar, emphasis.DelimiterCount);

            if (emphasis.DelimiterChar == '~' && emphasis.DelimiterCount == 2)
            {
                if (allowStrikethrough) newStrikethrough = true;
                else fmtDisabled = true;
            }
            else if (emphasis.DelimiterCount == 2)
            {
                if (allowBold) newBold = true;
                else fmtDisabled = true;
            }
            else
            {
                if (allowItalic) newItalic = true;
                else fmtDisabled = true;
            }

            if (fmtDisabled) docxContent.AddText(delim, isBold, isItalic, isUnderline, isStrikethrough);
            ProcessInlines(emphasis, newBold, newItalic, isUnderline, newStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
            if (fmtDisabled) docxContent.AddText(delim, isBold, isItalic, isUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.CodeInline codeInline)
        {
            docxContent.AddText(codeInline.Content, isBold, isItalic, isUnderline, isStrikethrough, isLiteral: true);
        }
        else if (inline is Markdig.Syntax.Inlines.LinkInline linkInline && !linkInline.IsImage)
        {
            var linkTarget = linkInline.Url;
            if (!string.IsNullOrEmpty(linkTarget) && linkTarget.StartsWith("#", StringComparison.Ordinal))
            {
                string targetId = linkTarget.Substring(1);
                RegisterCrossReferenceReference(targetId, $"link '{linkTarget}'");

                if (IsValidTargetId(targetId))
                {
                    docxContent.BeginInternalHyperlink(targetId);
                    try
                    {
                        ProcessInlines(linkInline, isBold, isItalic, isUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
                    }
                    finally
                    {
                        docxContent.EndHyperlink();
                    }

                    continue;
                }
            }

            if (IsExternalHyperlinkTarget(linkTarget))
            {
                docxContent.BeginExternalHyperlink(linkTarget!);
                try
                {
                    ProcessInlines(linkInline, isBold, isItalic, isUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
                }
                finally
                {
                    docxContent.EndHyperlink();
                }

                continue;
            }

            // Non-anchor links are rendered as plain text (link text only) to preserve existing behavior.
            ProcessInlines(linkInline, isBold, isItalic, isUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
        }
        else if (inline is Markdig.Extensions.SmartyPants.SmartyPant smartyPant)
        {
            string smartText = smartyPant.Type switch
            {
                Markdig.Extensions.SmartyPants.SmartyPantType.Quote => "\u2018",
                Markdig.Extensions.SmartyPants.SmartyPantType.LeftQuote => "\u2018",      // '
                Markdig.Extensions.SmartyPants.SmartyPantType.RightQuote => "\u2019",     // '
                Markdig.Extensions.SmartyPants.SmartyPantType.DoubleQuote => "\u201C",
                Markdig.Extensions.SmartyPants.SmartyPantType.LeftDoubleQuote => "\u201C", // "
                Markdig.Extensions.SmartyPants.SmartyPantType.RightDoubleQuote => "\u201D", // "
                Markdig.Extensions.SmartyPants.SmartyPantType.Ellipsis => "\u2026",       // …
                Markdig.Extensions.SmartyPants.SmartyPantType.Dash2 => "\u2013",          // –
                Markdig.Extensions.SmartyPants.SmartyPantType.Dash3 => "\u2014",          // —
                _ => smartyPant.ToString()
            };
            docxContent.AddText(smartText, isBold, isItalic, isUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.HtmlInline htmlInline)
        {
            // <u>/<\/u> are handled in ProcessInlinesWithUnderline; emit literally here
            docxContent.AddText(htmlInline.Tag, isBold, isItalic, isUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.LineBreakInline)
        {
            docxContent.AddLineBreak();
        }
        else if (inline is Markdig.Syntax.Inlines.ContainerInline nestedContainer)
        {
            ProcessInlines(nestedContainer, isBold, isItalic, isUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
        }
    }
}

// Process inline content with HTML underline tag detection
void ProcessInlinesWithUnderline(Markdig.Syntax.Inlines.ContainerInline? container, bool isBold = false, bool isItalic = false, bool isStrikethrough = false,
    bool allowBold = true, bool allowItalic = true, bool allowUnderline = true, bool allowStrikethrough = true)
{
    if (container == null) return;

    bool currentUnderline = false;
    var inlines = container.ToList();

    for (int i = 0; i < inlines.Count; i++)
    {
        var inline = inlines[i];

        if (inline is Markdig.Syntax.Inlines.HtmlInline htmlInline)
        {
            string tag = htmlInline.Tag.ToLowerInvariant();
            if (tag == "<u>")
            {
                if (allowUnderline) currentUnderline = true;
                else docxContent.AddText(htmlInline.Tag, isBold, isItalic, currentUnderline, isStrikethrough);
            }
            else if (tag == "</u>")
            {
                if (allowUnderline) currentUnderline = false;
                else docxContent.AddText(htmlInline.Tag, isBold, isItalic, currentUnderline, isStrikethrough);
            }
        }
        else if (inline is Markdig.Syntax.Inlines.LiteralInline literal)
        {
            AddTextWithCitations(literal.Content.ToString(), isBold, isItalic, currentUnderline, isStrikethrough, false);
        }
        else if (inline is Markdig.Syntax.Inlines.EmphasisInline emphasis)
        {
            bool newBold = isBold;
            bool newItalic = isItalic;
            bool newStrikethrough = isStrikethrough;
            bool fmtDisabled = false;
            string delim = new string(emphasis.DelimiterChar, emphasis.DelimiterCount);

            if (emphasis.DelimiterChar == '~' && emphasis.DelimiterCount == 2)
            {
                if (allowStrikethrough) newStrikethrough = true;
                else fmtDisabled = true;
            }
            else if (emphasis.DelimiterCount == 2)
            {
                if (allowBold) newBold = true;
                else fmtDisabled = true;
            }
            else
            {
                if (allowItalic) newItalic = true;
                else fmtDisabled = true;
            }

            if (fmtDisabled) docxContent.AddText(delim, isBold, isItalic, currentUnderline, isStrikethrough);
            ProcessInlines(emphasis, newBold, newItalic, currentUnderline, newStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
            if (fmtDisabled) docxContent.AddText(delim, isBold, isItalic, currentUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.CodeInline codeInline)
        {
            docxContent.AddText(codeInline.Content, isBold, isItalic, currentUnderline, isStrikethrough, isLiteral: true);
        }
        else if (inline is Markdig.Syntax.Inlines.LinkInline linkInline && !linkInline.IsImage)
        {
            var linkTarget = linkInline.Url;
            if (!string.IsNullOrEmpty(linkTarget) && linkTarget.StartsWith("#", StringComparison.Ordinal))
            {
                string targetId = linkTarget.Substring(1);
                RegisterCrossReferenceReference(targetId, $"link '{linkTarget}'");

                if (IsValidTargetId(targetId))
                {
                    docxContent.BeginInternalHyperlink(targetId);
                    try
                    {
                        ProcessInlines(linkInline, isBold, isItalic, currentUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
                    }
                    finally
                    {
                        docxContent.EndHyperlink();
                    }

                    continue;
                }
            }

            if (IsExternalHyperlinkTarget(linkTarget))
            {
                docxContent.BeginExternalHyperlink(linkTarget!);
                try
                {
                    ProcessInlines(linkInline, isBold, isItalic, currentUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
                }
                finally
                {
                    docxContent.EndHyperlink();
                }

                continue;
            }

            // Non-anchor links are rendered as plain text (link text only) to preserve existing behavior.
            ProcessInlines(linkInline, isBold, isItalic, currentUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
        }
        else if (inline is Markdig.Extensions.SmartyPants.SmartyPant smartyPant)
        {
            string smartText = smartyPant.Type switch
            {
                Markdig.Extensions.SmartyPants.SmartyPantType.Quote => "'",
                Markdig.Extensions.SmartyPants.SmartyPantType.LeftQuote => "\u2018",      // '
                Markdig.Extensions.SmartyPants.SmartyPantType.RightQuote => "\u2019",     // '
                Markdig.Extensions.SmartyPants.SmartyPantType.DoubleQuote => "\"",
                Markdig.Extensions.SmartyPants.SmartyPantType.LeftDoubleQuote => "\u201C", // "
                Markdig.Extensions.SmartyPants.SmartyPantType.RightDoubleQuote => "\u201D", // "
                Markdig.Extensions.SmartyPants.SmartyPantType.Ellipsis => "\u2026",       // …
                Markdig.Extensions.SmartyPants.SmartyPantType.Dash2 => "\u2013",          // –
                Markdig.Extensions.SmartyPants.SmartyPantType.Dash3 => "\u2014",          // —
                _ => smartyPant.ToString()
            };
            docxContent.AddText(smartText, isBold, isItalic, currentUnderline, isStrikethrough);
        }
        else if (inline is Markdig.Syntax.Inlines.LineBreakInline)
        {
            docxContent.AddLineBreak();
        }
        else if (inline is Markdig.Syntax.Inlines.ContainerInline nestedContainer)
        {
            ProcessInlines(nestedContainer, isBold, isItalic, currentUnderline, isStrikethrough, allowBold, allowItalic, allowUnderline, allowStrikethrough);
        }
    }
}

// Process list items recursively with nesting level tracking
void ProcessListBlock(ListBlock listBlock, int nestingLevel = 0)
{
    bool isOrdered = listBlock.IsOrdered;
    DocxListType listType = isOrdered ? DocxListType.Numbered : DocxListType.Bulleted;

    // Only create a new list at the top level
    if (nestingLevel == 0)
    {
        docxContent.NewList(listType);
    }

    int listItemNumber = 0;
    foreach (var item in listBlock)
    {
        if (item is ListItemBlock listItemBlock)
        {
            listItemNumber++;

            foreach (var child in listItemBlock)
            {
                if (child is ParagraphBlock paragraphBlock)
                {
                    string location = $"list item {listItemNumber} at level {nestingLevel + 1}";
                    string? listItemTargetId = GetTargetIdFromAttributes(paragraphBlock)
                        ?? GetTargetIdFromAttributes(listItemBlock)
                        ?? ConsumeTrailingTargetMarkerFromInline(paragraphBlock.Inline, location);
                    if (!string.IsNullOrEmpty(listItemTargetId))
                    {
                        RegisterCrossReferenceTarget(listItemTargetId, location);
                    }

                    docxContent.NewListItem(level: nestingLevel, bookmarkName: listItemTargetId);
                    ProcessInlinesWithUnderline(paragraphBlock.Inline);
                }
                else if (child is ListBlock nestedList)
                {
                    // Process nested list with increased nesting level
                    ProcessListBlock(nestedList, nestingLevel + 1);
                }
            }
        }
    }
}

void ProcessTableCellContent(Markdig.Extensions.Tables.TableCell tableCell)
{
    bool hasRenderedLine = false;

    void EnsureLineStart()
    {
        if (hasRenderedLine)
        {
            docxContent.AddLineBreak();
        }
        hasRenderedLine = true;
    }

    void ProcessTableCellBlock(Block block)
    {
        switch (block)
        {
            case ParagraphBlock paragraphBlock:
                EnsureLineStart();
                ProcessInlinesWithUnderline(paragraphBlock.Inline);
                break;

            case QuoteBlock quoteBlock:
                foreach (var quoteChild in quoteBlock)
                {
                    if (quoteChild is ParagraphBlock quotedParagraph)
                    {
                        EnsureLineStart();
                        docxContent.AddText("> ");
                        ProcessInlinesWithUnderline(quotedParagraph.Inline);
                    }
                    else
                    {
                        ProcessTableCellBlock(quoteChild);
                    }
                }
                break;

            case ListBlock listBlock:
                int orderedIndex = 1;
                foreach (var listItem in listBlock)
                {
                    if (listItem is not ListItemBlock listItemBlock) continue;

                    foreach (var listItemChild in listItemBlock)
                    {
                        if (listItemChild is ParagraphBlock listItemParagraph)
                        {
                            EnsureLineStart();
                            docxContent.AddText(listBlock.IsOrdered ? $"{orderedIndex}. " : "• ");
                            ProcessInlinesWithUnderline(listItemParagraph.Inline);
                        }
                        else
                        {
                            ProcessTableCellBlock(listItemChild);
                        }
                    }

                    if (listBlock.IsOrdered) orderedIndex++;
                }
                break;

            case ContainerBlock containerBlock:
                foreach (var childBlock in containerBlock)
                {
                    ProcessTableCellBlock(childBlock);
                }
                break;

            case LeafBlock leafBlock when leafBlock.Inline != null:
                EnsureLineStart();
                ProcessInlinesWithUnderline(leafBlock.Inline);
                break;
        }
    }

    foreach (var cellContent in tableCell)
    {
        ProcessTableCellBlock(cellContent);
    }
}

void ProcessTable(Markdig.Extensions.Tables.Table table, string? caption = null, int? widthPercent = null, string? captionTargetId = null)
{
    JustificationValues? GetColumnAlignment(int columnIndex)
    {
        if (table.ColumnDefinitions == null || columnIndex < 0 || columnIndex >= table.ColumnDefinitions.Count)
        {
            return null;
        }

        return table.ColumnDefinitions[columnIndex].Alignment switch
        {
            Markdig.Extensions.Tables.TableColumnAlign.Left => JustificationValues.Left,
            Markdig.Extensions.Tables.TableColumnAlign.Center => JustificationValues.Center,
            Markdig.Extensions.Tables.TableColumnAlign.Right => JustificationValues.Right,
            _ => null
        };
    }

    docxContent.NewTable(caption, widthPercent, captionTargetId);

    // Tracks pending vertical merges per column index (remaining continuation rows).
    var activeVerticalMerges = new Dictionary<int, int>();

    foreach (var row in table)
    {
        if (row is not Markdig.Extensions.Tables.TableRow tableRow)
        {
            continue;
        }

        docxContent.NewRow();
        int columnIndex = 0;

        foreach (var cell in tableRow)
        {
            if (cell is not Markdig.Extensions.Tables.TableCell tableCell)
            {
                continue;
            }

            while (activeVerticalMerges.TryGetValue(columnIndex, out int pendingContinuations) && pendingContinuations > 0)
            {
                docxContent.NewCell(verticalMergeContinue: true);
                pendingContinuations--;

                if (pendingContinuations == 0) activeVerticalMerges.Remove(columnIndex);
                else activeVerticalMerges[columnIndex] = pendingContinuations;

                columnIndex++;
            }

            int columnSpan = Math.Max(1, tableCell.ColumnSpan);
            int rowSpan = tableCell.RowSpan;

            bool isVerticalMergeStart = rowSpan > 1;
            bool isVerticalMergeContinue = rowSpan <= 0;
            JustificationValues? columnAlignment = GetColumnAlignment(columnIndex);

            docxContent.NewCell(
                columnSpan: columnSpan,
                verticalMergeStart: isVerticalMergeStart,
                verticalMergeContinue: isVerticalMergeContinue,
                alignment: columnAlignment
            );

            if (isVerticalMergeStart)
            {
                int continuationRows = rowSpan - 1;
                for (int i = 0; i < columnSpan; i++)
                {
                    activeVerticalMerges[columnIndex + i] = continuationRows;
                }
            }

            if (!isVerticalMergeContinue)
            {
                ProcessTableCellContent(tableCell);
            }

            columnIndex += columnSpan;
        }

        while (activeVerticalMerges.TryGetValue(columnIndex, out int trailingContinuations) && trailingContinuations > 0)
        {
            docxContent.NewCell(verticalMergeContinue: true);
            trailingContinuations--;

            if (trailingContinuations == 0) activeVerticalMerges.Remove(columnIndex);
            else activeVerticalMerges[columnIndex] = trailingContinuations;

            columnIndex++;
        }
    }
}

string GetNormalizedTopLevelHeadingText(HeadingBlock headingBlock)
{
    string headingText = GetPlainText(headingBlock.Inline).Trim();
    headingText = Regex.Replace(headingText, @"\s+\{#[^}]+\}\s*$", string.Empty).Trim();
    return headingText;
}

(int startIndex, int endExclusive)? FindTopLevelSectionRange(IReadOnlyList<Block> topLevelBlocks, string headingText)
{
    for (int i = 0; i < topLevelBlocks.Count; i++)
    {
        if (topLevelBlocks[i] is not HeadingBlock headingBlock || headingBlock.Level != 1)
        {
            continue;
        }

        if (!string.Equals(GetNormalizedTopLevelHeadingText(headingBlock), headingText, StringComparison.OrdinalIgnoreCase))
        {
            continue;
        }

        int sectionEndExclusive = i + 1;
        while (sectionEndExclusive < topLevelBlocks.Count)
        {
            if (topLevelBlocks[sectionEndExclusive] is HeadingBlock nextHeadingBlock && nextHeadingBlock.Level == 1)
            {
                break;
            }

            sectionEndExclusive++;
        }

        return (i, sectionEndExclusive);
    }

    return null;
}

List<Block> ExtractTopLevelSection(List<Block> topLevelBlocks, string headingText)
{
    var sectionRange = FindTopLevelSectionRange(topLevelBlocks, headingText);

    if (!sectionRange.HasValue)
    {
        return new List<Block>();
    }

    int sectionStart = sectionRange.Value.startIndex;
    int sectionCount = sectionRange.Value.endExclusive - sectionRange.Value.startIndex;
    var extractedSectionBlocks = topLevelBlocks.GetRange(sectionStart, sectionCount);
    topLevelBlocks.RemoveRange(sectionStart, sectionCount);

    return extractedSectionBlocks;
}

(List<Block> remainingBlocks, List<Block> acknowledgementsBlocks, List<Block> abstractBlocks, List<Block> appendixBlocks) SplitTopLevelDocumentSections(MarkdownDocument document)
{
    var topLevelBlocks = document.Cast<Block>().ToList();
    var acknowledgementsBlocks = ExtractTopLevelSection(topLevelBlocks, "Acknowledgements");
    var abstractBlocks = ExtractTopLevelSection(topLevelBlocks, "Abstract");
    var appendixBlocks = ExtractTopLevelSection(topLevelBlocks, "Appendix");

    return (topLevelBlocks, acknowledgementsBlocks, abstractBlocks, appendixBlocks);
}

void ProcessTopLevelBlocks(IEnumerable<Block> blocks)
{
    foreach (var topLevelBlock in blocks)
    {
        ProcessMarkdownNode(topLevelBlock);
    }
}

const string frontMatterHeadingStyleName = "Heading (Non-Numbered)";
const string frontMatterParagraphStyleName = "Normal";

void ProcessFrontMatterSectionBlocks(List<Block> sectionBlocks, string sectionLabel)
{
    bool renderedSectionHeading = false;

    foreach (var block in sectionBlocks)
    {
        if (!renderedSectionHeading && block is HeadingBlock headingBlock && headingBlock.Level == 1)
        {
            string location = $"{sectionLabel} heading";
            string? headingTargetId = GetTargetIdFromAttributes(headingBlock)
                ?? ConsumeTrailingTargetMarkerFromInline(headingBlock.Inline, location);

            if (!string.IsNullOrEmpty(headingTargetId))
            {
                RegisterCrossReferenceTarget(headingTargetId, location);
            }

            docxContent.NewParagraphWithStyleName(frontMatterHeadingStyleName, headingTargetId);
            ProcessInlinesWithUnderline(headingBlock.Inline);
            renderedSectionHeading = true;
            continue;
        }

        if (block is ParagraphBlock paragraphBlock)
        {
            docxContent.NewParagraphWithStyleName(frontMatterParagraphStyleName);
            ProcessInlinesWithUnderline(paragraphBlock.Inline);
            continue;
        }

        ProcessMarkdownNode(block);
    }
}

void ProcessFrontMatterSections(List<Block> acknowledgementsBlocks, List<Block> abstractBlocks)
{
    bool hasRenderedFrontMatterSection = false;

    if (acknowledgementsBlocks.Count > 0)
    {
        ProcessFrontMatterSectionBlocks(acknowledgementsBlocks, "Acknowledgements");
        hasRenderedFrontMatterSection = true;
    }

    if (abstractBlocks.Count > 0)
    {
        if (hasRenderedFrontMatterSection)
        {
            docxContent.AddPageBreak();
        }

        ProcessFrontMatterSectionBlocks(abstractBlocks, "Abstract");
        hasRenderedFrontMatterSection = true;
    }
}

bool IsHeading1Paragraph(Paragraph paragraph)
{
    var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
    return string.Equals(styleId, "Heading1", StringComparison.OrdinalIgnoreCase);
}

string GetNormalizedParagraphText(Paragraph paragraph)
{
    string paragraphText = string.Concat(paragraph.Descendants<Text>().Select(text => text.Text)).Trim();
    paragraphText = Regex.Replace(paragraphText, @"\s+\{#[^}]+\}\s*$", string.Empty).Trim();
    return paragraphText;
}

bool ParagraphContainsBibliographyField(Paragraph paragraph)
{
    bool hasSimpleField = paragraph.Descendants<SimpleField>().Any(simpleField =>
        simpleField.Instruction?.Value?.IndexOf("BIBLIOGRAPHY", StringComparison.OrdinalIgnoreCase) >= 0);

    if (hasSimpleField)
    {
        return true;
    }

    return paragraph.Descendants<FieldCode>().Any(fieldCode =>
        fieldCode.InnerText.IndexOf("BIBLIOGRAPHY", StringComparison.OrdinalIgnoreCase) >= 0);
}

OpenXmlElement? GetDirectBodyChild(Body body, OpenXmlElement element)
{
    OpenXmlElement? current = element;
    while (current?.Parent != null && current.Parent != body)
    {
        current = current.Parent;
    }

    return current?.Parent == body ? current : null;
}

OpenXmlElement? FindAppendixInsertAnchorAfterReferences(Body body)
{
    var bibliographyParagraph = body
        .Descendants<Paragraph>()
        .LastOrDefault(ParagraphContainsBibliographyField);

    if (bibliographyParagraph != null)
    {
        return GetDirectBodyChild(body, bibliographyParagraph);
    }

    var referencesHeadingParagraph = body
        .Descendants<Paragraph>()
        .LastOrDefault(paragraph =>
            IsHeading1Paragraph(paragraph) &&
            string.Equals(GetNormalizedParagraphText(paragraph), "References", StringComparison.OrdinalIgnoreCase));

    return referencesHeadingParagraph == null
        ? null
        : GetDirectBodyChild(body, referencesHeadingParagraph);
}

// Main recursive processor for markdown nodes
void ProcessMarkdownNode(MarkdownObject node)
{
    switch (node)
    {
        case HeadingBlock headingBlock:
            string? headingTargetId = GetTargetIdFromAttributes(headingBlock)
                ?? ConsumeTrailingTargetMarkerFromInline(headingBlock.Inline, "heading");

            if (!string.IsNullOrEmpty(headingTargetId))
            {
                RegisterCrossReferenceTarget(headingTargetId, "heading");
            }

            docxContent.NewHeading(headingBlock.Level, headingTargetId);
            ProcessInlinesWithUnderline(headingBlock.Inline);
            break;

        case ParagraphBlock paragraphBlock:
            // Check if this paragraph contains only an image link
            var firstInline = paragraphBlock.Inline?.FirstChild;
            if (firstInline is Markdig.Syntax.Inlines.LinkInline { IsImage: true } &&
                paragraphBlock.Inline?.LastChild == firstInline)
            {
                WarnContentOutsideFigureBlock("Image", paragraphBlock);

                // This is a standalone image paragraph - it will be handled by the Figure container
                // Skip processing here as it's handled by the Figure case
                break;
            }

            docxContent.NewParagraph();
            ProcessInlinesWithUnderline(paragraphBlock.Inline);
            break;

        case ListBlock listBlock:
            ProcessListBlock(listBlock);
            break;

        case FencedCodeBlock fencedCodeBlock:
            // Check if this is a references block
            if (string.Equals(fencedCodeBlock.Info?.Trim(), "references", StringComparison.OrdinalIgnoreCase))
            {

                var yamlContent = string.Join("\n", fencedCodeBlock.Lines);
                try
                {
                    var references = ParseReferencesYaml(yamlContent);
                    if (references.Count > 0)
                    {
                        // Remove the heading that preceded this references block
                        docxContent.RemoveLastHeading();
                        docxContent.AddBibliography(references);
                        Console.WriteLine($"[INFO] Parsed {references.Count} bibliography source(s).");
                    }
                    else
                    {
                        Console.WriteLine("[WARNING] References block found but contained no valid entries.");
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"[ERROR] Failed to parse references YAML: {ex.Message}");
                }
                break;
            }

            // Check for fmt: formatted code block
            var fmtInfo = fencedCodeBlock.Info?.Trim() ?? "";
            if (fmtInfo == "fmt" || fmtInfo.StartsWith("fmt:"))
            {
                bool allowBold = false, allowItalic = false, allowUnderline = false, allowStrikethrough = false;

                if (fmtInfo == "fmt")
                {
                    allowBold = allowItalic = allowUnderline = allowStrikethrough = true;
                }
                else
                {
                    var opts = fmtInfo.Substring(4).Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                    foreach (var opt in opts)
                    {
                        switch (opt.ToLowerInvariant())
                        {
                            case "bold": allowBold = true; break;
                            case "italic": allowItalic = true; break;
                            case "underline": allowUnderline = true; break;
                            case "strikethrough": allowStrikethrough = true; break;
                            default:
                                Console.WriteLine($"[WARNING] Unknown fmt option '{opt}'. Valid options: bold, italic, underline, strikethrough.");
                                break;
                        }
                    }
                }

                docxContent.NewParagraph(isLiteral: true);
                var fmtLineGroup = fencedCodeBlock.Lines;
                int fmtLineCount = fmtLineGroup.Count;
                for (int fmtLineIdx = 0; fmtLineIdx < fmtLineCount; fmtLineIdx++)
                {
                    var lineText = fmtLineGroup.Lines[fmtLineIdx].Slice.ToString();

                    // Separate leading whitespace so Markdig doesn't interpret
                    // indented lines as indented code blocks when parsing inline content.
                    var trimmedLine = lineText.TrimStart();
                    var leadingWhitespace = lineText.Substring(0, lineText.Length - trimmedLine.Length);

                    if (leadingWhitespace.Length > 0)
                        docxContent.AddText(leadingWhitespace);

                    var lineDoc = Markdown.Parse(trimmedLine, pipeline);
                    if (lineDoc.FirstOrDefault() is ParagraphBlock fmtPara)
                    {
                        ProcessInlinesWithUnderline(fmtPara.Inline,
                            allowBold: allowBold, allowItalic: allowItalic,
                            allowUnderline: allowUnderline, allowStrikethrough: allowStrikethrough);
                    }
                    else
                    {
                        docxContent.AddText(trimmedLine);
                    }

                    if (fmtLineIdx < fmtLineCount - 1)
                        docxContent.AddLineBreak();
                }
                break;
            }

            var codeContent = string.Join("\n", fencedCodeBlock.Lines);
            docxContent.NewParagraph(isLiteral: true);
            docxContent.AddText(codeContent);
            break;

        case ThematicBreakBlock:
            // Horizontal rule - skip or could add a separator
            break;

        case Markdig.Extensions.CustomContainers.CustomContainer customContainer
            when string.Equals(customContainer.Info?.Trim(), "pagebreak", StringComparison.OrdinalIgnoreCase):
            docxContent.AddPageBreak();
            break;

        case Markdig.Extensions.Figures.Figure figure:
            // Figure can contain Table, ParagraphBlock (with image), and FigureCaption
            string? figureCaption = null;
            int? figureWidth = null;
            string? figureTargetId = null;

            // First, find the caption to get width info
            foreach (var child in figure)
            {
                if (child is Markdig.Extensions.Figures.FigureCaption caption)
                {
                    // FigureCaption is a LeafBlock with Inline content
                    var rawCaption = GetPlainText(caption.Inline);
                    (figureCaption, figureWidth, figureTargetId) = ParseCaptionMetadata(rawCaption, "figure caption");
                    figureTargetId ??= GetTargetIdFromAttributes(caption);
                }
            }

            if (!string.IsNullOrEmpty(figureTargetId))
            {
                RegisterCrossReferenceTarget(figureTargetId, "figure/table caption");
            }

            // Now process the figure contents
            bool hasTable = false;
            bool hasImage = false;
            bool captionTargetApplied = false;

            foreach (var child in figure)
            {
                if (child is Markdig.Extensions.Tables.Table table)
                {
                    hasTable = true;
                    string? bookmarkName = !captionTargetApplied ? figureTargetId : null;
                    ProcessTable(table, figureCaption, figureWidth, bookmarkName);
                    if (!string.IsNullOrEmpty(bookmarkName))
                    {
                        captionTargetApplied = true;
                    }
                }
                else if (child is ParagraphBlock figParagraphBlock)
                {
                    // Check for image inside paragraph
                    foreach (var inline in figParagraphBlock.Inline ?? Enumerable.Empty<Markdig.Syntax.Inlines.Inline>())
                    {
                        if (inline is Markdig.Syntax.Inlines.LinkInline { IsImage: true } figImageLink)
                        {
                            hasImage = true;
                            string? imagePath = figImageLink.Url;
                            if (!string.IsNullOrEmpty(imagePath) && !string.IsNullOrEmpty(inputMDFileDir))
                            {
                                // Resolve relative path
                                imagePath = Path.Combine(inputMDFileDir, imagePath);
                            }

                            // Convert width from pixels to inches (assuming 96 DPI)
                            double? widthInches = figureWidth.HasValue ? figureWidth.Value / 96.0 : null;
                            string? bookmarkName = !captionTargetApplied ? figureTargetId : null;
                            docxContent.AddImage(imagePath ?? string.Empty, figureCaption, widthInches, bookmarkName);
                            if (!string.IsNullOrEmpty(bookmarkName))
                            {
                                captionTargetApplied = true;
                            }
                        }
                    }
                }
                // FigureCaption already processed above
            }

            // Check for missing caption and print warning
            if (string.IsNullOrWhiteSpace(figureCaption))
            {
                if (hasTable)
                {
                    Console.WriteLine("[WARNING] Table is missing a caption.");
                }
                else if (hasImage)
                {
                    Console.WriteLine("[WARNING] Image is missing a caption.");
                }
            }
            break;

        case Markdig.Extensions.Tables.Table table:
            // Standalone table (not in a figure)
            WarnContentOutsideFigureBlock("Table", table);
            ProcessTable(table);
            break;

        case LinkReferenceDefinitionGroup:
            // Skip link reference definitions
            break;

        case MarkdownDocument document:
            foreach (var child in document)
            {
                ProcessMarkdownNode(child);
            }
            break;

        case ContainerBlock containerBlock:
            foreach (var child in containerBlock)
            {
                ProcessMarkdownNode(child);
            }
            break;
    }
}

try
{
    var splitSections = SplitTopLevelDocumentSections(markdownContent);

    if (splitSections.acknowledgementsBlocks.Count > 0 || splitSections.abstractBlocks.Count > 0)
    {
        docxContent = frontMatterDocxContent;
        ProcessFrontMatterSections(splitSections.acknowledgementsBlocks, splitSections.abstractBlocks);
        docxContent = mainDocxContent;
    }

    ProcessTopLevelBlocks(splitSections.remainingBlocks);

    appendixBlocks = splitSections.appendixBlocks;

    if (appendixBlocks.Count > 0)
    {
        docxContent = appendixDocxContent;
        docxContent.AddPageBreak();
        foreach (var appendixBlock in appendixBlocks)
        {
            ProcessMarkdownNode(appendixBlock);
        }

        docxContent = mainDocxContent;
    }

    ValidateCrossReferencesOrThrow();
}
catch (InvalidOperationException ex)
{
    Console.Error.WriteLine($"[ERROR] {ex.Message}");
    return;
}

// Copy DOCX Template to Output File

System.IO.File.Copy(inputDOCXTemplateFilePath, outputDOCXFilePath, true);

// Write to DOCX File

var doc = WordprocessingDocument.Open(outputDOCXFilePath, true);
if (doc.MainDocumentPart is null)
{
    Console.Error.WriteLine("[ERROR] The DOCX template does not contain a valid main document part.");
    return;
}
if (doc.MainDocumentPart.Document is null)
{
    Console.Error.WriteLine("[ERROR] The DOCX template does not contain a valid document.");
    return;
}

var body = doc.MainDocumentPart.Document.Body;
if (body is null)
{
    Console.Error.WriteLine("[ERROR] The DOCX template does not contain a valid body.");
    return;
}

var sectionBreakParagraphs = body
    .Elements<Paragraph>()
    .Where(paragraph => paragraph.Descendants<SectionProperties>().Any())
    .ToList();

if (sectionBreakParagraphs.Count < 2)
{
    Console.Error.WriteLine("[ERROR] The DOCX template must contain at least two section breaks (cover -> front matter and front matter -> main content).");
    return;
}

var firstSectionEndPara = sectionBreakParagraphs[0];
var secondSectionEndPara = sectionBreakParagraphs[1];

// var newPara = new Paragraph(new Run(new Text("Hello, World!")));
// body.InsertAfter(newPara, firstSectionEndPara);

if (frontMatterDocxContent.Elements.Count > 0)
{
    // Insert front matter at the beginning of section 2 (after the cover section break).
    frontMatterDocxContent.WriteTo(doc, firstSectionEndPara);
}

// Insert main content at the beginning of section 3 (after the front-matter section break).
mainDocxContent.WriteTo(doc, secondSectionEndPara);

if (appendixBlocks.Count > 0)
{
    var appendixInsertAnchor = FindAppendixInsertAnchorAfterReferences(body);
    if (appendixInsertAnchor == null)
    {
        Console.WriteLine("[WARNING] Could not find the references anchor in the document. Appending appendix at the end of the body.");
        appendixInsertAnchor = body.LastChild ?? secondSectionEndPara;
    }

    appendixDocxContent.WriteTo(doc, appendixInsertAnchor);
}

doc.Save();