using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using md_to_docx_sync.Enums;
using md_to_docx_sync.Interfaces;
using md_to_docx_sync.Models;
using SixLabors.ImageSharp;

namespace md_to_docx_sync.Services;

public class DocxContent : IDocxContent
{
    // State tracking
    private int _currentHeadingLevel = 1;
    private DocxListType? _currentListType = null;
    private int _listItemIndex = 0;
    private int _localListId = 0; // Local list ID (0-based, will be offset when writing)
    private int _figureSequenceNumber = 0; // Counter for figure numbering
    private int _tableSequenceNumber = 0; // Counter for table numbering
    private readonly Dictionary<string, ReferenceSource> _referencesByKey = new(StringComparer.OrdinalIgnoreCase); // Bibliography sources keyed by citation tag

    // Current elements being built
    private Paragraph? _currentParagraph = null;
    private Table? _currentTable = null;
    private TableRow? _currentTableRow = null;
    private TableCell? _currentTableCell = null;
    private int _currentTableRowIndex = 0;

    // Collected content
    private readonly List<OpenXmlElement> _elements = new();

    // List definitions (stored with local IDs and heading level, will be remapped when writing)
    private readonly List<(int localId, DocxListType type, int headingLevel)> _listDefinitions = new();

    /// <summary>
    /// Gets all the collected OpenXML elements to be inserted into the document body.
    /// </summary>
    public IReadOnlyList<OpenXmlElement> Elements => _elements;

    /// <summary>
    /// Gets the style name for body text based on current heading level.
    /// </summary>
    private string BodyStyleId => $"Body{_currentHeadingLevel}";

    /// <summary>
    /// Gets the style name for list paragraphs based on current heading level.
    /// </summary>
    private string ListStyleId => $"LP{_currentHeadingLevel}";

    /// <summary>
    /// Gets the style name for code blocks based on current heading level.
    /// </summary>
    private string CodeStyleId => $"Code{_currentHeadingLevel}";

    /// <summary>
    /// Gets the heading style name based on level.
    /// </summary>
    private static string GetHeadingStyleId(int level) => $"Heading{level}";

    public void NewHeading(int level)
    {
        FinalizeCurrentContext();

        _currentHeadingLevel = Math.Clamp(level, 1, 4);

        _currentParagraph = new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId() { Val = GetHeadingStyleId(_currentHeadingLevel) }
            )
        );
        _elements.Add(_currentParagraph);
    }

    public void NewParagraph(bool isLiteral = false)
    {
        FinalizeCurrentContext();

        string styleId = isLiteral ? CodeStyleId : BodyStyleId;

        _currentParagraph = new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId() { Val = styleId }
            )
        );
        _elements.Add(_currentParagraph);
    }

    private static Paragraph CreateTableSpacerParagraph()
    {
        return new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId() { Val = "Normal" },
                new SpacingBetweenLines()
                {
                    After = "0",
                    Line = "360",
                    LineRule = LineSpacingRuleValues.Auto
                }
            )
        );
    }

    public void AddText(string text, bool isBold = false, bool isItalic = false, bool isUnderline = false, bool isStrikethrough = false, bool isLiteral = false)
    {
        EnsureCurrentParagraph();

        var runProperties = new RunProperties();

        if (isBold)
        {
            runProperties.Append(new Bold());
        }

        if (isItalic)
        {
            runProperties.Append(new Italic());
        }

        if (isUnderline)
        {
            runProperties.Append(new Underline() { Val = UnderlineValues.Single });
        }

        if (isStrikethrough)
        {
            runProperties.Append(new Strike());
        }

        // // For inline code (literal), apply monospace font styling
        // // but keep the same paragraph style (Body1, Body2, etc.)
        // if (isLiteral)
        // {
        //     runProperties.Append(new RunFonts() { Ascii = "Courier New", HighAnsi = "Courier New", ComplexScript = "Courier New" });
        //     runProperties.Append(new Shading() { Val = ShadingPatternValues.Clear, Fill = "E0E0E0" });
        // }

        // Handle text with newlines (for code blocks)
        var lines = text.Split('\n');
        for (int i = 0; i < lines.Length; i++)
        {
            var run = new Run();

            if (runProperties.HasChildren)
            {
                run.Append(runProperties.CloneNode(true));
            }

            run.Append(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve });
            _currentParagraph!.Append(run);

            // Add line break between lines (not after the last one)
            if (i < lines.Length - 1)
            {
                var breakRun = new Run(new Break());
                _currentParagraph.Append(breakRun);
            }
        }
    }

    public void AddLineBreak()
    {
        EnsureCurrentParagraph();
        _currentParagraph!.Append(new Run(new Break()));
    }

    public void AddPageBreak()
    {
        FinalizeCurrentContext();

        var paragraph = new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId() { Val = BodyStyleId }
            ),
            new Run(new Break() { Type = BreakValues.Page })
        );

        _elements.Add(paragraph);
    }

    public void NewList(DocxListType type)
    {
        FinalizeCurrentContext();

        _currentListType = type;
        _listItemIndex = 0;

        // Store list definition with local ID and current heading level (will be remapped when writing to document)
        _listDefinitions.Add((_localListId, type, _currentHeadingLevel));
        _localListId++;
    }

    public void NewListItem(int level = 0)
    {
        if (_currentListType == null)
        {
            throw new InvalidOperationException("Cannot create a list item without first calling NewList.");
        }

        // Clamp level to valid range (0-3)
        level = Math.Clamp(level, 0, 3);

        // Finalize any previous paragraph but keep list context
        _currentParagraph = null;

        // Use local list ID (will be remapped when writing to document)
        var localNumberingId = _localListId - 1;

        // Apply style first, then numbering properties
        var paragraphProperties = new ParagraphProperties();

        // 1. First apply the list style (LP1, LP2, LP3, LP4)
        paragraphProperties.Append(new ParagraphStyleId() { Val = ListStyleId });

        // 2. Then apply the numbering properties with local ID (will be remapped)
        // We use a placeholder value that will be updated in WriteTo
        paragraphProperties.Append(new NumberingProperties(
            new NumberingLevelReference() { Val = level },
            new NumberingId() { Val = localNumberingId }
        ));

        _currentParagraph = new Paragraph(paragraphProperties);
        _elements.Add(_currentParagraph);
        _listItemIndex++;
    }

    public void NewTable(string? caption = null, int? widthPercent = null)
    {
        FinalizeCurrentContext();

        // Add caption if provided
        if (!string.IsNullOrEmpty(caption))
        {
            // Increment table sequence number
            _tableSequenceNumber++;

            // Create a proper caption with SEQ field for automatic table numbering
            var captionParagraph = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Caption" },
                    new Justification() { Val = JustificationValues.Center }
                )
            );

            // Add "Table " text
            var tableRun = new Run(new Text("Table ") { Space = SpaceProcessingModeValues.Preserve });
            captionParagraph.Append(tableRun);

            // Add SEQ field for automatic numbering - properly structured
            // Field begin
            var fieldBeginRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
            captionParagraph.Append(fieldBeginRun);

            // Field code
            var fieldCodeRun = new Run(new FieldCode(" SEQ Table \\* ARABIC "));
            captionParagraph.Append(fieldCodeRun);

            // Field separator (required between code and result)
            var fieldSeparateRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate });
            captionParagraph.Append(fieldSeparateRun);

            // Field result placeholder - use the actual sequence number so TOC fields display correctly
            var fieldResultRun = new Run(new Text(_tableSequenceNumber.ToString()));
            captionParagraph.Append(fieldResultRun);

            // Field end
            var fieldEndRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });
            captionParagraph.Append(fieldEndRun);

            // Add ": " separator and caption text
            var separatorRun = new Run(new Text(": ") { Space = SpaceProcessingModeValues.Preserve });
            captionParagraph.Append(separatorRun);

            var captionRun = new Run(new Text(caption));
            captionParagraph.Append(captionRun);

            _elements.Add(captionParagraph);
        }

        // Calculate table width (default to 100% if not specified)
        // In OpenXML, table width in Pct is specified as 50ths of a percent (so 5000 = 100%)
        int effectiveWidthPercent = Math.Clamp(widthPercent ?? 100, 1, 100);
        string tableWidthValue = (effectiveWidthPercent * 50).ToString();

        // Create table with styling
        var tableProperties = new TableProperties(
            // Black borders
            new TableBorders(
                new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 },
                new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 },
                new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 },
                new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 },
                new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 },
                new InsideVerticalBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 }
            ),
            // Width based on parameter (default 100%)
            new TableWidth() { Width = tableWidthValue, Type = TableWidthUnitValues.Pct },
            // Center alignment
            new TableJustification() { Val = TableRowAlignmentValues.Center }
        );

        _currentTable = new Table(tableProperties);
        _elements.Add(_currentTable);
        // Add a spacer paragraph to keep separation between the table and following block content.
        _elements.Add(CreateTableSpacerParagraph());
        _currentTableRowIndex = 0;
    }

    public void NewRow()
    {
        if (_currentTable == null)
        {
            throw new InvalidOperationException("Cannot create a row without first calling NewTable.");
        }

        _currentTableRow = new TableRow();
        _currentTable.Append(_currentTableRow);
        _currentTableCell = null;
        _currentParagraph = null;
        _currentTableRowIndex++;
    }

    public void NewCell(int columnSpan = 1, bool verticalMergeStart = false, bool verticalMergeContinue = false, JustificationValues? alignment = null)
    {
        if (_currentTableRow == null)
        {
            throw new InvalidOperationException("Cannot create a cell without first calling NewRow.");
        }

        var cellProperties = new TableCellProperties(
            new TableCellBorders(
                new TopBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 },
                new BottomBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 },
                new LeftBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 },
                new RightBorder() { Val = BorderValues.Single, Color = "000000", Size = 4 }
            )
        );

        columnSpan = Math.Max(1, columnSpan);
        if (columnSpan > 1)
        {
            cellProperties.Append(new GridSpan() { Val = columnSpan });
        }

        if (verticalMergeStart)
        {
            cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Restart });
        }
        else if (verticalMergeContinue)
        {
            cellProperties.Append(new VerticalMerge() { Val = MergedCellValues.Continue });
        }

        _currentTableCell = new TableCell(cellProperties);
        _currentTableRow.Append(_currentTableCell);

        // Create a paragraph inside the cell for text content using 'Normal' style.
        var paragraphProps = new ParagraphProperties(
            new ParagraphStyleId() { Val = "Normal" }
        );

        if (alignment.HasValue)
        {
            paragraphProps.Append(new Justification() { Val = alignment.Value });
        }

        _currentParagraph = new Paragraph(paragraphProps);
        _currentTableCell.Append(_currentParagraph);
    }

    public void AddImage(string imagePath, string? caption = null, double? widthInches = null)
    {
        FinalizeCurrentContext();

        if (!System.IO.File.Exists(imagePath))
        {
            throw new FileNotFoundException($"Image file not found: {imagePath}");
        }

        // Store the image information to be added when WriteTo is called
        // We'll create a placeholder paragraph that will be replaced with the actual image
        var imageParagraph = new Paragraph(
            new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center }
            )
        );

        // Add a marker run to identify this as an image placeholder
        // Include width info in the marker (use -1 for auto/original size)
        string widthMarker = widthInches.HasValue ? widthInches.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "-1";
        var markerRun = new Run(new Text($"[IMAGE:{imagePath}|WIDTH:{widthMarker}]") { Space = SpaceProcessingModeValues.Preserve });
        imageParagraph.Append(markerRun);
        _elements.Add(imageParagraph);

        // Add caption if provided
        if (!string.IsNullOrEmpty(caption))
        {
            // Increment figure sequence number
            _figureSequenceNumber++;

            // Create a proper caption with SEQ field for automatic figure numbering
            var captionParagraph = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId() { Val = "Caption" },
                    new Justification() { Val = JustificationValues.Center }
                )
            );

            // Add "Figure " text
            var figureRun = new Run(new Text("Figure ") { Space = SpaceProcessingModeValues.Preserve });
            captionParagraph.Append(figureRun);

            // Add SEQ field for automatic numbering
            var fieldBeginRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
            captionParagraph.Append(fieldBeginRun);

            var fieldCodeRun = new Run(new FieldCode(" SEQ Figure \\* ARABIC "));
            captionParagraph.Append(fieldCodeRun);

            var fieldSeparateRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate });
            captionParagraph.Append(fieldSeparateRun);

            // Field result placeholder - use the actual sequence number so TOC fields display correctly
            var fieldResultRun = new Run(new Text(_figureSequenceNumber.ToString()));
            captionParagraph.Append(fieldResultRun);

            // Field end
            var fieldEndRun = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });
            captionParagraph.Append(fieldEndRun);

            // Add ": " separator and caption text
            var separatorRun = new Run(new Text(": ") { Space = SpaceProcessingModeValues.Preserve });
            captionParagraph.Append(separatorRun);

            var captionRun = new Run(new Text(caption));
            captionParagraph.Append(captionRun);

            _elements.Add(captionParagraph);
        }
    }

    public void AddCitation(params string[] keys)
    {
        if (keys.Length == 0) return;
        EnsureCurrentParagraph();

        for (int i = 0; i < keys.Length; i++)
        {
            // Build field code:
            //   First citation:      CITATION key \l 1033
            //   Subsequent (merged):  CITATION key \l 1033  \m firstKey
            string fieldCode = i == 0
                ? $" CITATION {keys[i]} \\l 1033 "
                : $" CITATION {keys[i]} \\l 1033  \\m {keys[0]}";

            // CITATION field begin
            _currentParagraph!.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }));

            // CITATION field code
            _currentParagraph.Append(new Run(
                new FieldCode(fieldCode) { Space = SpaceProcessingModeValues.Preserve }));

            // CITATION field separator
            _currentParagraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }));

            // Placeholder result text with NoProof (Word will replace this when fields are updated)
            var resultRun = new Run();
            resultRun.Append(new RunProperties(new NoProof()));
            resultRun.Append(new Text($"[{keys[i]}]") { Space = SpaceProcessingModeValues.Preserve });
            _currentParagraph.Append(resultRun);

            // CITATION field end
            _currentParagraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.End }));
        }
    }

    public void AddBibliography(List<ReferenceSource> references)
    {
        int addedCount = 0;
        int duplicateCount = 0;
        int conflictingCount = 0;

        foreach (var reference in references)
        {
            if (_referencesByKey.TryGetValue(reference.Key, out var existingReference))
            {
                if (HasSameReferenceDefinition(existingReference, reference))
                {
                    duplicateCount++;
                }
                else
                {
                    conflictingCount++;
                    Console.WriteLine($"[WARNING] Duplicate reference key '{reference.Key}' has conflicting definitions. Keeping the first definition.");
                }

                continue;
            }

            _referencesByKey[reference.Key] = CloneReference(reference);
            addedCount++;
        }

        Console.WriteLine($"[INFO] Accepted {addedCount} unique bibliography source(s), skipped {duplicateCount} duplicate(s), found {conflictingCount} conflicting duplicate(s).");
    }

    private static ReferenceSource CloneReference(ReferenceSource reference)
    {
        return new ReferenceSource
        {
            Key = reference.Key,
            Type = reference.Type,
            Fields = new Dictionary<string, string>(reference.Fields)
        };
    }

    private static bool HasSameReferenceDefinition(ReferenceSource left, ReferenceSource right)
    {
        if (!string.Equals(left.Type, right.Type, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (left.Fields.Count != right.Fields.Count)
        {
            return false;
        }

        var rightFields = new Dictionary<string, string>(right.Fields, StringComparer.OrdinalIgnoreCase);
        foreach (var (fieldName, fieldValue) in left.Fields)
        {
            if (!rightFields.TryGetValue(fieldName, out var rightValue))
            {
                return false;
            }

            if (!string.Equals(fieldValue, rightValue, StringComparison.Ordinal))
            {
                return false;
            }
        }

        return true;
    }

    public void RemoveLastHeading()
    {
        // Walk backwards to find the last paragraph with a heading style
        for (int i = _elements.Count - 1; i >= 0; i--)
        {
            if (_elements[i] is Paragraph paragraph)
            {
                var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
                if (styleId != null && styleId.StartsWith("Heading"))
                {
                    _elements.RemoveAt(i);
                    return;
                }
            }
        }
    }

    /// <summary>
    /// Writes the collected content to a WordprocessingDocument.
    /// </summary>
    public void WriteTo(WordprocessingDocument document, OpenXmlElement insertAfterElement)
    {
        var body = document.MainDocumentPart?.Document?.Body;
        if (body == null)
        {
            throw new InvalidOperationException("Document does not have a valid body.");
        }

        // Enable automatic field updates when the document is opened
        // This ensures SEQ fields for figures and tables are updated correctly
        EnableUpdateFieldsOnOpen(document);

        // Mark TOC fields as dirty to force Word to regenerate the Table of Contents
        MarkTocFieldsAsDirty(document);

        // Write bibliography sources to CustomXmlPart if any references were provided
        WriteBibliographySources(document);

        // Calculate ID offset based on existing numbering definitions
        int idOffset = 0;
        if (_listDefinitions.Count > 0)
        {
            idOffset = AddNumberingDefinitions(document);
        }

        // Clone elements and remap numbering IDs / process image placeholders
        var previousElement = insertAfterElement;
        foreach (var element in _elements)
        {
            var clonedElement = element.CloneNode(true);

            // Check if this is an image placeholder paragraph
            if (clonedElement is Paragraph paragraph)
            {
                RemapNumberingId(paragraph, idOffset);

                // Check for image placeholder marker
                var textRun = paragraph.Descendants<Text>().FirstOrDefault();
                if (textRun?.Text != null && textRun.Text.StartsWith("[IMAGE:"))
                {
                    // Parse marker: [IMAGE:path|WIDTH:value]
                    var markerContent = textRun.Text.Substring(7, textRun.Text.Length - 8); // Remove "[IMAGE:" and "]"
                    var parts = markerContent.Split("|WIDTH:");
                    var imagePath = parts[0];
                    double? widthInches = null;

                    if (parts.Length > 1 && double.TryParse(parts[1], System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double parsedWidth) && parsedWidth > 0)
                    {
                        widthInches = parsedWidth;
                    }

                    // Replace paragraph with actual image
                    clonedElement = CreateImageParagraph(document, imagePath, widthInches);
                }
            }

            body.InsertAfter(clonedElement, previousElement);
            previousElement = previousElement.NextSibling() ?? previousElement;
        }
    }

    /// <summary>
    /// Enables the UpdateFieldsOnOpen setting in the document.
    /// This ensures SEQ fields (for figure and table numbering) are updated when the document is opened.
    /// </summary>
    private static void EnableUpdateFieldsOnOpen(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return;

        // Get or create the DocumentSettingsPart
        var settingsPart = mainPart.DocumentSettingsPart;
        if (settingsPart == null)
        {
            settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();
        }

        var settings = settingsPart.Settings;
        if (settings == null)
        {
            settings = new Settings();
            settingsPart.Settings = settings;
        }

        // Check if UpdateFieldsOnOpen already exists
        var updateFields = settings.Elements<UpdateFieldsOnOpen>().FirstOrDefault();
        if (updateFields == null)
        {
            // Add UpdateFieldsOnOpen element
            settings.PrependChild(new UpdateFieldsOnOpen() { Val = true });
        }
        else
        {
            // Ensure it's set to true
            updateFields.Val = true;
        }

        settings.Save();

        Console.WriteLine("[NOTE] You will get a message to update the fields when opening the document. Please accept this to ensure figure and table numbers are correct.");
    }

    /// <summary>
    /// Marks all TOC (Table of Contents) fields in the document as dirty,
    /// forcing Word to regenerate them when the document is opened.
    /// This is necessary because UpdateFieldsOnOpen alone may not reliably update TOC fields,
    /// even though it works fine for Table of Figures and Table of Tables.
    /// </summary>
    private static void MarkTocFieldsAsDirty(WordprocessingDocument document)
    {
        var body = document.MainDocumentPart?.Document?.Body;
        if (body == null) return;

        // Handle complex fields (fldChar + instrText pattern)
        // Track the most recent Begin FieldChar so we can mark it dirty when we find a TOC instruction
        var fieldCharBeginStack = new Stack<FieldChar>();

        foreach (var run in body.Descendants<Run>())
        {
            var fieldChar = run.GetFirstChild<FieldChar>();
            if (fieldChar != null)
            {
                if (fieldChar.FieldCharType?.Value == FieldCharValues.Begin)
                {
                    fieldCharBeginStack.Push(fieldChar);
                }
                else if (fieldChar.FieldCharType?.Value == FieldCharValues.End && fieldCharBeginStack.Count > 0)
                {
                    fieldCharBeginStack.Pop();
                }
            }

            var fieldCode = run.GetFirstChild<FieldCode>();
            if (fieldCode != null && fieldCode.Text?.TrimStart().StartsWith("TOC") == true)
            {
                // Mark the associated Begin FieldChar as dirty
                if (fieldCharBeginStack.Count > 0)
                {
                    fieldCharBeginStack.Peek().Dirty = true;
                }
            }
        }

        // Handle simple fields (<w:fldSimple>)
        foreach (var simpleField in body.Descendants<SimpleField>())
        {
            if (simpleField.Instruction?.Value?.TrimStart().StartsWith("TOC") == true)
            {
                simpleField.Dirty = true;
            }
        }
    }

    /// <summary>
    /// Writes bibliography sources to a CustomXmlPart using Word's bibliography XML schema.
    /// If the template already contains a bibliography CustomXmlPart, the sources are merged into it.
    /// Otherwise a new CustomXmlPart is created.
    /// This allows Word to recognize the sources in "Manage Sources" and render CITATION / BIBLIOGRAPHY fields.
    /// </summary>
    private void WriteBibliographySources(WordprocessingDocument document)
    {
        if (_referencesByKey.Count == 0) return;

        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return;

        XNamespace bNs = "http://schemas.openxmlformats.org/officeDocument/2006/bibliography";

        // Try to find an existing bibliography CustomXmlPart in the template
        CustomXmlPart? existingBibPart = null;
        XDocument? existingDoc = null;

        foreach (var customXmlPart in mainPart.CustomXmlParts)
        {
            try
            {
                using var stream = customXmlPart.GetStream(FileMode.Open);
                var doc = XDocument.Load(stream);
                if (doc.Root?.Name.Namespace == bNs)
                {
                    existingBibPart = customXmlPart;
                    existingDoc = doc;
                    break;
                }
            }
            catch
            {
                // Not a valid XML part or not bibliography data — skip
            }
        }

        if (existingBibPart != null && existingDoc?.Root != null)
        {
            // Merge new sources into the existing bibliography part
            var existingTags = existingDoc.Root
                .Elements(bNs + "Source")
                .Select(source => source.Element(bNs + "Tag")?.Value)
                .Where(tag => !string.IsNullOrWhiteSpace(tag))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            int mergedCount = 0;
            int skippedExistingCount = 0;

            foreach (var reference in _referencesByKey.Values)
            {
                if (existingTags.Contains(reference.Key))
                {
                    skippedExistingCount++;
                    continue;
                }

                existingDoc.Root.Add(CreateSourceXElement(reference, bNs));
                existingTags.Add(reference.Key);
                mergedCount++;
            }

            using var writeStream = existingBibPart.GetStream(FileMode.Create);
            existingDoc.Save(writeStream);

            Console.WriteLine($"[INFO] Merged {mergedCount} bibliography source(s) into existing bibliography data; skipped {skippedExistingCount} source(s) already present in template.");
        }
        else
        {
            // Create a brand-new bibliography CustomXmlPart
            var sourcesElement = new XElement(bNs + "Sources",
                new XAttribute(XNamespace.Xmlns + "b", bNs.NamespaceName),
                new XAttribute("SelectedStyle", "\\APASixthEditionOfficeOnline.xsl"),
                new XAttribute("StyleName", "APA"),
                new XAttribute("Version", "6"),
                _referencesByKey.Values.Select(r => CreateSourceXElement(r, bNs))
            );

            var xmlDoc = new XDocument(
                new XDeclaration("1.0", "UTF-8", "yes"),
                sourcesElement
            );

            var customXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            using (var stream = customXmlPart.GetStream(FileMode.Create))
            {
                xmlDoc.Save(stream);
            }

            // Add CustomXmlPropertiesPart using SDK types for correct XML serialization
            var propsPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
            propsPart.DataStoreItem = new DataStoreItem
            {
                ItemId = $"{{{Guid.NewGuid()}}}"
            };
            propsPart.DataStoreItem.Append(
                new SchemaReferences(
                    new SchemaReference { Uri = bNs.NamespaceName }
                )
            );
            propsPart.DataStoreItem.Save();

            Console.WriteLine($"[INFO] Created new bibliography data with {_referencesByKey.Count} unique source(s).");
        }
    }

    /// <summary>
    /// Creates an XElement representing a single bibliography source in Word's XML format.
    /// </summary>
    private static XElement CreateSourceXElement(ReferenceSource reference, XNamespace bNs)
    {
        var sourceElement = new XElement(bNs + "Source");

        // Tag (citation key)
        sourceElement.Add(new XElement(bNs + "Tag", reference.Key));

        // SourceType (mapped from user-friendly name to Word's internal name)
        string wordSourceType = MapSourceType(reference.Type);
        sourceElement.Add(new XElement(bNs + "SourceType", wordSourceType));

        // GUID
        sourceElement.Add(new XElement(bNs + "Guid", $"{{{Guid.NewGuid()}}}"));

        // LCID (locale identifier — 0 = default)
        sourceElement.Add(new XElement(bNs + "LCID", "0"));

        // Author container (holds all person-type fields)
        var authorContainer = new XElement(bNs + "Author");
        bool hasAuthors = false;

        // Process all fields
        foreach (var (fieldName, fieldValue) in reference.Fields)
        {
            if (string.IsNullOrWhiteSpace(fieldValue)) continue;

            // Check if this is a person/name field
            string? authorRole = GetAuthorRole(fieldName);
            if (authorRole != null)
            {
                var nameListElement = ParseNameList(fieldValue, bNs);
                authorContainer.Add(new XElement(bNs + authorRole, nameListElement));
                hasAuthors = true;
                continue;
            }

            // Check if this is a simple field
            string? xmlElementName = MapFieldToXmlElement(fieldName);
            if (xmlElementName != null)
            {
                sourceElement.Add(new XElement(bNs + xmlElementName, fieldValue));
            }
        }

        if (hasAuthors)
        {
            sourceElement.Add(authorContainer);
        }

        return sourceElement;
    }

    /// <summary>
    /// Maps user-friendly source type names to Word's internal SourceType values.
    /// </summary>
    private static string MapSourceType(string type)
    {
        return type switch
        {
            "Book" => "Book",
            "BookSection" => "BookSection",
            "JournalArticle" => "JournalArticle",
            "ArticleInAPeriodical" => "ArticleInAPeriodical",
            "ConferenceProceedings" => "ConferenceProceedings",
            "Report" => "Report",
            "WebSite" => "InternetSite",
            "DocumentFromWebSite" => "DocumentFromInternetSite",
            "ElectronicSource" => "ElectronicSource",
            "Art" => "Art",
            "SoundRecording" => "SoundRecording",
            "Performance" => "Performance",
            "Film" => "Film",
            "Interview" => "Interview",
            "Patent" => "Patent",
            "Case" => "Case",
            "Misc" => "Misc",
            _ => "Misc"
        };
    }

    /// <summary>
    /// Returns the Word XML author role element name for person-type YAML fields, or null if not a person field.
    /// </summary>
    private static string? GetAuthorRole(string fieldName)
    {
        return fieldName switch
        {
            "author" => "Author",
            "editor" => "Editor",
            "artist" => "Artist",
            "composer" => "Composer",
            "director" => "Director",
            "writer" => "Writer",
            "interviewee" => "Interviewee",
            "inventor" => "Inventor",
            _ => null
        };
    }

    /// <summary>
    /// Maps YAML field names to Word bibliography XML element names.
    /// Returns null for unrecognized fields.
    /// </summary>
    private static string? MapFieldToXmlElement(string fieldName)
    {
        return fieldName switch
        {
            "title" => "Title",
            "year" => "Year",
            "city" => "City",
            "publisher" => "Publisher",
            "edition" => "Edition",
            "pages" => "Pages",
            "volume" => "Volume",
            "issue" => "Issue",
            "month" => "Month",
            "day" => "Day",
            "yearAccessed" => "YearAccessed",
            "monthAccessed" => "MonthAccessed",
            "dayAccessed" => "DayAccessed",
            "url" => "URL",
            "bookTitle" => "BookTitle",
            "journalName" => "JournalName",
            "periodicalTitle" => "PeriodicalTitle",
            "nameOfWebPage" => "InternetSiteTitle",
            "institution" => "Institution",
            "productionCompany" => "ProductionCompany",
            "countryRegion" => "CountryRegion",
            "patentNumber" => "PatentNumber",
            "reporter" => "Reporter",
            _ => null
        };
    }

    /// <summary>
    /// Parses a name string (e.g., "Kramer, James D; Chen, Jacky") into a Word NameList XElement.
    /// Names are semicolon-separated, each in "Last, First" format.
    /// </summary>
    private static XElement ParseNameList(string nameString, XNamespace bNs)
    {
        var nameList = new XElement(bNs + "NameList");

        var people = nameString.Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        foreach (var person in people)
        {
            var personElement = new XElement(bNs + "Person");
            var parts = person.Split(',', 2, StringSplitOptions.TrimEntries);

            if (parts.Length >= 1)
            {
                personElement.Add(new XElement(bNs + "Last", parts[0]));
            }
            if (parts.Length >= 2)
            {
                personElement.Add(new XElement(bNs + "First", parts[1]));
            }

            nameList.Add(personElement);
        }

        return nameList;
    }

    /// <summary>
    /// Remaps local numbering IDs to actual document IDs.
    /// </summary>
    private static void RemapNumberingId(Paragraph paragraph, int idOffset)
    {
        var numId = paragraph.ParagraphProperties?.NumberingProperties?.NumberingId;
        if (numId?.Val != null)
        {
            numId.Val = numId.Val + idOffset;
        }
    }

    /// <summary>
    /// Creates a paragraph containing an image with border and center alignment.
    /// </summary>
    /// <param name="document">The Word document.</param>
    /// <param name="imagePath">Path to the image file.</param>
    /// <param name="widthInches">Optional width in inches. If null, uses original image dimensions.</param>
    private static Paragraph CreateImageParagraph(WordprocessingDocument document, string imagePath, double? widthInches = null)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null)
        {
            throw new InvalidOperationException("Document does not have a valid main part.");
        }

        // Get original image dimensions (cross-platform)
        var (originalWidthPx, originalHeightPx) = GetImageDimensions(imagePath);

        // Add the image to the document
        ImagePart imagePart;
        using (var stream = new FileStream(imagePath, FileMode.Open))
        {
            imagePart = mainPart.AddImagePart(GetImagePartType(imagePath));
            imagePart.FeedData(stream);
        }

        // Calculate dimensions in EMUs (914400 EMUs = 1 inch)
        const long emusPerInch = 914400;
        long imageWidthEmu, imageHeightEmu;

        if (widthInches.HasValue)
        {
            // Use specified width and calculate height to maintain aspect ratio
            imageWidthEmu = (long)(widthInches.Value * emusPerInch);
            double aspectRatio = (double)originalHeightPx / originalWidthPx;
            imageHeightEmu = (long)(imageWidthEmu * aspectRatio);
        }
        else
        {
            // Use original image dimensions (assuming 96 DPI for screen images)
            const double dpi = 96.0;
            imageWidthEmu = (long)(originalWidthPx / dpi * emusPerInch);
            imageHeightEmu = (long)(originalHeightPx / dpi * emusPerInch);
        }

        var relationshipId = mainPart.GetIdOfPart(imagePart);

        // Generate unique IDs for the drawing
        var drawingId = GenerateUniqueId();
        var imageId = GenerateUniqueId();

        const int borderWidthEmu = 12700;
        const long borderEffectExtentEmu = borderWidthEmu / 2L;

        // Create the drawing element with border
        var element = new DocumentFormat.OpenXml.Wordprocessing.Drawing(
            new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = imageWidthEmu, Cy = imageHeightEmu },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                {
                    LeftEdge = borderEffectExtentEmu,
                    TopEdge = borderEffectExtentEmu,
                    RightEdge = borderEffectExtentEmu,
                    BottomEdge = borderEffectExtentEmu
                },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties() { Id = (uint)drawingId, Name = $"Picture {drawingId}" },
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                    new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }
                ),
                new DocumentFormat.OpenXml.Drawing.Graphic(
                    new DocumentFormat.OpenXml.Drawing.GraphicData(
                        new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                            new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties() { Id = (uint)imageId, Name = System.IO.Path.GetFileName(imagePath) },
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                            ),
                            new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                new DocumentFormat.OpenXml.Drawing.Blip() { Embed = relationshipId },
                                new DocumentFormat.OpenXml.Drawing.Stretch(
                                    new DocumentFormat.OpenXml.Drawing.FillRectangle()
                                )
                            ),
                            new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                new DocumentFormat.OpenXml.Drawing.Transform2D(
                                    new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                    new DocumentFormat.OpenXml.Drawing.Extents() { Cx = imageWidthEmu, Cy = imageHeightEmu }
                                ),
                                new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                    new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                )
                                { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle },
                                // Add black border (1 pt = 12700 EMUs)
                                new DocumentFormat.OpenXml.Drawing.Outline(
                                    new DocumentFormat.OpenXml.Drawing.SolidFill(
                                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex() { Val = "000000" }
                                    )
                                )
                                { Width = borderWidthEmu }
                            )
                        )
                    )
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                )
            )
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            }
        );

        // Create centered paragraph with the image
        var paragraph = new Paragraph(
            new ParagraphProperties(
                new Justification() { Val = JustificationValues.Center }
            ),
            new Run(element)
        );

        return paragraph;
    }

    /// <summary>
    /// Determines the ImagePartType based on file extension.
    /// </summary>
    private static PartTypeInfo GetImagePartType(string imagePath)
    {
        var extension = System.IO.Path.GetExtension(imagePath).ToLower();
        return extension switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tif" or ".tiff" => ImagePartType.Tiff,
            _ => ImagePartType.Png
        };
    }

    /// <summary>
    /// Gets image dimensions using ImageSharp (cross-platform).
    /// </summary>
    private static (int width, int height) GetImageDimensions(string imagePath)
    {
        using var image = Image.Load(imagePath);
        return (image.Width, image.Height);
    }

    /// <summary>
    /// Generates a unique ID for drawing elements.
    /// </summary>
    private static int GenerateUniqueId()
    {
        return (int)(DateTime.Now.Ticks & 0x7FFFFFFF);
    }

    /// <summary>
    /// Adds numbering definitions to the document and returns the ID offset used.
    /// </summary>
    private int AddNumberingDefinitions(WordprocessingDocument document)
    {
        var mainPart = document.MainDocumentPart;
        if (mainPart == null) return 0;

        NumberingDefinitionsPart numberingPart;
        if (mainPart.NumberingDefinitionsPart == null)
        {
            numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering = new Numbering();
        }
        else
        {
            numberingPart = mainPart.NumberingDefinitionsPart;
            if (numberingPart.Numbering == null)
            {
                numberingPart.Numbering = new Numbering();
            }
        }

        // Find the highest existing AbstractNumId and NumberID
        int maxAbstractNumId = 0;
        int maxNumberId = 0;

        foreach (var existingAbstractNum in numberingPart.Numbering.Elements<AbstractNum>())
        {
            if (existingAbstractNum.AbstractNumberId?.Value > maxAbstractNumId)
            {
                maxAbstractNumId = existingAbstractNum.AbstractNumberId.Value;
            }
        }

        foreach (var existingNumInstance in numberingPart.Numbering.Elements<NumberingInstance>())
        {
            if (existingNumInstance.NumberID?.Value > maxNumberId)
            {
                maxNumberId = existingNumInstance.NumberID.Value;
            }
        }

        // Use the higher of the two as our starting offset
        int idOffset = Math.Max(maxAbstractNumId, maxNumberId) + 1;

        // In OpenXML, all AbstractNum elements MUST come before all NumberingInstance elements
        // Find the last AbstractNum to insert after, or first NumberingInstance to insert before
        var lastAbstractNum = numberingPart.Numbering.Elements<AbstractNum>().LastOrDefault();
        var firstNumInstance = numberingPart.Numbering.Elements<NumberingInstance>().FirstOrDefault();

        // Create all AbstractNum elements first
        var newAbstractNums = new List<AbstractNum>();
        var newNumberingInstances = new List<NumberingInstance>();

        foreach (var (localId, listType, headingLevel) in _listDefinitions)
        {
            int actualId = localId + idOffset;

            var abstractNum = CreateAbstractNum(document, actualId, listType, headingLevel);
            newAbstractNums.Add(abstractNum);

            var numberingInstance = new NumberingInstance(
                new AbstractNumId() { Val = actualId }
            )
            { NumberID = actualId };
            newNumberingInstances.Add(numberingInstance);
        }

        // Insert AbstractNum elements after the last existing AbstractNum (but before NumberingInstances)
        if (lastAbstractNum != null)
        {
            var insertAfter = lastAbstractNum;
            foreach (var abstractNum in newAbstractNums)
            {
                numberingPart.Numbering.InsertAfter(abstractNum, insertAfter);
                insertAfter = abstractNum;
            }
        }
        else if (firstNumInstance != null)
        {
            // No existing AbstractNum, insert before first NumberingInstance
            foreach (var abstractNum in newAbstractNums.AsEnumerable().Reverse())
            {
                numberingPart.Numbering.InsertBefore(abstractNum, firstNumInstance);
            }
        }
        else
        {
            // No existing elements, just append
            foreach (var abstractNum in newAbstractNums)
            {
                numberingPart.Numbering.Append(abstractNum);
            }
        }

        // Append NumberingInstance elements at the end
        foreach (var numInstance in newNumberingInstances)
        {
            numberingPart.Numbering.Append(numInstance);
        }

        numberingPart.Numbering.Save();

        return idOffset;
    }

    /// <summary>
    /// Creates an abstract numbering definition for bullet or numbered lists.
    /// Defines all 9 levels (0-8) for proper multi-level list support.
    /// </summary>
    private static AbstractNum CreateAbstractNum(WordprocessingDocument document, int abstractNumId, DocxListType listType, int headingLevel)
    {
        var abstractNum = new AbstractNum() { AbstractNumberId = abstractNumId };

        // Must use MultiLevelType to enable proper multi-level behavior like pressing Tab in Word
        abstractNum.Append(new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel });

        // Get base left indent from the LP style (e.g., LP1, LP2, LP3, LP4)
        int baseLeftIndent = GetStyleLeftIndent(document, $"LP{headingLevel}");

        // Define all 9 levels (0-8) as Word expects for proper Tab key behavior
        for (int i = 0; i < 9; i++)
        {
            var level = CreateLevel(i, listType, baseLeftIndent);
            abstractNum.Append(level);
        }

        return abstractNum;
    }

    /// <summary>
    /// Gets the left indentation value from a paragraph style in the document.
    /// </summary>
    /// <param name="document">The Word document.</param>
    /// <param name="styleId">The style ID to look up (e.g., "LP1").</param>
    /// <returns>The left indent in twips, or a default value if not found.</returns>
    private static int GetStyleLeftIndent(WordprocessingDocument document, string styleId)
    {
        const int defaultLeftIndent = 720; // Default 0.5 inch if style not found

        var stylesPart = document.MainDocumentPart?.StyleDefinitionsPart;
        if (stylesPart?.Styles == null)
            return defaultLeftIndent;

        var style = stylesPart.Styles.Elements<Style>()
            .FirstOrDefault(s => s.StyleId?.Value == styleId && s.Type?.Value == StyleValues.Paragraph);

        if (style == null)
            return defaultLeftIndent;

        // Get the left indent from the style's paragraph properties
        var indent = style.StyleParagraphProperties?.Indentation;
        if (indent?.Left?.Value != null && int.TryParse(indent.Left.Value, out int leftIndent))
        {
            return leftIndent + 720; // 720 = Extra 0.5 inches
        }

        return defaultLeftIndent;
    }

    /// <summary>
    /// Creates a level definition for a multi-level list.
    /// Uses Word's default bullet/number patterns and indentation based on style.
    /// </summary>
    /// <param name="levelIndex">The level index (0-8).</param>
    /// <param name="listType">The type of list (bulleted or numbered).</param>
    /// <param name="baseLeftIndent">The base left indent from the LP style in twips.</param>
    private static Level CreateLevel(int levelIndex, DocxListType listType, int baseLeftIndent)
    {
        var level = new Level() { LevelIndex = levelIndex };

        // Add template ID for hybrid multi-level (required for proper Tab behavior)
        level.Append(new TemplateCode() { Val = Guid.NewGuid().ToString("N").Substring(0, 8).ToUpper() });

        if (listType == DocxListType.Bulleted)
        {
            level.Append(new NumberingFormat() { Val = NumberFormatValues.Bullet });

            // Word's default bullet pattern: solid bullet, hollow circle, square, then repeats
            (string bulletChar, string fontName) = (levelIndex % 3) switch
            {
                0 => ("\uF0B7", "Symbol"),       // Solid bullet (●)
                1 => ("o", "Courier New"),       // Hollow circle (○)
                _ => ("\uF0A7", "Wingdings")     // Square (■)
            };

            level.Append(new LevelText() { Val = bulletChar });
            level.Append(new NumberingSymbolRunProperties(
                new RunFonts() { Ascii = fontName, HighAnsi = fontName, Hint = FontTypeHintValues.Default }
            ));
        }
        else // Numbered
        {
            level.Append(new StartNumberingValue() { Val = 1 });

            // Word's default numbered list pattern
            (NumberFormatValues format, string textPattern) = (levelIndex % 3) switch
            {
                0 => (NumberFormatValues.Decimal, $"%{levelIndex + 1}."),      // 1., 2., 3.
                1 => (NumberFormatValues.LowerLetter, $"%{levelIndex + 1}."),  // a., b., c.
                _ => (NumberFormatValues.LowerRoman, $"%{levelIndex + 1}.")    // i., ii., iii.
            };

            level.Append(new NumberingFormat() { Val = format });
            level.Append(new LevelText() { Val = textPattern });
        }

        // Use LevelSuffix to add tab after the number/bullet (Word's default)
        level.Append(new LevelSuffix() { Val = LevelSuffixValues.Tab });

        // Calculate indentation based on the LP style's left indent
        // - Level 0 uses the base left indent from the LP style
        // - Each subsequent level adds 720 twips (0.5 inch)
        // - Hanging indent of 360 twips (0.25 inch) for the bullet/number
        int leftIndent = baseLeftIndent + (720 * levelIndex);
        int hangingIndent = 360;

        level.Append(new PreviousParagraphProperties(
            new Indentation()
            {
                Left = leftIndent.ToString(),
                Hanging = hangingIndent.ToString()
            }
        ));

        return level;
    }

    /// <summary>
    /// Ensures there's a current paragraph to add content to.
    /// </summary>
    private void EnsureCurrentParagraph()
    {
        if (_currentParagraph == null)
        {
            NewParagraph();
        }
    }

    /// <summary>
    /// Finalizes the current context (list, table, etc.) when switching to a new context.
    /// </summary>
    private void FinalizeCurrentContext()
    {
        _currentParagraph = null;
        _currentListType = null;
        _currentTable = null;
        _currentTableRow = null;
        _currentTableCell = null;
        _currentTableRowIndex = 0;
    }

    /// <summary>
    /// Clears all collected content and resets state.
    /// </summary>
    public void Clear()
    {
        _elements.Clear();
        _listDefinitions.Clear();
        _currentHeadingLevel = 1;
        _localListId = 0;
        _figureSequenceNumber = 0;
        _tableSequenceNumber = 0;
        _referencesByKey.Clear();
        FinalizeCurrentContext();
    }
}