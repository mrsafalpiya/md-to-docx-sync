using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using md_to_docx_sync.Enums;
using md_to_docx_sync.Models;

namespace md_to_docx_sync.Interfaces;

public interface IDocxContent
{
    void NewHeading(int level);
    void NewParagraph(bool isLiteral = false);

    // Text
    void AddText(string text, bool isBold = false, bool isItalic = false, bool isUnderline = false, bool isStrikethrough = false, bool isLiteral = false);
    void AddLineBreak();

    // List
    void NewList(DocxListType type);
    void NewListItem(int level = 0);

    // Table
    /// <summary>
    /// Creates a new table.
    /// </summary>
    /// <param name="caption">Optional caption for the table.</param>
    /// <param name="widthPercent">Optional width as percentage (1-100). Defaults to 100 (full width).</param>
    void NewTable(string? caption = null, int? widthPercent = null);
    void NewRow();
    void NewCell();

    // Images
    /// <summary>
    /// Adds an image to the document.
    /// </summary>
    /// <param name="imagePath">Path to the image file.</param>
    /// <param name="caption">Optional caption for the image.</param>
    /// <param name="widthInches">Optional width in inches. If not provided, uses the original image width. Height is auto-calculated to maintain aspect ratio.</param>
    void AddImage(string imagePath, string? caption = null, double? widthInches = null);

    // References
    /// <summary>
    /// Adds inline CITATION field(s) referencing bibliography source(s) by key.
    /// When multiple keys are provided, they are merged into a grouped citation using the \m switch.
    /// </summary>
    void AddCitation(params string[] keys);

    /// <summary>
    /// Stores bibliography sources to be written into the document's CustomXmlPart during WriteTo().
    /// Does not produce visible output; the template's existing bibliography section renders the sources.
    /// </summary>
    void AddBibliography(List<ReferenceSource> references);

    /// <summary>
    /// Removes the last heading element from the collected content.
    /// Used to strip the heading above a references block so it doesn't appear in the output.
    /// </summary>
    void RemoveLastHeading();

    // Output
    IReadOnlyList<OpenXmlElement> Elements { get; }
    void WriteTo(WordprocessingDocument document, OpenXmlElement insertAfterElement);
    void Clear();
}