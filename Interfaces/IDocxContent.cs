using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using md_to_docx_sync.Enums;

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

    // Output
    IReadOnlyList<OpenXmlElement> Elements { get; }
    void WriteTo(WordprocessingDocument document, OpenXmlElement insertAfterElement);
    void Clear();
}