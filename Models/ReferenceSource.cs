namespace md_to_docx_sync.Models;

public class ReferenceSource
{
    public required string Key { get; set; }
    public required string Type { get; set; }
    public Dictionary<string, string> Fields { get; set; } = new();
}
