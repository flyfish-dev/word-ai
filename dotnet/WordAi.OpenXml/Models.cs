using System.Text.Json;
using System.Text.Json.Serialization;

namespace WordAi.OpenXml;

public static class JsonSupport
{
    public static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };
}

public sealed record Anchor(
    string AnchorId,
    string Kind,
    string Label,
    string Path,
    string? TextPreview,
    string? StyleId,
    int? Level,
    Dictionary<string, object?> Extra);

public sealed record DocumentProfile(
    string Path,
    string Sha256,
    long SizeBytes,
    int PartsCount,
    int ParagraphCount,
    int TableCount,
    int ImageCount,
    int FieldCount,
    int CommentCount,
    int CommentReferenceCount,
    int TrackedChangeCount,
    int ContentControlCount,
    int HeadingCount,
    int BookmarkCount,
    IReadOnlyList<Anchor> Anchors);

public sealed class PatchGuard
{
    public bool RequirePreconditions { get; set; }
    public bool AllowOverwrite { get; set; }
}

[JsonPolymorphic(TypeDiscriminatorPropertyName = "op")]
[JsonDerivedType(typeof(ReplaceContentControlTextOperation), "replace_content_control_text")]
[JsonDerivedType(typeof(AppendContentControlTextOperation), "append_content_control_text")]
[JsonDerivedType(typeof(PrependContentControlTextOperation), "prepend_content_control_text")]
[JsonDerivedType(typeof(ReplaceTextInContentControlOperation), "replace_text_in_content_control")]
[JsonDerivedType(typeof(ReplaceParagraphTextOperation), "replace_paragraph_text")]
[JsonDerivedType(typeof(InsertParagraphAfterOperation), "insert_paragraph_after")]
[JsonDerivedType(typeof(InsertParagraphBeforeOperation), "insert_paragraph_before")]
[JsonDerivedType(typeof(ReplaceTableCellTextOperation), "replace_table_cell_text")]
[JsonDerivedType(typeof(AppendTableRowOperation), "append_table_row")]
[JsonDerivedType(typeof(WrapParagraphWithContentControlOperation), "wrap_paragraph_with_content_control")]
[JsonDerivedType(typeof(AddCommentOperation), "add_comment")]
public abstract class PatchOperation
{
    public string? ExpectedOldText { get; set; }
    public string? ExpectedOldSha256 { get; set; }
}

public sealed class ReplaceContentControlTextOperation : PatchOperation
{
    public string Tag { get; set; } = "";
    public string Text { get; set; } = "";
    public bool PreserveStyle { get; set; } = true;
    public bool AllowParagraphCountChange { get; set; }
    public bool AllowComplexContent { get; set; }
}

public sealed class AppendContentControlTextOperation : PatchOperation
{
    public string Tag { get; set; } = "";
    public string Text { get; set; } = "";
}

public sealed class PrependContentControlTextOperation : PatchOperation
{
    public string Tag { get; set; } = "";
    public string Text { get; set; } = "";
}

public sealed class ReplaceTextInContentControlOperation : PatchOperation
{
    public string Tag { get; set; } = "";
    public string Find { get; set; } = "";
    public string Replace { get; set; } = "";
    public string Occurrence { get; set; } = "all";
    public bool RequireMatch { get; set; } = true;
}

public abstract class ParagraphScopedOperation : PatchOperation
{
    public string? ParaId { get; set; }
    public int? ParagraphIndex { get; set; }
    public string? Tag { get; set; }
    public string? ContentControlTag { get; set; }
    public string? TargetTag { get; set; }
}

public sealed class ReplaceParagraphTextOperation : ParagraphScopedOperation
{
    public string Text { get; set; } = "";
    public bool PreserveStyle { get; set; } = true;
    public bool AllowComplexContent { get; set; }
}

public sealed class InsertParagraphAfterOperation : ParagraphScopedOperation
{
    public string Text { get; set; } = "";
    public bool InheritStyle { get; set; } = true;
    public bool InheritHeadingStyle { get; set; }
}

public sealed class InsertParagraphBeforeOperation : ParagraphScopedOperation
{
    public string Text { get; set; } = "";
    public bool InheritStyle { get; set; } = true;
    public bool InheritHeadingStyle { get; set; }
}

public sealed class ReplaceTableCellTextOperation : PatchOperation
{
    public string? ScopeTag { get; set; }
    public int TableIndex { get; set; } = 1;
    public int Row { get; set; } = 1;
    public int Col { get; set; } = 1;
    public string Text { get; set; } = "";
    public bool PreserveStyle { get; set; } = true;
    public bool AllowParagraphCountChange { get; set; }
    public bool AllowComplexContent { get; set; }
}

public sealed class AppendTableRowOperation : PatchOperation
{
    public string? ScopeTag { get; set; }
    public int TableIndex { get; set; } = 1;
    public int? TemplateRow { get; set; }
    public List<string> Values { get; set; } = [];
}

public sealed class WrapParagraphWithContentControlOperation : ParagraphScopedOperation
{
    public new string Tag { get; set; } = "";
    public string? Alias { get; set; }
    public string? Title { get; set; }
    public bool Lock { get; set; } = true;
}

public sealed class AddCommentOperation : ParagraphScopedOperation
{
    public string Text { get; set; } = "";
    public string Author { get; set; } = "Word AI";
    public string Initials { get; set; } = "AI";
}

public sealed class PatchSet
{
    public string SchemaVersion { get; set; } = "2.0";
    public bool Strict { get; set; } = true;
    public string? Reason { get; set; }
    public string? SourceSha256 { get; set; }
    public PatchGuard Guard { get; set; } = new();
    public bool Overwrite { get; set; }
    public bool AbortOnValidationError { get; set; } = true;
    public bool KeepInvalidOutput { get; set; }
    public List<PatchOperation> Operations { get; set; } = [];
}

public sealed record Risk(string Severity, string Code, string Message, int? OperationIndex = null, Dictionary<string, object?>? Extra = null);

public sealed record TouchedScope(
    IReadOnlyList<string> ContentControlTags,
    IReadOnlyList<string> ParaIds,
    IReadOnlyList<int> ParagraphIndices,
    IReadOnlyList<int> TableIndices,
    IReadOnlyList<string> TableCells);

public sealed record PatchAssessment(
    bool Ok,
    int OperationCount,
    bool RequiresStructuralChange,
    TouchedScope Touched,
    IReadOnlyList<Risk> Risks);

public sealed record ValidationIssue(string Severity, string Code, string Message, string? Path = null);

public sealed record ValidationReport(
    bool Ok,
    string SourcePath,
    string TargetPath,
    IReadOnlyList<ValidationIssue> Issues,
    Dictionary<string, object?> Metrics);

public sealed class ValidationOptions
{
    public HashSet<string> AllowedPartChanges { get; } = new(StringComparer.OrdinalIgnoreCase) { "word/document.xml" };
    public HashSet<string> AllowedCountChanges { get; } = new(StringComparer.OrdinalIgnoreCase);
    public HashSet<string> TouchedContentControlTags { get; } = new(StringComparer.Ordinal);
    public HashSet<string> TouchedParaIds { get; } = new(StringComparer.Ordinal);
    public HashSet<int> TouchedParagraphIndices { get; } = [];
    public HashSet<int> TouchedTableIndices { get; } = [];
    public HashSet<string> TouchedTableCells { get; } = new(StringComparer.Ordinal);
    public HashSet<string> AllowedAddedContentControlTags { get; } = new(StringComparer.Ordinal);
    public bool AllowTableDimensionChange { get; set; }
    public bool AllowParagraphCountChange { get; set; }
}

public sealed record PatchAudit(
    string SourcePath,
    string OutputPath,
    string? AuditPath,
    string? PatchsetReason,
    PatchAssessment SafetyAssessment,
    IReadOnlyList<Dictionary<string, object?>> Applied,
    ValidationReport Validation,
    Dictionary<string, int> BeforeMetrics,
    Dictionary<string, int> AfterMetrics,
    string ChangedTextDiff,
    bool DryRun,
    bool KeptOutput,
    string CreatedAt);
