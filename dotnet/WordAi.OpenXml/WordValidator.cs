using System.IO.Compression;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordAi.OpenXml;

public sealed class WordValidator
{
    public ValidationReport Validate(string sourceDocx, string targetDocx, bool strict = true, ValidationOptions? options = null)
    {
        options ??= new ValidationOptions();
        var issues = new List<ValidationIssue>();
        var metrics = new Dictionary<string, object?>();

        try
        {
            using var zip = ZipFile.OpenRead(targetDocx);
            var bad = zip.Entries.FirstOrDefault(e =>
            {
                try
                {
                    using var s = e.Open();
                    Span<byte> buffer = stackalloc byte[1024];
                    while (s.Read(buffer) > 0) { }
                    return false;
                }
                catch
                {
                    return true;
                }
            });
            if (bad is not null) issues.Add(new ValidationIssue("error", "zip_corrupt", $"Corrupt ZIP member: {bad.FullName}"));
            metrics["target_part_count"] = zip.Entries.Count(e => !string.IsNullOrEmpty(e.Name));
        }
        catch (Exception ex)
        {
            issues.Add(new ValidationIssue("error", "zip_open_failed", ex.Message));
            return new ValidationReport(false, sourceDocx, targetDocx, issues, metrics);
        }

        AddOpenXmlValidatorIssues(sourceDocx, targetDocx, issues, strict, metrics);
        ComparePackageParts(sourceDocx, targetDocx, strict, options, issues, metrics);

        try
        {
            var before = new WordInspector().Inspect(sourceDocx);
            var after = new WordInspector().Inspect(targetDocx);
            CompareCounts(before, after, strict, options, issues, metrics);
            CheckProtectedObjects(sourceDocx, targetDocx, strict, options, issues, metrics);
        }
        catch (Exception ex)
        {
            issues.Add(new ValidationIssue("error", "inspection_failed", ex.Message));
        }

        return new ValidationReport(!issues.Any(i => i.Severity == "error"), sourceDocx, targetDocx, issues, metrics);
    }

    private static void AddOpenXmlValidatorIssues(string sourceDocx, string targetDocx, List<ValidationIssue> issues, bool strict, Dictionary<string, object?> metrics)
    {
        try
        {
            var sourceErrors = ValidatorFingerprints(sourceDocx);
            var targetErrors = ValidatorFingerprints(targetDocx);
            var newErrors = targetErrors.Where(e => !sourceErrors.Contains(e)).Take(100).ToList();
            metrics["source_openxml_validator_error_count"] = sourceErrors.Count;
            metrics["target_openxml_validator_error_count"] = targetErrors.Count;
            metrics["new_openxml_validator_error_count"] = newErrors.Count;
            foreach (var error in newErrors)
            {
                issues.Add(new ValidationIssue(
                    strict ? "error" : "warning",
                    "openxml_validator",
                    error,
                    null));
            }
        }
        catch (Exception ex)
        {
            issues.Add(new ValidationIssue("error", "openxml_validator_failed", ex.Message));
        }
    }

    private static HashSet<string> ValidatorFingerprints(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        var validator = new OpenXmlValidator();
        return validator.Validate(doc)
            .Select(error => $"{error.Part?.Uri}|{error.Path?.XPath}|{error.Description}")
            .ToHashSet(StringComparer.Ordinal);
    }

    private static void ComparePackageParts(
        string sourceDocx,
        string targetDocx,
        bool strict,
        ValidationOptions options,
        List<ValidationIssue> issues,
        Dictionary<string, object?> metrics)
    {
        var sourceHashes = WordInspector.HashParts(sourceDocx);
        var targetHashes = WordInspector.HashParts(targetDocx);
        var removed = sourceHashes.Keys.Except(targetHashes.Keys, StringComparer.OrdinalIgnoreCase).Order().ToList();
        var added = targetHashes.Keys.Except(sourceHashes.Keys, StringComparer.OrdinalIgnoreCase).Order().ToList();
        var changed = sourceHashes.Keys.Intersect(targetHashes.Keys, StringComparer.OrdinalIgnoreCase)
            .Where(k => !string.Equals(sourceHashes[k], targetHashes[k], StringComparison.OrdinalIgnoreCase))
            .Order()
            .ToList();

        metrics["parts_removed"] = removed;
        metrics["parts_added"] = added;
        metrics["changed_parts"] = changed;

        if (removed.Count > 0) issues.Add(new ValidationIssue("error", "parts_removed", $"Removed ZIP parts: {string.Join(", ", removed)}"));

        var unexpectedAdded = added.Where(p => !options.AllowedPartChanges.Contains(p)).ToList();
        if (unexpectedAdded.Count > 0)
        {
            issues.Add(new ValidationIssue(strict ? "error" : "warning", "parts_added", $"Added ZIP parts: {string.Join(", ", unexpectedAdded)}"));
        }

        var unexpectedChanged = changed.Where(p => !options.AllowedPartChanges.Contains(p)).ToList();
        if (unexpectedChanged.Count > 0)
        {
            issues.Add(new ValidationIssue(strict ? "error" : "warning", "unexpected_part_change", $"Unexpected changed parts: {string.Join(", ", unexpectedChanged)}"));
        }
    }

    private static void CompareCounts(DocumentProfile before, DocumentProfile after, bool strict, ValidationOptions options, List<ValidationIssue> issues, Dictionary<string, object?> metrics)
    {
        var beforeCounts = Metrics(before);
        var afterCounts = Metrics(after);
        metrics["before"] = beforeCounts;
        metrics["after"] = afterCounts;

        foreach (var key in new[] { "table_count", "image_count", "field_count", "comment_count", "comment_reference_count", "tracked_change_count", "content_control_count", "heading_count" })
        {
            if (beforeCounts[key] != afterCounts[key] && !options.AllowedCountChanges.Contains(key))
            {
                issues.Add(new ValidationIssue(strict ? "error" : "warning", "structure_count_changed", $"{key} changed from {beforeCounts[key]} to {afterCounts[key]}"));
            }
        }

        if (before.ParagraphCount != after.ParagraphCount && !options.AllowParagraphCountChange && !options.AllowedCountChanges.Contains("paragraph_count"))
        {
            issues.Add(new ValidationIssue(strict ? "error" : "warning", "paragraph_count_changed", $"paragraph_count changed from {before.ParagraphCount} to {after.ParagraphCount}"));
        }

        var beforeTags = before.Anchors.Where(a => a.Kind == "content_control").Select(a => a.Extra.GetValueOrDefault("tag")?.ToString()).Where(x => !string.IsNullOrEmpty(x)).Order().ToList();
        var afterTags = after.Anchors.Where(a => a.Kind == "content_control").Select(a => a.Extra.GetValueOrDefault("tag")?.ToString()).Where(x => !string.IsNullOrEmpty(x) && !options.AllowedAddedContentControlTags.Contains(x!)).Order().ToList();
        if (!beforeTags.SequenceEqual(afterTags))
        {
            issues.Add(new ValidationIssue("error", "content_control_tags_changed", "Content control tag set changed"));
        }
    }

    private static void CheckProtectedObjects(string sourceDocx, string targetDocx, bool strict, ValidationOptions options, List<ValidationIssue> issues, Dictionary<string, object?> metrics)
    {
        var isolationRequested = options.TouchedContentControlTags.Count > 0
                                 || options.TouchedParaIds.Count > 0
                                 || options.TouchedParagraphIndices.Count > 0
                                 || options.TouchedTableIndices.Count > 0
                                 || options.TouchedTableCells.Count > 0;
        if (!isolationRequested)
        {
            metrics["protected_object_checks"] = new Dictionary<string, object?> { ["skipped"] = "no touched targets supplied; structural package checks only" };
            return;
        }

        using var beforeDoc = WordprocessingDocument.Open(sourceDocx, false);
        using var afterDoc = WordprocessingDocument.Open(targetDocx, false);
        var beforeBody = beforeDoc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Missing source body");
        var afterBody = afterDoc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Missing target body");

        CheckContentControlHashes(beforeBody, afterBody, strict, options, issues);
        CheckTableHashes(beforeBody, afterBody, strict, options, issues);
        CheckParagraphHashes(beforeBody, afterBody, strict, options, issues);
        CheckBodyBlockSequence(beforeBody, afterBody, strict, options, issues, metrics);

        metrics["protected_object_checks"] = new Dictionary<string, object?>
        {
            ["touched_content_control_tags"] = options.TouchedContentControlTags.Order().ToList(),
            ["touched_para_ids"] = options.TouchedParaIds.Order().ToList(),
            ["touched_paragraph_indices"] = options.TouchedParagraphIndices.Order().ToList(),
            ["touched_table_indices"] = options.TouchedTableIndices.Order().ToList(),
            ["touched_table_cells"] = options.TouchedTableCells.Order().ToList(),
        };
    }

    private static void CheckContentControlHashes(Body beforeBody, Body afterBody, bool strict, ValidationOptions options, List<ValidationIssue> issues)
    {
        var before = beforeBody.Descendants<SdtElement>()
            .Select(s => (Tag: WordInspector.ContentControlTag(s), Hash: WordInspector.Sha256String(s.OuterXml)))
            .Where(x => !string.IsNullOrEmpty(x.Tag))
            .ToDictionary(x => x.Tag!, x => x.Hash, StringComparer.Ordinal);
        var after = afterBody.Descendants<SdtElement>()
            .Select(s => (Tag: WordInspector.ContentControlTag(s), Hash: WordInspector.Sha256String(s.OuterXml)))
            .Where(x => !string.IsNullOrEmpty(x.Tag))
            .ToDictionary(x => x.Tag!, x => x.Hash, StringComparer.Ordinal);
        var changed = before.Where(kv => !options.TouchedContentControlTags.Contains(kv.Key) && after.GetValueOrDefault(kv.Key) != kv.Value).Select(kv => kv.Key).ToList();
        if (changed.Count > 0)
        {
            issues.Add(new ValidationIssue(strict ? "error" : "warning", "protected_content_control_changed", $"Untouched content controls changed: {string.Join(", ", changed.Take(20))}"));
        }
    }

    private static void CheckTableHashes(Body beforeBody, Body afterBody, bool strict, ValidationOptions options, List<ValidationIssue> issues)
    {
        var beforeTables = beforeBody.Descendants<Table>().ToList();
        var afterTables = afterBody.Descendants<Table>().ToList();
        for (var i = 0; i < Math.Min(beforeTables.Count, afterTables.Count); i++)
        {
            var index = i + 1;
            var beforeTable = beforeTables[i];
            var afterTable = afterTables[i];
            if (options.TouchedTableIndices.Contains(index))
            {
                var beforeRows = beforeTable.Elements<TableRow>().ToList();
                var afterRows = afterTable.Elements<TableRow>().ToList();
                var beforeColumns = beforeRows.Select(r => r.Elements<TableCell>().Count()).ToList();
                var afterColumns = afterRows.Select(r => r.Elements<TableCell>().Count()).ToList();
                if (!options.AllowTableDimensionChange && (beforeRows.Count != afterRows.Count || !beforeColumns.SequenceEqual(afterColumns)))
                {
                    issues.Add(new ValidationIssue("error", "touched_table_dimension_changed_without_permission", $"Table {index} dimensions changed without permission"));
                }
                CheckTableCellHashes(index, beforeTable, afterTable, strict, options, issues);
                continue;
            }

            if (WordInspector.Sha256String(beforeTable.OuterXml) != WordInspector.Sha256String(afterTable.OuterXml))
            {
                issues.Add(new ValidationIssue(strict ? "error" : "warning", "protected_table_changed", $"Untouched table changed: {index}"));
            }
        }
    }

    private static void CheckTableCellHashes(int tableIndex, Table beforeTable, Table afterTable, bool strict, ValidationOptions options, List<ValidationIssue> issues)
    {
        var beforeRows = beforeTable.Elements<TableRow>().ToList();
        var afterRows = afterTable.Elements<TableRow>().ToList();
        for (var r = 0; r < Math.Min(beforeRows.Count, afterRows.Count); r++)
        {
            var beforeCells = beforeRows[r].Elements<TableCell>().ToList();
            var afterCells = afterRows[r].Elements<TableCell>().ToList();
            for (var c = 0; c < Math.Min(beforeCells.Count, afterCells.Count); c++)
            {
                var cellRef = $"{tableIndex}:{r + 1}:{c + 1}";
                if (options.TouchedTableCells.Contains(cellRef)) continue;
                if (WordInspector.Sha256String(beforeCells[c].OuterXml) != WordInspector.Sha256String(afterCells[c].OuterXml))
                {
                    issues.Add(new ValidationIssue(strict ? "error" : "warning", "protected_table_cell_changed", $"Untouched table cell changed: {cellRef}"));
                }
            }
        }
    }

    private static void CheckParagraphHashes(Body beforeBody, Body afterBody, bool strict, ValidationOptions options, List<ValidationIssue> issues)
    {
        var beforeByParaId = beforeBody.Descendants<Paragraph>()
            .Select(p => (ParaId: WordInspector.ParaId(p), Paragraph: p))
            .Where(x => !string.IsNullOrEmpty(x.ParaId))
            .ToDictionary(x => x.ParaId!, x => x.Paragraph, StringComparer.Ordinal);
        var afterByParaId = afterBody.Descendants<Paragraph>()
            .Select(p => (ParaId: WordInspector.ParaId(p), Paragraph: p))
            .Where(x => !string.IsNullOrEmpty(x.ParaId))
            .ToDictionary(x => x.ParaId!, x => x.Paragraph, StringComparer.Ordinal);

        var changed = new List<string>();
        foreach (var (paraId, beforeParagraph) in beforeByParaId)
        {
            if (options.TouchedParaIds.Contains(paraId)) continue;
            if (WordInspector.AncestorContentControlTag(beforeParagraph) is { } tag && options.TouchedContentControlTags.Contains(tag)) continue;
            if (WordInspector.AncestorTableIndex(beforeParagraph, beforeBody) is { } tableIndex && options.TouchedTableIndices.Contains(tableIndex)) continue;
            if (!afterByParaId.TryGetValue(paraId, out var afterParagraph)) continue;
            if (WordInspector.Sha256String(beforeParagraph.OuterXml) != WordInspector.Sha256String(afterParagraph.OuterXml)) changed.Add(paraId);
        }
        if (changed.Count > 0)
        {
            issues.Add(new ValidationIssue(strict ? "error" : "warning", "protected_paragraph_changed", $"Untouched paragraphs with paraId changed: {string.Join(", ", changed.Take(20))}"));
        }
    }

    private static void CheckBodyBlockSequence(Body beforeBody, Body afterBody, bool strict, ValidationOptions options, List<ValidationIssue> issues, Dictionary<string, object?> metrics)
    {
        var beforeBlocks = BodyBlocks(beforeBody, options).Where(b => !b.Touched).ToList();
        var afterBlocks = BodyBlocks(afterBody, options).Where(b => !b.Touched).ToList();
        var cursor = 0;
        var missing = new List<Dictionary<string, object?>>();
        foreach (var block in beforeBlocks)
        {
            while (cursor < afterBlocks.Count && afterBlocks[cursor].Hash != block.Hash) cursor++;
            if (cursor >= afterBlocks.Count)
            {
                missing.Add(new Dictionary<string, object?>
                {
                    ["block_index"] = block.Index,
                    ["kind"] = block.Kind,
                    ["text_sha256"] = block.TextSha256,
                    ["content_control_tags"] = block.ContentControlTags,
                    ["table_indices"] = block.TableIndices,
                    ["paragraph_indices"] = block.ParagraphIndices,
                });
            }
            else
            {
                cursor++;
            }
        }

        metrics["body_block_sequence"] = new Dictionary<string, object?>
        {
            ["ok"] = missing.Count == 0,
            ["protected_source_block_count"] = beforeBlocks.Count,
            ["protected_target_block_count"] = afterBlocks.Count,
            ["missing_or_modified_protected_blocks"] = missing.Take(50).ToList(),
        };
        if (missing.Count > 0)
        {
            issues.Add(new ValidationIssue(strict ? "error" : "warning", "protected_body_block_changed", "Untouched body blocks were modified, reordered, or removed"));
        }
    }

    private static List<BodyBlock> BodyBlocks(Body body, ValidationOptions options)
    {
        var allTables = body.Descendants<Table>().ToList();
        var allParagraphs = body.Descendants<Paragraph>().ToList();
        var result = new List<BodyBlock>();
        var index = 0;
        foreach (var child in body.ChildElements)
        {
            if (child is SectionProperties) continue;
            index++;
            var tags = child.Descendants<SdtElement>().Select(WordInspector.ContentControlTag).Where(t => !string.IsNullOrEmpty(t)).Cast<string>().Distinct().Order().ToList();
            if (child is SdtElement sdt && WordInspector.ContentControlTag(sdt) is { } ownTag && !tags.Contains(ownTag)) tags.Insert(0, ownTag);
            var tableIndices = child.Descendants<Table>().Select(t => allTables.IndexOf(t) + 1).Where(i => i > 0).Distinct().Order().ToList();
            if (child is Table table)
            {
                var tableIndex = allTables.IndexOf(table) + 1;
                if (tableIndex > 0 && !tableIndices.Contains(tableIndex)) tableIndices.Insert(0, tableIndex);
            }
            var paragraphIndices = child.Descendants<Paragraph>().Select(p => allParagraphs.IndexOf(p) + 1).Where(i => i > 0).Distinct().Order().ToList();
            if (child is Paragraph paragraph)
            {
                var paragraphIndex = allParagraphs.IndexOf(paragraph) + 1;
                if (paragraphIndex > 0 && !paragraphIndices.Contains(paragraphIndex)) paragraphIndices.Insert(0, paragraphIndex);
            }
            var touched = tags.Any(options.TouchedContentControlTags.Contains)
                          || tableIndices.Any(options.TouchedTableIndices.Contains)
                          || paragraphIndices.Any(options.TouchedParagraphIndices.Contains);
            var text = child is Paragraph p ? WordInspector.ParagraphText(p) : string.Join("\n", child.Descendants<Paragraph>().Select(WordInspector.ParagraphText));
            result.Add(new BodyBlock(index, child.GetType().Name, WordInspector.Sha256String(child.OuterXml), WordInspector.Sha256String(text), tags, tableIndices, paragraphIndices, touched));
        }
        return result;
    }

    private static Dictionary<string, int> Metrics(DocumentProfile profile) => new()
    {
        ["parts_count"] = profile.PartsCount,
        ["paragraph_count"] = profile.ParagraphCount,
        ["table_count"] = profile.TableCount,
        ["image_count"] = profile.ImageCount,
        ["field_count"] = profile.FieldCount,
        ["comment_count"] = profile.CommentCount,
        ["comment_reference_count"] = profile.CommentReferenceCount,
        ["tracked_change_count"] = profile.TrackedChangeCount,
        ["content_control_count"] = profile.ContentControlCount,
        ["heading_count"] = profile.HeadingCount,
        ["bookmark_count"] = profile.BookmarkCount,
    };

    private sealed record BodyBlock(
        int Index,
        string Kind,
        string Hash,
        string TextSha256,
        IReadOnlyList<string> ContentControlTags,
        IReadOnlyList<int> TableIndices,
        IReadOnlyList<int> ParagraphIndices,
        bool Touched);
}
