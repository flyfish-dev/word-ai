using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordAi.OpenXml;

public sealed class WordPatchApplier
{
    private static readonly HashSet<Type> HighRiskOperationTypes =
    [
        typeof(ReplaceParagraphTextOperation),
        typeof(InsertParagraphAfterOperation),
        typeof(InsertParagraphBeforeOperation),
        typeof(ReplaceTableCellTextOperation),
        typeof(AppendTableRowOperation),
        typeof(WrapParagraphWithContentControlOperation),
        typeof(AddCommentOperation),
    ];

    public PatchAssessment AssessPatchSet(string sourceDocx, PatchSet patchSet)
    {
        if (patchSet.Operations.Count == 0) throw new InvalidOperationException("patchset.operations must contain at least one operation");

        var risks = new List<Risk>();
        var touchedTags = new HashSet<string>(StringComparer.Ordinal);
        var touchedParaIds = new HashSet<string>(StringComparer.Ordinal);
        var touchedParagraphIndices = new HashSet<int>();
        var touchedTableIndices = new HashSet<int>();
        var touchedTableCells = new HashSet<string>(StringComparer.Ordinal);
        var requiresStructuralChange = false;

        if (patchSet.Guard.RequirePreconditions && string.IsNullOrWhiteSpace(patchSet.SourceSha256))
        {
            risks.Add(new Risk("error", "missing_source_sha256", "guard.require_preconditions=true requires patchset.source_sha256."));
        }

        using var doc = WordprocessingDocument.Open(sourceDocx, false);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Missing document body.");

        for (var i = 0; i < patchSet.Operations.Count; i++)
        {
            var op = patchSet.Operations[i];
            try
            {
                switch (op)
                {
                    case ReplaceContentControlTextOperation cc:
                        var sdt = FindContentControlByTag(body, cc.Tag);
                        touchedTags.Add(cc.Tag);
                        var complexity = WordInspector.ComplexSummary(sdt);
                        if (complexity.Values.Any(v => v > 0) && !cc.AllowComplexContent)
                        {
                            risks.Add(new Risk("error", "complex_content_control", "Replacement target contains complex Word objects.", i, complexity.ToDictionary(kv => kv.Key, kv => (object?)kv.Value)));
                        }
                        AddMissingPreconditionRisk(op, patchSet, risks, i, forceError: false, "content-control operation");
                        break;
                    case AppendContentControlTextOperation append:
                        FindContentControlByTag(body, append.Tag);
                        touchedTags.Add(append.Tag);
                        requiresStructuralChange = true;
                        AddMissingPreconditionRisk(op, patchSet, risks, i, forceError: false, "content-control append operation");
                        break;
                    case PrependContentControlTextOperation prepend:
                        FindContentControlByTag(body, prepend.Tag);
                        touchedTags.Add(prepend.Tag);
                        requiresStructuralChange = true;
                        AddMissingPreconditionRisk(op, patchSet, risks, i, forceError: false, "content-control prepend operation");
                        break;
                    case ReplaceTextInContentControlOperation textOp:
                        FindContentControlByTag(body, textOp.Tag);
                        touchedTags.Add(textOp.Tag);
                        AddMissingPreconditionRisk(op, patchSet, risks, i, forceError: false, "content-control find/replace operation");
                        break;
                    case ParagraphScopedOperation paragraphScoped when op is ReplaceParagraphTextOperation or InsertParagraphAfterOperation or InsertParagraphBeforeOperation or AddCommentOperation:
                        var p = FindTargetParagraph(body, paragraphScoped, wrapMode: false);
                        TrackParagraph(body, p, touchedParaIds, touchedParagraphIndices);
                        requiresStructuralChange = op is InsertParagraphAfterOperation or InsertParagraphBeforeOperation or AddCommentOperation;
                        if (op is ReplaceParagraphTextOperation replacePara && WordInspector.HasComplexContent(p) && !replacePara.AllowComplexContent)
                        {
                            risks.Add(new Risk("error", "complex_paragraph", "Target paragraph contains complex Word objects.", i));
                        }
                        AddMissingPreconditionRisk(op, patchSet, risks, i, forceError: true, "paragraph-scoped operation");
                        break;
                    case ReplaceTableCellTextOperation cell:
                        var table = FindTable(body, cell.TableIndex, cell.ScopeTag);
                        var globalTableIndex = GlobalTableIndex(body, table);
                        var tableCell = FindCell(table, cell.Row, cell.Col);
                        touchedTableIndices.Add(globalTableIndex);
                        touchedTableCells.Add($"{globalTableIndex}:{cell.Row}:{cell.Col}");
                        if (WordInspector.HasComplexContent(tableCell) && !cell.AllowComplexContent)
                        {
                            risks.Add(new Risk("error", "complex_table_cell", "Target table cell contains complex Word objects.", i));
                        }
                        AddMissingPreconditionRisk(op, patchSet, risks, i, forceError: true, "table-cell operation");
                        break;
                    case AppendTableRowOperation row:
                        var rowTable = FindTable(body, row.TableIndex, row.ScopeTag);
                        touchedTableIndices.Add(GlobalTableIndex(body, rowTable));
                        requiresStructuralChange = true;
                        AddMissingPreconditionRisk(op, patchSet, risks, i, forceError: true, "table-row operation");
                        break;
                    case WrapParagraphWithContentControlOperation wrap:
                        if (string.IsNullOrWhiteSpace(wrap.Tag)) throw new InvalidOperationException("tag is required");
                        if (FindContentControlsByTag(body, wrap.Tag).Count > 0) throw new InvalidOperationException($"content control tag already exists: {wrap.Tag}");
                        var wrapTarget = FindTargetParagraph(body, wrap, wrapMode: true);
                        TrackParagraph(body, wrapTarget, touchedParaIds, touchedParagraphIndices);
                        touchedTags.Add(wrap.Tag);
                        requiresStructuralChange = true;
                        AddMissingPreconditionRisk(op, patchSet, risks, i, forceError: true, "content-control wrapping operation");
                        break;
                    default:
                        risks.Add(new Risk("error", "unsupported_op", $"Unsupported op: {op.GetType().Name}", i));
                        break;
                }
            }
            catch (Exception ex)
            {
                risks.Add(new Risk("error", "target_resolution_failed", ex.Message, i));
            }
        }

        return new PatchAssessment(
            !risks.Any(r => r.Severity == "error"),
            patchSet.Operations.Count,
            requiresStructuralChange,
            new TouchedScope(
                touchedTags.Order().ToList(),
                touchedParaIds.Order().ToList(),
                touchedParagraphIndices.Order().ToList(),
                touchedTableIndices.Order().ToList(),
                touchedTableCells.Order().ToList()),
            risks);
    }

    public PatchAudit ApplyPatchSet(string sourceDocx, PatchSet patchSet, string outputDocx)
    {
        var sourceFullPath = Path.GetFullPath(sourceDocx);
        var outputFullPath = Path.GetFullPath(outputDocx);
        if (!string.IsNullOrWhiteSpace(patchSet.SourceSha256) &&
            !string.Equals(patchSet.SourceSha256, WordInspector.Sha256File(sourceFullPath), StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("patchset.source_sha256 does not match current DOCX; refusing stale write.");
        }

        var allowOverwrite = patchSet.Guard.AllowOverwrite || patchSet.Overwrite;
        if (string.Equals(sourceFullPath, outputFullPath, StringComparison.OrdinalIgnoreCase) && !allowOverwrite)
        {
            throw new InvalidOperationException("Refusing to overwrite the source DOCX. Use a distinct output path or set guard.allow_overwrite=true after backup.");
        }
        if (File.Exists(outputFullPath) && !allowOverwrite)
        {
            throw new IOException($"output path already exists: {outputFullPath}");
        }

        var assessment = AssessPatchSet(sourceFullPath, patchSet);
        if (!assessment.Ok)
        {
            throw new InvalidOperationException("PatchSet failed safety assessment: " + string.Join("; ", assessment.Risks.Select(r => $"{r.Code}: {r.Message}")));
        }

        Directory.CreateDirectory(Path.GetDirectoryName(outputFullPath)!);
        var candidate = outputFullPath + $".candidate.{Environment.ProcessId}";
        File.Delete(candidate);
        File.Copy(sourceFullPath, candidate, overwrite: true);

        var before = new WordInspector().Inspect(sourceFullPath);
        var applied = new List<Dictionary<string, object?>>();
        var validationOptions = BuildValidationOptions(assessment, patchSet);

        try
        {
            using (var doc = WordprocessingDocument.Open(candidate, true))
            {
                var body = doc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Missing document body.");
                foreach (var op in patchSet.Operations)
                {
                    ApplyOperation(doc, body, op, applied);
                }
                doc.MainDocumentPart!.Document.Save();
            }

            var report = new WordValidator().Validate(sourceFullPath, candidate, patchSet.Strict, validationOptions);
            var after = new WordInspector().Inspect(candidate);
            var audit = CreateAudit(sourceFullPath, outputFullPath, null, patchSet, assessment, applied, report, before, after, candidate, dryRun: false, keptOutput: true);

            if (patchSet.AbortOnValidationError && !report.Ok)
            {
                var invalidAuditPath = Path.ChangeExtension(outputFullPath, ".invalid.audit.json");
                if (patchSet.KeepInvalidOutput)
                {
                    var invalidDocx = Path.ChangeExtension(outputFullPath, ".invalid.docx");
                    File.Delete(invalidDocx);
                    File.Move(candidate, invalidDocx);
                }
                else
                {
                    File.Delete(candidate);
                }
                WriteJson(invalidAuditPath, audit with { AuditPath = invalidAuditPath });
                throw new InvalidOperationException($"Validation failed; final DOCX was not committed. See audit: {invalidAuditPath}");
            }

            File.Delete(outputFullPath);
            File.Move(candidate, outputFullPath);
            var finalReport = report with { TargetPath = outputFullPath };
            var finalAfter = new WordInspector().Inspect(outputFullPath);
            var auditPath = Path.ChangeExtension(outputFullPath, ".audit.json");
            var finalAudit = CreateAudit(sourceFullPath, outputFullPath, auditPath, patchSet, assessment, applied, finalReport, before, finalAfter, outputFullPath, dryRun: false, keptOutput: true);
            WriteJson(auditPath, finalAudit);
            return finalAudit;
        }
        catch
        {
            if (File.Exists(candidate)) File.Delete(candidate);
            throw;
        }
    }

    public PatchAudit DryRunPatchSet(string sourceDocx, PatchSet patchSet, bool keepOutput = false)
    {
        var dryDir = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(sourceDocx))!, ".wordai", "dryruns");
        Directory.CreateDirectory(dryDir);
        var outPath = Path.Combine(dryDir, $"{Path.GetFileNameWithoutExtension(sourceDocx)}.dryrun.{DateTimeOffset.UtcNow:yyyyMMddTHHmmssfffZ}.docx");
        var patchCopy = ClonePatchSet(patchSet);
        patchCopy.Guard.AllowOverwrite = true;
        patchCopy.Overwrite = true;
        var audit = ApplyPatchSet(sourceDocx, patchCopy, outPath) with { DryRun = true, KeptOutput = keepOutput };
        if (!keepOutput)
        {
            File.Delete(outPath);
            if (audit.AuditPath is not null) File.Delete(audit.AuditPath);
            audit = audit with { OutputPath = "", AuditPath = null };
        }
        return audit;
    }

    private static void ApplyOperation(WordprocessingDocument doc, Body body, PatchOperation op, List<Dictionary<string, object?>> applied)
    {
        switch (op)
        {
            case ReplaceContentControlTextOperation cc:
                var sdt = FindContentControlByTag(body, cc.Tag);
                var oldText = WordInspector.ContentControlText(sdt);
                AssertExpected(oldText, cc.ExpectedOldText, cc.ExpectedOldSha256, $"content control {cc.Tag}");
                ReplaceContentControlText(sdt, cc.Text, cc.AllowParagraphCountChange, cc.PreserveStyle, cc.AllowComplexContent);
                applied.Add(new Dictionary<string, object?> { ["op"] = "replace_content_control_text", ["tag"] = cc.Tag, ["old_sha256"] = WordInspector.Sha256String(oldText), ["new_sha256"] = WordInspector.Sha256String(cc.Text), ["path"] = WordInspector.GetOpenXmlPath(sdt) });
                break;
            case AppendContentControlTextOperation append:
                var appendSdt = FindContentControlByTag(body, append.Tag);
                var appendOld = WordInspector.ContentControlText(appendSdt);
                AssertExpected(appendOld, append.ExpectedOldText, append.ExpectedOldSha256, $"content control {append.Tag}");
                AppendOrPrependContentControlText(appendSdt, append.Text, append: true);
                applied.Add(new Dictionary<string, object?> { ["op"] = "append_content_control_text", ["tag"] = append.Tag, ["old_sha256"] = WordInspector.Sha256String(appendOld), ["inserted_sha256"] = WordInspector.Sha256String(append.Text), ["path"] = WordInspector.GetOpenXmlPath(appendSdt) });
                break;
            case PrependContentControlTextOperation prepend:
                var prependSdt = FindContentControlByTag(body, prepend.Tag);
                var prependOld = WordInspector.ContentControlText(prependSdt);
                AssertExpected(prependOld, prepend.ExpectedOldText, prepend.ExpectedOldSha256, $"content control {prepend.Tag}");
                AppendOrPrependContentControlText(prependSdt, prepend.Text, append: false);
                applied.Add(new Dictionary<string, object?> { ["op"] = "prepend_content_control_text", ["tag"] = prepend.Tag, ["old_sha256"] = WordInspector.Sha256String(prependOld), ["inserted_sha256"] = WordInspector.Sha256String(prepend.Text), ["path"] = WordInspector.GetOpenXmlPath(prependSdt) });
                break;
            case ReplaceTextInContentControlOperation replaceText:
                var textSdt = FindContentControlByTag(body, replaceText.Tag);
                var textOld = WordInspector.ContentControlText(textSdt);
                AssertExpected(textOld, replaceText.ExpectedOldText, replaceText.ExpectedOldSha256, $"content control {replaceText.Tag}");
                var count = ReplaceTextInContentControl(textSdt, replaceText.Find, replaceText.Replace, replaceText.Occurrence);
                if (count == 0 && replaceText.RequireMatch) throw new InvalidOperationException("find text not found");
                applied.Add(new Dictionary<string, object?> { ["op"] = "replace_text_in_content_control", ["tag"] = replaceText.Tag, ["replace_count"] = count, ["path"] = WordInspector.GetOpenXmlPath(textSdt) });
                break;
            case ReplaceParagraphTextOperation para:
                var p = FindTargetParagraph(body, para, wrapMode: false);
                var paraOld = WordInspector.ParagraphText(p);
                AssertExpected(paraOld, para.ExpectedOldText, para.ExpectedOldSha256, "paragraph");
                ReplaceParagraphText(p, para.Text, para.PreserveStyle, para.AllowComplexContent);
                applied.Add(ParagraphApplied("replace_paragraph_text", body, p, paraOld, para.Text));
                break;
            case InsertParagraphAfterOperation after:
                InsertParagraph(body, after, after.Text, insertAfter: true, after.InheritStyle, after.InheritHeadingStyle, applied);
                break;
            case InsertParagraphBeforeOperation before:
                InsertParagraph(body, before, before.Text, insertAfter: false, before.InheritStyle, before.InheritHeadingStyle, applied);
                break;
            case ReplaceTableCellTextOperation cell:
                var table = FindTable(body, cell.TableIndex, cell.ScopeTag);
                var globalTableIndex = GlobalTableIndex(body, table);
                var tableCell = FindCell(table, cell.Row, cell.Col);
                var cellOld = WordInspector.CellText(tableCell);
                AssertExpected(cellOld, cell.ExpectedOldText, cell.ExpectedOldSha256, "table cell");
                ReplaceCellText(tableCell, cell.Text, cell.PreserveStyle, cell.AllowParagraphCountChange, cell.AllowComplexContent);
                applied.Add(new Dictionary<string, object?> { ["op"] = "replace_table_cell_text", ["table_index"] = globalTableIndex, ["row"] = cell.Row, ["col"] = cell.Col, ["old_sha256"] = WordInspector.Sha256String(cellOld), ["new_sha256"] = WordInspector.Sha256String(cell.Text) });
                break;
            case AppendTableRowOperation row:
                var rowTable = FindTable(body, row.TableIndex, row.ScopeTag);
                var rowGlobalTableIndex = GlobalTableIndex(body, rowTable);
                var tableOld = WordInspector.TableText(rowTable);
                AssertExpected(tableOld, row.ExpectedOldText, row.ExpectedOldSha256, "table");
                AppendTableRow(rowTable, row);
                applied.Add(new Dictionary<string, object?> { ["op"] = "append_table_row", ["table_index"] = rowGlobalTableIndex, ["old_sha256"] = WordInspector.Sha256String(tableOld), ["appended_cells"] = row.Values.Count });
                break;
            case WrapParagraphWithContentControlOperation wrap:
                var target = FindTargetParagraph(body, wrap, wrapMode: true);
                var wrapOld = WordInspector.ParagraphText(target);
                AssertExpected(wrapOld, wrap.ExpectedOldText, wrap.ExpectedOldSha256, "paragraph");
                var wrapped = WrapParagraphWithContentControl(target, wrap.Tag, wrap.Alias ?? wrap.Title, wrap.Lock);
                applied.Add(new Dictionary<string, object?> { ["op"] = "wrap_paragraph_with_content_control", ["tag"] = wrap.Tag, ["alias"] = wrap.Alias ?? wrap.Title, ["old_sha256"] = WordInspector.Sha256String(wrapOld), ["path"] = WordInspector.GetOpenXmlPath(wrapped) });
                break;
            case AddCommentOperation comment:
                var commentTarget = FindTargetParagraph(body, comment, wrapMode: false);
                var commentOld = WordInspector.ParagraphText(commentTarget);
                AssertExpected(commentOld, comment.ExpectedOldText, comment.ExpectedOldSha256, "paragraph");
                var commentId = AddComment(doc, commentTarget, comment.Text, comment.Author, comment.Initials);
                applied.Add(new Dictionary<string, object?> { ["op"] = "add_comment", ["comment_id"] = commentId, ["old_sha256"] = WordInspector.Sha256String(commentOld), ["target_paragraph_index"] = ParagraphIndex(body, commentTarget), ["target_paraId"] = WordInspector.ParaId(commentTarget), ["path"] = WordInspector.GetOpenXmlPath(commentTarget) });
                break;
            default:
                throw new NotSupportedException($"Unsupported patch operation {op.GetType().Name}");
        }
    }

    private static ValidationOptions BuildValidationOptions(PatchAssessment assessment, PatchSet patchSet)
    {
        var options = new ValidationOptions();
        foreach (var tag in assessment.Touched.ContentControlTags) options.TouchedContentControlTags.Add(tag);
        foreach (var id in assessment.Touched.ParaIds) options.TouchedParaIds.Add(id);
        foreach (var idx in assessment.Touched.ParagraphIndices) options.TouchedParagraphIndices.Add(idx);
        foreach (var idx in assessment.Touched.TableIndices) options.TouchedTableIndices.Add(idx);
        foreach (var cell in assessment.Touched.TableCells) options.TouchedTableCells.Add(cell);

        foreach (var op in patchSet.Operations)
        {
            switch (op)
            {
                case ReplaceContentControlTextOperation cc when cc.AllowParagraphCountChange:
                    options.AllowedCountChanges.Add("paragraph_count");
                    options.AllowParagraphCountChange = true;
                    break;
                case AppendContentControlTextOperation:
                case PrependContentControlTextOperation:
                case InsertParagraphAfterOperation:
                case InsertParagraphBeforeOperation:
                    options.AllowedCountChanges.Add("paragraph_count");
                    options.AllowParagraphCountChange = true;
                    break;
                case ReplaceTableCellTextOperation cell when cell.AllowParagraphCountChange:
                    options.AllowedCountChanges.Add("paragraph_count");
                    options.AllowParagraphCountChange = true;
                    break;
                case AppendTableRowOperation:
                    options.AllowedCountChanges.Add("paragraph_count");
                    options.AllowParagraphCountChange = true;
                    options.AllowTableDimensionChange = true;
                    break;
                case WrapParagraphWithContentControlOperation wrap:
                    options.AllowedCountChanges.Add("content_control_count");
                    options.AllowedAddedContentControlTags.Add(wrap.Tag);
                    break;
                case AddCommentOperation:
                    options.AllowedPartChanges.Add("word/comments.xml");
                    options.AllowedPartChanges.Add("word/_rels/document.xml.rels");
                    options.AllowedPartChanges.Add("[Content_Types].xml");
                    options.AllowedCountChanges.Add("comment_count");
                    options.AllowedCountChanges.Add("comment_reference_count");
                    break;
            }
        }
        return options;
    }

    private static void ReplaceContentControlText(SdtElement sdt, string text, bool allowParagraphCountChange, bool preserveStyle, bool allowComplexContent)
    {
        if (!allowComplexContent && WordInspector.HasComplexContent(sdt))
        {
            throw new InvalidOperationException($"Refusing to replace complex content control {WordInspector.ContentControlTag(sdt)}.");
        }
        var content = SdtContent(sdt);
        var paragraphs = content.Elements<Paragraph>().ToList();
        if (paragraphs.Count == 0)
        {
            content.AppendChild(NewParagraph(text, null, null));
            return;
        }
        ReplaceParagraphsText(paragraphs, text, allowParagraphCountChange, preserveStyle, allowComplexContent);
    }

    private static void AppendOrPrependContentControlText(SdtElement sdt, string text, bool append)
    {
        var summary = WordInspector.ComplexSummary(sdt);
        if (summary["tables"] > 0 || summary["drawings"] > 0)
        {
            throw new InvalidOperationException("Appending/prepending text to a content control containing tables/images is blocked.");
        }
        var content = SdtContent(sdt);
        var paragraphs = content.Elements<Paragraph>().ToList();
        var template = paragraphs.LastOrDefault();
        var pPr = template?.ParagraphProperties?.CloneNode(true) as ParagraphProperties;
        var rPr = template?.Descendants<RunProperties>().FirstOrDefault()?.CloneNode(true) as RunProperties;
        var newParagraphs = SplitLines(text).Select(line => NewParagraph(line, pPr, rPr)).ToList();
        if (append)
        {
            foreach (var p in newParagraphs) content.AppendChild(p);
        }
        else
        {
            var first = content.Elements<OpenXmlElement>().FirstOrDefault();
            foreach (var p in newParagraphs)
            {
                if (first is null) content.AppendChild(p);
                else content.InsertBefore(p, first);
            }
        }
    }

    private static int ReplaceTextInContentControl(SdtElement sdt, string find, string replace, string occurrence)
    {
        var summary = WordInspector.ComplexSummary(sdt);
        if (summary["fields"] > 0 || summary["tracked_changes"] > 0)
        {
            throw new InvalidOperationException("Target content control contains fields or tracked changes; find/replace is blocked.");
        }
        var count = 0;
        foreach (var t in sdt.Descendants<Text>())
        {
            if (string.IsNullOrEmpty(t.Text) || !t.Text.Contains(find, StringComparison.Ordinal)) continue;
            var before = t.Text;
            if (string.Equals(occurrence, "first", StringComparison.OrdinalIgnoreCase))
            {
                t.Text = ReplaceFirst(t.Text, find, replace);
                count++;
                break;
            }
            count += CountOccurrences(before, find);
            t.Text = before.Replace(find, replace, StringComparison.Ordinal);
        }
        return count;
    }

    private static void ReplaceParagraphText(Paragraph p, string text, bool preserveStyle, bool allowComplexContent)
    {
        if (!allowComplexContent && WordInspector.HasComplexContent(p))
        {
            throw new InvalidOperationException($"Refusing to replace paragraph containing complex Word objects at {WordInspector.GetOpenXmlPath(p)}");
        }
        var pPr = preserveStyle ? p.ParagraphProperties?.CloneNode(true) as ParagraphProperties : null;
        var rPr = preserveStyle ? p.Descendants<RunProperties>().FirstOrDefault()?.CloneNode(true) as RunProperties : null;
        p.RemoveAllChildren();
        if (pPr is not null) p.AppendChild(pPr);
        p.AppendChild(NewRun(text, rPr));
    }

    private static void ReplaceParagraphsText(List<Paragraph> paragraphs, string text, bool allowParagraphCountChange, bool preserveStyle, bool allowComplexContent)
    {
        var lines = SplitLines(text).ToList();
        if (!allowParagraphCountChange && lines.Count != paragraphs.Count)
        {
            throw new InvalidOperationException($"Paragraph count would change from {paragraphs.Count} to {lines.Count}.");
        }
        foreach (var p in paragraphs)
        {
            if (!allowComplexContent && WordInspector.HasComplexContent(p))
            {
                throw new InvalidOperationException($"Refusing to replace paragraph containing complex Word objects at {WordInspector.GetOpenXmlPath(p)}");
            }
        }
        if (lines.Count == paragraphs.Count)
        {
            for (var i = 0; i < lines.Count; i++) ReplaceParagraphText(paragraphs[i], lines[i], preserveStyle, allowComplexContent);
            return;
        }
        var parent = paragraphs[0].Parent ?? throw new InvalidOperationException("Paragraph has no parent");
        var insertBefore = paragraphs[^1].NextSibling();
        var pPr = paragraphs[0].ParagraphProperties?.CloneNode(true) as ParagraphProperties;
        var rPr = preserveStyle ? paragraphs[0].Descendants<RunProperties>().FirstOrDefault()?.CloneNode(true) as RunProperties : null;
        foreach (var p in paragraphs) p.Remove();
        foreach (var line in lines)
        {
            var newP = NewParagraph(line, pPr, rPr);
            if (insertBefore is null) parent.AppendChild(newP);
            else parent.InsertBefore(newP, insertBefore);
        }
    }

    private static void InsertParagraph(Body body, ParagraphScopedOperation op, string text, bool insertAfter, bool inheritStyle, bool inheritHeadingStyle, List<Dictionary<string, object?>> applied)
    {
        var target = FindTargetParagraph(body, op, wrapMode: false);
        var old = WordInspector.ParagraphText(target);
        AssertExpected(old, op.ExpectedOldText, op.ExpectedOldSha256, "anchor paragraph");
        if (WordInspector.HeadingLevel(target.ParagraphProperties?.ParagraphStyleId?.Val?.Value) is not null && !inheritHeadingStyle)
        {
            inheritStyle = false;
        }
        var pPr = inheritStyle ? target.ParagraphProperties?.CloneNode(true) as ParagraphProperties : null;
        var rPr = inheritStyle ? target.Descendants<RunProperties>().FirstOrDefault()?.CloneNode(true) as RunProperties : null;
        var parent = target.Parent ?? throw new InvalidOperationException("Paragraph has no parent");
        var reference = insertAfter ? target.NextSibling() : target;
        var insertedCount = 0;
        foreach (var line in SplitLines(text))
        {
            var p = NewParagraph(line, pPr, rPr);
            if (reference is null) parent.AppendChild(p);
            else parent.InsertBefore(p, reference);
            insertedCount++;
        }
        applied.Add(new Dictionary<string, object?> { ["op"] = insertAfter ? "insert_paragraph_after" : "insert_paragraph_before", ["anchor_paraId"] = WordInspector.ParaId(target), ["anchor_paragraph_index"] = ParagraphIndex(body, target), ["old_sha256"] = WordInspector.Sha256String(old), ["inserted_paragraph_count"] = insertedCount, ["inserted_sha256"] = WordInspector.Sha256String(text) });
    }

    private static void ReplaceCellText(TableCell cell, string text, bool preserveStyle, bool allowParagraphCountChange, bool allowComplexContent)
    {
        var paragraphs = cell.Elements<Paragraph>().ToList();
        if (paragraphs.Count == 0)
        {
            cell.AppendChild(NewParagraph(text, null, null));
            return;
        }
        ReplaceParagraphsText(paragraphs, text, allowParagraphCountChange, preserveStyle, allowComplexContent);
    }

    private static void AppendTableRow(Table table, AppendTableRowOperation op)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0) throw new InvalidOperationException("Cannot append row to empty table");
        var templateIndex = (op.TemplateRow ?? rows.Count) - 1;
        if (templateIndex < 0 || templateIndex >= rows.Count) throw new InvalidOperationException("template_row out of range");
        var newRow = (TableRow)rows[templateIndex].CloneNode(true);
        var cells = newRow.Elements<TableCell>().ToList();
        if (op.Values.Count > cells.Count) throw new InvalidOperationException("values has more items than table columns");
        for (var i = 0; i < op.Values.Count; i++)
        {
            ReplaceCellText(cells[i], op.Values[i], preserveStyle: true, allowParagraphCountChange: true, allowComplexContent: true);
        }
        table.AppendChild(newRow);
    }

    private static SdtBlock WrapParagraphWithContentControl(Paragraph paragraph, string tag, string? alias, bool locked)
    {
        if (WordInspector.AncestorContentControlTag(paragraph) is not null)
        {
            throw new InvalidOperationException("Paragraph is already inside a content control");
        }
        var parent = paragraph.Parent ?? throw new InvalidOperationException("Paragraph has no parent");
        var reference = paragraph.NextSibling();
        var sdt = new SdtBlock();
        var pr = new SdtProperties(
            new SdtAlias { Val = alias ?? tag },
            new Tag { Val = tag },
            new SdtId { Val = Math.Abs(tag.GetHashCode()) });
        if (locked)
        {
            pr.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Lock { Val = LockingValues.SdtLocked });
        }
        sdt.AppendChild(pr);
        var content = new SdtContentBlock();
        paragraph.Remove();
        content.AppendChild(paragraph);
        sdt.AppendChild(content);
        if (reference is null) parent.AppendChild(sdt);
        else parent.InsertBefore(sdt, reference);
        return sdt;
    }

    private static int AddComment(WordprocessingDocument doc, Paragraph target, string text, string author, string initials)
    {
        var main = doc.MainDocumentPart ?? throw new InvalidOperationException("Missing main document part");
        var commentsPart = main.WordprocessingCommentsPart ?? main.AddNewPart<WordprocessingCommentsPart>();
        commentsPart.Comments ??= new Comments();
        var comments = commentsPart.Comments;
        var id = comments.Elements<Comment>().Select(c => int.TryParse(c.Id?.Value, out var n) ? n : -1).DefaultIfEmpty(-1).Max() + 1;

        var start = new CommentRangeStart { Id = id.ToString() };
        var end = new CommentRangeEnd { Id = id.ToString() };
        var refRun = new Run(
            new RunProperties(new RunStyle { Val = "CommentReference" }),
            new CommentReference { Id = id.ToString() });
        var pPr = target.GetFirstChild<ParagraphProperties>();
        target.InsertAfter(start, pPr);
        target.AppendChild(end);
        target.AppendChild(refRun);

        var comment = new Comment
        {
            Id = id.ToString(),
            Author = author,
            Initials = initials,
            Date = DateTime.UtcNow,
        };
        comment.AppendChild(new Paragraph(new Run(new Text(text))));
        comments.AppendChild(comment);
        comments.Save();
        return id;
    }

    private static SdtElement FindContentControlByTag(Body body, string tag)
    {
        var found = FindContentControlsByTag(body, tag);
        return found.Count switch
        {
            0 => throw new InvalidOperationException($"Content control tag not found: {tag}"),
            1 => found[0],
            _ => throw new InvalidOperationException($"Duplicate content control tag: {tag}"),
        };
    }

    private static List<SdtElement> FindContentControlsByTag(Body body, string tag)
        => body.Descendants<SdtElement>().Where(s => WordInspector.ContentControlTag(s) == tag).ToList();

    private static OpenXmlCompositeElement SdtContent(SdtElement sdt)
        => sdt switch
        {
            SdtBlock block => block.SdtContentBlock ?? throw new InvalidOperationException("Missing block content control content"),
            SdtRun run => run.SdtContentRun ?? throw new InvalidOperationException("Missing run content control content"),
            SdtCell cell => cell.SdtContentCell ?? throw new InvalidOperationException("Missing cell content control content"),
            _ => throw new InvalidOperationException("Unsupported content control shape"),
        };

    private static Paragraph FindTargetParagraph(Body body, ParagraphScopedOperation op, bool wrapMode)
    {
        var tag = wrapMode ? op.ContentControlTag ?? op.TargetTag : op.Tag ?? op.ContentControlTag ?? op.TargetTag;
        if (!string.IsNullOrWhiteSpace(tag))
        {
            var sdt = FindContentControlByTag(body, tag);
            return sdt.Descendants<Paragraph>().FirstOrDefault() ?? throw new InvalidOperationException($"Content control has no paragraph: {tag}");
        }
        if (!string.IsNullOrWhiteSpace(op.ParaId))
        {
            var p = body.Descendants<Paragraph>().FirstOrDefault(x => WordInspector.ParaId(x) == op.ParaId);
            if (p is not null) return p;
        }
        if (op.ParagraphIndex is { } index)
        {
            var paragraphs = body.Descendants<Paragraph>().ToList();
            if (index >= 1 && index <= paragraphs.Count) return paragraphs[index - 1];
        }
        throw new InvalidOperationException("Target paragraph not found");
    }

    private static Table FindTable(Body body, int tableIndex, string? scopeTag)
    {
        var root = string.IsNullOrWhiteSpace(scopeTag) ? (OpenXmlElement)body : FindContentControlByTag(body, scopeTag);
        var tables = root.Descendants<Table>().ToList();
        if (tableIndex < 1 || tableIndex > tables.Count) throw new InvalidOperationException("table_index out of range");
        return tables[tableIndex - 1];
    }

    private static TableCell FindCell(Table table, int row, int col)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (row < 1 || row > rows.Count) throw new InvalidOperationException("row out of range");
        var cells = rows[row - 1].Elements<TableCell>().ToList();
        if (col < 1 || col > cells.Count) throw new InvalidOperationException("col out of range");
        return cells[col - 1];
    }

    private static int GlobalTableIndex(Body body, Table table)
    {
        var tables = body.Descendants<Table>().ToList();
        var idx = tables.IndexOf(table);
        if (idx < 0) throw new InvalidOperationException("table is not inside body");
        return idx + 1;
    }

    private static int? ParagraphIndex(Body body, Paragraph paragraph)
    {
        var paragraphs = body.Descendants<Paragraph>().ToList();
        var idx = paragraphs.IndexOf(paragraph);
        return idx >= 0 ? idx + 1 : null;
    }

    private static void TrackParagraph(Body body, Paragraph paragraph, HashSet<string> paraIds, HashSet<int> paragraphIndices)
    {
        if (WordInspector.ParaId(paragraph) is { Length: > 0 } paraId) paraIds.Add(paraId);
        if (ParagraphIndex(body, paragraph) is { } idx) paragraphIndices.Add(idx);
    }

    private static Dictionary<string, object?> ParagraphApplied(string op, Body body, Paragraph paragraph, string oldText, string newText) => new()
    {
        ["op"] = op,
        ["path"] = WordInspector.GetOpenXmlPath(paragraph),
        ["paraId"] = WordInspector.ParaId(paragraph),
        ["paragraph_index"] = ParagraphIndex(body, paragraph),
        ["old_sha256"] = WordInspector.Sha256String(oldText),
        ["new_sha256"] = WordInspector.Sha256String(newText),
    };

    private static void AssertExpected(string actual, string? expectedText, string? expectedSha256, string label)
    {
        if (expectedText is not null && actual != expectedText) throw new InvalidOperationException($"{label}: expected_old_text does not match current text");
        if (expectedSha256 is not null && WordInspector.Sha256String(actual) != expectedSha256) throw new InvalidOperationException($"{label}: expected_old_sha256 does not match current text");
    }

    private static void AddMissingPreconditionRisk(PatchOperation op, PatchSet patchSet, List<Risk> risks, int index, bool forceError, string label)
    {
        if (op.ExpectedOldText is not null || op.ExpectedOldSha256 is not null) return;
        var severity = patchSet.Guard.RequirePreconditions || forceError || HighRiskOperationTypes.Contains(op.GetType()) ? "error" : "warning";
        risks.Add(new Risk(severity, "missing_precondition", $"No expected_old_text/expected_old_sha256 supplied for {label}.", index));
    }

    private static Run NewRun(string text, RunProperties? rPr)
    {
        var run = new Run();
        if (rPr is not null) run.AppendChild((RunProperties)rPr.CloneNode(true));
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return run;
    }

    private static Paragraph NewParagraph(string text, ParagraphProperties? pPr, RunProperties? rPr)
    {
        var p = new Paragraph();
        if (pPr is not null) p.AppendChild((ParagraphProperties)pPr.CloneNode(true));
        p.AppendChild(NewRun(text, rPr));
        return p;
    }

    private static IEnumerable<string> SplitLines(string text)
        => text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').Split('\n');

    private static string ReplaceFirst(string source, string find, string replace)
    {
        var index = source.IndexOf(find, StringComparison.Ordinal);
        return index < 0 ? source : string.Concat(source.AsSpan(0, index), replace, source.AsSpan(index + find.Length));
    }

    private static int CountOccurrences(string source, string find)
    {
        if (find.Length == 0) return 0;
        var count = 0;
        var index = 0;
        while ((index = source.IndexOf(find, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += find.Length;
        }
        return count;
    }

    private static PatchAudit CreateAudit(
        string source,
        string output,
        string? auditPath,
        PatchSet patchSet,
        PatchAssessment assessment,
        IReadOnlyList<Dictionary<string, object?>> applied,
        ValidationReport validation,
        DocumentProfile before,
        DocumentProfile after,
        string diffTarget,
        bool dryRun,
        bool keptOutput)
        => new(
            source,
            output,
            auditPath,
            patchSet.Reason,
            assessment,
            applied,
            validation,
            Metrics(before),
            Metrics(after),
            TextDiff(source, diffTarget),
            dryRun,
            keptOutput,
            DateTimeOffset.UtcNow.ToString("O"));

    private static Dictionary<string, int> Metrics(DocumentProfile profile) => new()
    {
        ["paragraph_count"] = profile.ParagraphCount,
        ["table_count"] = profile.TableCount,
        ["image_count"] = profile.ImageCount,
        ["field_count"] = profile.FieldCount,
        ["comment_count"] = profile.CommentCount,
        ["comment_reference_count"] = profile.CommentReferenceCount,
        ["tracked_change_count"] = profile.TrackedChangeCount,
        ["content_control_count"] = profile.ContentControlCount,
        ["heading_count"] = profile.HeadingCount,
    };

    private static string TextDiff(string source, string target)
    {
        var a = WordInspector.ExtractPlainText(source).Split('\n');
        var b = WordInspector.ExtractPlainText(target).Split('\n');
        var lines = new List<string> { $"--- {source}", $"+++ {target}" };
        var max = Math.Max(a.Length, b.Length);
        for (var i = 0; i < max; i++)
        {
            var left = i < a.Length ? a[i] : null;
            var right = i < b.Length ? b[i] : null;
            if (left == right) continue;
            if (left is not null) lines.Add("-" + left);
            if (right is not null) lines.Add("+" + right);
        }
        return string.Join("\n", lines);
    }

    private static void WriteJson(string path, object value)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(value, JsonSupport.Options));
    }

    private static PatchSet ClonePatchSet(PatchSet patchSet)
        => System.Text.Json.JsonSerializer.Deserialize<PatchSet>(System.Text.Json.JsonSerializer.Serialize(patchSet, JsonSupport.Options), JsonSupport.Options)
           ?? throw new InvalidOperationException("Unable to clone PatchSet");
}
