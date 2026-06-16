using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordAi.OpenXml;

public sealed class WordInspector
{
    private const string W14Ns = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly Regex HeadingNamePattern = new(@"^(?:heading|标题)([1-9])$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public sealed class ParagraphStyleInfo
    {
        public required string StyleId { get; init; }
        public string? Name { get; init; }
        public string? BasedOn { get; init; }
        public string? OutlineLevel { get; init; }
        public int? HeadingLevel { get; set; }
        public bool IsToc { get; set; }
    }

    public DocumentProfile Inspect(string docxPath, int previewLength = 240)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        var main = doc.MainDocumentPart ?? throw new InvalidOperationException("Missing main document part.");
        var body = main.Document?.Body ?? throw new InvalidOperationException("Missing document body.");
        var anchors = ListAnchors(docxPath, previewLength);

        return new DocumentProfile(
            docxPath,
            Sha256File(docxPath),
            new FileInfo(docxPath).Length,
            CountZipParts(docxPath),
            body.Descendants<Paragraph>().Count(),
            body.Descendants<Table>().Count(),
            main.Document.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Count(),
            main.Document.Descendants<FieldCode>().Count() + main.Document.Descendants<SimpleField>().Count(),
            main.WordprocessingCommentsPart?.Comments?.Elements<Comment>().Count() ?? 0,
            main.Document.Descendants<CommentRangeStart>().Count() + main.Document.Descendants<CommentReference>().Count(),
            CountTrackedChanges(main.Document),
            body.Descendants<SdtElement>().Count(),
            anchors.Count(a => a.Kind == "heading"),
            body.Descendants<BookmarkStart>().Count(b => !string.IsNullOrWhiteSpace(b.Name?.Value) && !b.Name!.Value.StartsWith('_')),
            anchors);
    }

    public IReadOnlyList<Anchor> ListAnchors(string docxPath, int previewLength = 240)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        var main = doc.MainDocumentPart ?? throw new InvalidOperationException("Missing main document part.");
        var body = main.Document?.Body ?? throw new InvalidOperationException("Missing document body.");
        var styles = LoadParagraphStyles(main);
        var tocParagraphs = TocParagraphIndices(body, styles);
        var anchors = new List<Anchor>();

        foreach (var sdt in body.Descendants<SdtElement>())
        {
            var tag = ContentControlTag(sdt);
            var alias = ContentControlAlias(sdt);
            var id = ContentControlId(sdt);
            var text = ContentControlText(sdt);
            anchors.Add(new Anchor(
                $"cc:{tag ?? id ?? Sha256String(GetOpenXmlPath(sdt))[..12]}",
                "content_control",
                alias ?? tag ?? id ?? "content_control",
                GetOpenXmlPath(sdt),
                Truncate(text, previewLength),
                null,
                null,
                new Dictionary<string, object?>
                {
                    ["tag"] = tag,
                    ["alias"] = alias,
                    ["id"] = id,
                    ["text_sha256"] = Sha256String(text),
                    ["xml_sha256"] = Sha256String(sdt.OuterXml),
                    ["complexity"] = ComplexSummary(sdt),
                }));
        }

        var outline = new List<string>();
        var index = 0;
        foreach (var p in body.Descendants<Paragraph>())
        {
            index++;
            var text = ParagraphText(p).Trim();
            var style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            if (tocParagraphs.Contains(index)) continue;
            var level = HeadingLevel(p, styles);
            if (level is not null && text.Length > 0)
            {
                while (outline.Count >= level.Value) outline.RemoveAt(outline.Count - 1);
                outline.Add(text);
                anchors.Add(new Anchor(
                    $"heading:{index}",
                    "heading",
                    string.Join(" > ", outline),
                    GetOpenXmlPath(p),
                    Truncate(text, previewLength),
                    style,
                    level,
                    new Dictionary<string, object?>
                    {
                        ["paragraph_index"] = index,
                        ["paraId"] = ParaId(p),
                        ["style_name"] = styles.TryGetValue(style ?? "", out var styleInfo) ? styleInfo.Name : null,
                        ["is_toc"] = false,
                        ["text_sha256"] = Sha256String(text),
                    }));
            }
        }

        foreach (var bm in body.Descendants<BookmarkStart>())
        {
            var name = bm.Name?.Value;
            if (!string.IsNullOrWhiteSpace(name) && !name.StartsWith('_'))
            {
                anchors.Add(new Anchor(
                    $"bookmark:{name}",
                    "bookmark",
                    name,
                    GetOpenXmlPath(bm),
                    null,
                    null,
                    null,
                    new Dictionary<string, object?> { ["name"] = name, ["id"] = bm.Id?.Value }));
            }
        }

        return anchors;
    }

    public static string ParagraphText(Paragraph p) => string.Concat(p.Descendants<Text>().Select(t => t.Text));

    public static string ElementText(OpenXmlElement element)
        => string.Join("\n", element.Descendants<Paragraph>().Select(ParagraphText).Where(x => x.Length > 0));

    public static string ContentControlText(SdtElement sdt)
    {
        var paragraphs = sdt.Descendants<Paragraph>().Select(ParagraphText).ToList();
        return paragraphs.Count > 0 ? string.Join("\n", paragraphs) : string.Concat(sdt.Descendants<Text>().Select(t => t.Text));
    }

    public static string TableText(Table table)
        => string.Join("\n", table.Elements<TableRow>().Select(r => string.Join("\t", r.Elements<TableCell>().Select(CellText))));

    public static string CellText(TableCell cell)
        => string.Join("\n", cell.Elements<Paragraph>().Select(ParagraphText));

    public static string? ContentControlTag(SdtElement sdt) => sdt.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;

    public static string? ContentControlAlias(SdtElement sdt) => sdt.SdtProperties?.GetFirstChild<SdtAlias>()?.Val?.Value;

    public static string? ContentControlId(SdtElement sdt) => sdt.SdtProperties?.GetFirstChild<SdtId>()?.Val?.Value.ToString();

    public static string? ParaId(Paragraph p)
    {
        var attr = p.GetAttributes().FirstOrDefault(a => a.LocalName == "paraId" && a.NamespaceUri == W14Ns);
        return string.IsNullOrEmpty(attr.Value) ? null : attr.Value;
    }

    public static IReadOnlyDictionary<string, ParagraphStyleInfo> LoadParagraphStyles(MainDocumentPart main)
    {
        var result = new Dictionary<string, ParagraphStyleInfo>(StringComparer.Ordinal);
        var styles = main.StyleDefinitionsPart?.Styles;
        if (styles is null) return result;

        foreach (var style in styles.Elements<Style>().Where(s => s.Type?.Value == StyleValues.Paragraph))
        {
            var styleId = style.StyleId?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) continue;
            result[styleId] = new ParagraphStyleInfo
            {
                StyleId = styleId,
                Name = style.StyleName?.Val?.Value,
                BasedOn = style.BasedOn?.Val?.Value,
                OutlineLevel = style.StyleParagraphProperties?.OutlineLevel?.Val?.Value.ToString(),
            };
        }

        foreach (var info in result.Values)
        {
            info.IsToc = IsTocStyle(info.StyleId, result);
            info.HeadingLevel = HeadingLevel(info.StyleId, result);
        }
        return result;
    }

    public static IReadOnlyDictionary<string, ParagraphStyleInfo> LoadParagraphStyles(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        var main = doc.MainDocumentPart ?? throw new InvalidOperationException("Missing main document part.");
        return LoadParagraphStyles(main);
    }

    public static int? HeadingLevel(Paragraph paragraph, IReadOnlyDictionary<string, ParagraphStyleInfo>? styles = null)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (IsTocStyle(styleId, styles)) return null;
        var direct = OutlineLevelValue(paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value.ToString());
        return direct ?? HeadingLevel(styleId, styles);
    }

    public static int? HeadingLevel(string? styleId, IReadOnlyDictionary<string, ParagraphStyleInfo>? styles = null)
    {
        if (string.IsNullOrWhiteSpace(styleId)) return null;
        if (IsTocStyle(styleId, styles)) return null;
        var direct = HeadingLevelFromName(styleId);
        if (direct is not null) return direct;
        if (styles is not null && styles.TryGetValue(styleId, out var info))
        {
            var nameLevel = HeadingLevelFromName(info.Name);
            if (nameLevel is not null) return nameLevel;
            var outlineLevel = OutlineLevelValue(info.OutlineLevel);
            if (outlineLevel is not null) return outlineLevel;
        }
        return null;
    }

    public static bool IsTocStyle(string? styleId, IReadOnlyDictionary<string, ParagraphStyleInfo>? styles = null)
    {
        var values = new List<string>();
        if (!string.IsNullOrWhiteSpace(styleId)) values.Add(styleId);
        if (!string.IsNullOrWhiteSpace(styleId) && styles is not null && styles.TryGetValue(styleId, out var info) && !string.IsNullOrWhiteSpace(info.Name))
        {
            values.Add(info.Name);
        }
        foreach (var value in values)
        {
            var normalized = Regex.Replace(value.Trim(), @"\s+", "").ToLowerInvariant();
            if (normalized.Length == 0) continue;
            if (Regex.IsMatch(normalized, @"^toc\d*$", RegexOptions.IgnoreCase)) return true;
            if (normalized == "tocheading" || normalized == "tableofcontents") return true;
            if (normalized.StartsWith("toc", StringComparison.OrdinalIgnoreCase) && (normalized.Contains("heading", StringComparison.OrdinalIgnoreCase) || normalized[3..].All(char.IsDigit))) return true;
            if (normalized.StartsWith("目录", StringComparison.OrdinalIgnoreCase) || normalized.StartsWith("wpsoffice手动目录", StringComparison.OrdinalIgnoreCase)) return true;
            if (normalized.StartsWith("tableofcontents", StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static int? HeadingLevelFromName(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var normalized = Regex.Replace(value.Trim(), @"\s+", "");
        var match = HeadingNamePattern.Match(normalized);
        if (!match.Success) return null;
        return int.TryParse(match.Groups[1].Value, out var level) ? level : null;
    }

    private static int? OutlineLevelValue(string? value)
    {
        if (!int.TryParse(value, out var raw)) return null;
        return raw is >= 0 and <= 8 ? raw + 1 : null;
    }

    private static HashSet<int> TocParagraphIndices(Body body, IReadOnlyDictionary<string, ParagraphStyleInfo> styles)
    {
        var result = new HashSet<int>();
        var fieldStack = new Stack<FieldContext>();
        var index = 0;

        bool ActiveTocField() => fieldStack.Any(ctx => ctx.IsToc);
        void ClearLeakedTocFields()
        {
            fieldStack = new Stack<FieldContext>(fieldStack.Reverse().Where(ctx => !ctx.IsToc));
        }

        foreach (var p in body.Descendants<Paragraph>())
        {
            index++;
            var style = p.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
            var explicitToc = IsTocStyle(style, styles) || IsTocSdtDescendant(p) || HasTocFieldInstruction(p) || HasTocReferenceField(p);
            if (ActiveTocField() && !explicitToc && !string.IsNullOrWhiteSpace(ParagraphText(p)))
            {
                // Some DOCX producers leave the TOC complex field unclosed. Do not
                // let a leaked TOC field consume the body outline after the visible TOC block.
                ClearLeakedTocFields();
            }
            var isToc = ActiveTocField() || explicitToc;
            if (p.Descendants<SimpleField>().Any(f => IsTocFieldInstruction(f.Instruction?.Value)))
            {
                isToc = true;
            }
            foreach (var el in p.Descendants<OpenXmlElement>())
            {
                if (el is FieldChar fld)
                {
                    var kind = fld.FieldCharType?.Value;
                    if (kind == FieldCharValues.Begin)
                    {
                        fieldStack.Push(new FieldContext());
                    }
                    else if (kind == FieldCharValues.Separate)
                    {
                        if (ActiveTocField()) isToc = true;
                    }
                    else if (kind == FieldCharValues.End)
                    {
                        if (ActiveTocField()) isToc = true;
                        if (fieldStack.Count > 0) fieldStack.Pop();
                    }
                }
                else if (el is FieldCode code && fieldStack.Count > 0)
                {
                    var ctx = fieldStack.Pop();
                    ctx.Instruction += code.Text ?? "";
                    if (IsTocFieldInstruction(ctx.Instruction))
                    {
                        ctx.IsToc = true;
                        isToc = true;
                    }
                    fieldStack.Push(ctx);
                }
            }
            if (ActiveTocField()) isToc = true;
            if (isToc) result.Add(index);
        }

        return result;
    }

    private sealed class FieldContext
    {
        public string Instruction { get; set; } = "";
        public bool IsToc { get; set; }
    }

    private static bool IsTocFieldInstruction(string? instr)
        => !string.IsNullOrWhiteSpace(instr) && Regex.IsMatch(instr, @"^\s*TOC(?:\s|\\|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static bool HasTocFieldInstruction(Paragraph paragraph)
        => paragraph.Descendants<SimpleField>().Any(f => IsTocFieldInstruction(f.Instruction?.Value))
           || paragraph.Descendants<FieldCode>().Any(code => IsTocFieldInstruction(code.Text));

    private static bool HasTocReferenceField(Paragraph paragraph)
    {
        static bool Match(string? value)
            => !string.IsNullOrWhiteSpace(value)
               && Regex.IsMatch(value, @"\b(?:PAGEREF|HYPERLINK)\b.*_Toc", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        return paragraph.Descendants<SimpleField>().Any(f => Match(f.Instruction?.Value))
               || paragraph.Descendants<FieldCode>().Any(code => Match(code.Text));
    }

    private static bool IsTocSdtDescendant(OpenXmlElement element)
    {
        foreach (var sdt in element.Ancestors<SdtElement>())
        {
            var xml = sdt.SdtProperties?.OuterXml?.ToLowerInvariant() ?? "";
            if (xml.Contains("table of contents", StringComparison.OrdinalIgnoreCase)
                || xml.Contains("tableofcontents", StringComparison.OrdinalIgnoreCase)
                || xml.Contains("目录", StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(xml, @"(^|[^a-z])toc([^a-z]|$)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant))
            {
                return true;
            }
        }
        return false;
    }

    public static bool HasComplexContent(OpenXmlElement element)
        => element.Descendants<Table>().Any()
           || element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any()
           || element.Descendants<Picture>().Any()
           || element.Descendants<SimpleField>().Any()
           || element.Descendants<FieldCode>().Any()
           || element.Descendants<CommentRangeStart>().Any()
           || element.Descendants<CommentRangeEnd>().Any()
           || element.Descendants<CommentReference>().Any()
           || CountTrackedChanges(element) > 0;

    public static Dictionary<string, int> ComplexSummary(OpenXmlElement element) => new()
    {
        ["tables"] = element.Descendants<Table>().Count(),
        ["drawings"] = element.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Count() + element.Descendants<Picture>().Count(),
        ["fields"] = element.Descendants<SimpleField>().Count() + element.Descendants<FieldCode>().Count(),
        ["comments"] = element.Descendants<CommentRangeStart>().Count() + element.Descendants<CommentReference>().Count(),
        ["tracked_changes"] = CountTrackedChanges(element),
    };

    public static int CountTrackedChanges(OpenXmlElement element)
        => element.Descendants<InsertedRun>().Count()
           + element.Descendants<DeletedRun>().Count()
           + element.Descendants<MoveFrom>().Count()
           + element.Descendants<MoveTo>().Count();

    public static string? AncestorContentControlTag(OpenXmlElement element)
    {
        for (var current = element.Parent; current is not null; current = current.Parent)
        {
            if (current is SdtElement sdt) return ContentControlTag(sdt);
        }
        return null;
    }

    public static int? AncestorTableIndex(OpenXmlElement element, Body body)
    {
        var table = element.Ancestors<Table>().FirstOrDefault();
        if (table is null) return null;
        var tables = body.Descendants<Table>().ToList();
        var idx = tables.IndexOf(table);
        return idx >= 0 ? idx + 1 : null;
    }

    public static string GetOpenXmlPath(OpenXmlElement element)
    {
        var stack = new Stack<string>();
        for (var current = element; current is not null; current = current.Parent)
        {
            var parent = current.Parent;
            var name = current.GetType().Name;
            if (parent is not null)
            {
                var same = parent.ChildElements.Where(c => c.GetType() == current.GetType()).ToList();
                stack.Push($"{name}[{same.IndexOf(current) + 1}]");
            }
            else
            {
                stack.Push(name);
            }
        }
        return "/" + string.Join("/", stack);
    }

    public static string Sha256String(string value)
        => Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();

    public static string Sha256File(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    public static Dictionary<string, string> HashParts(string path)
    {
        using var zip = ZipFile.OpenRead(path);
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var e in zip.Entries.Where(x => !string.IsNullOrEmpty(x.Name)))
        {
            using var s = e.Open();
            result[e.FullName!] = Convert.ToHexString(SHA256.HashData(s)).ToLowerInvariant();
        }
        return result;
    }

    public static string ExtractPlainText(string docxPath)
    {
        using var doc = WordprocessingDocument.Open(docxPath, false);
        var body = doc.MainDocumentPart?.Document?.Body ?? throw new InvalidOperationException("Missing document body.");
        return string.Join("\n", body.Descendants<Paragraph>().Select(ParagraphText));
    }

    private static int CountZipParts(string docxPath)
    {
        using var zip = ZipFile.OpenRead(docxPath);
        return zip.Entries.Count(e => !string.IsNullOrEmpty(e.Name));
    }

    private static string Truncate(string? value, int max) => string.IsNullOrEmpty(value) ? string.Empty : value.Length <= max ? value : value[..max];
}
