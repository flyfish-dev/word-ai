using System.Text.Json;
using WordAi.OpenXml;

if (args.Length < 2)
{
    PrintUsage();
    return 2;
}

try
{
    switch (args[0])
    {
        case "inspect":
            Console.WriteLine(JsonSerializer.Serialize(new WordInspector().Inspect(args[1]), JsonSupport.Options));
            return 0;

        case "assess" when args.Length >= 3:
        {
            var patch = ReadPatchSet(args[2]);
            Console.WriteLine(JsonSerializer.Serialize(new WordPatchApplier().AssessPatchSet(args[1], patch), JsonSupport.Options));
            return 0;
        }

        case "dry-run" when args.Length >= 3:
        {
            var patch = ReadPatchSet(args[2]);
            var keep = args.Length >= 4 && bool.TryParse(args[3], out var parsed) && parsed;
            Console.WriteLine(JsonSerializer.Serialize(new WordPatchApplier().DryRunPatchSet(args[1], patch, keep), JsonSupport.Options));
            return 0;
        }

        case "apply" when args.Length >= 4:
        {
            var patch = ReadPatchSet(args[2]);
            Console.WriteLine(JsonSerializer.Serialize(new WordPatchApplier().ApplyPatchSet(args[1], patch, args[3]), JsonSupport.Options));
            return 0;
        }

        case "validate" when args.Length >= 3:
        {
            var validation = ReadValidationOptions(args.Length >= 4 ? args[3] : null);
            Console.WriteLine(JsonSerializer.Serialize(new WordValidator().Validate(args[1], args[2], validation.Strict, validation.Options), JsonSupport.Options));
            return 0;
        }

        default:
            PrintUsage();
            return 2;
    }
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex);
    return 1;
}

static PatchSet ReadPatchSet(string path)
    => JsonSerializer.Deserialize<PatchSet>(File.ReadAllText(path), JsonSupport.Options)
       ?? throw new InvalidOperationException("Invalid patchset JSON.");

static (bool Strict, ValidationOptions Options) ReadValidationOptions(string? path)
{
    var options = new ValidationOptions();
    var strict = true;
    if (string.IsNullOrWhiteSpace(path))
    {
        return (strict, options);
    }

    using var doc = JsonDocument.Parse(File.ReadAllText(path));
    var root = doc.RootElement;
    if (root.TryGetProperty("strict", out var strictEl) && strictEl.ValueKind is JsonValueKind.True or JsonValueKind.False)
    {
        strict = strictEl.GetBoolean();
    }

    AddStrings(root, "allowed_part_changes", options.AllowedPartChanges);
    AddStrings(root, "allowed_count_changes", options.AllowedCountChanges);
    AddStrings(root, "touched_content_control_tags", options.TouchedContentControlTags);
    AddStrings(root, "touched_para_ids", options.TouchedParaIds);
    AddInts(root, "touched_paragraph_indices", options.TouchedParagraphIndices);
    AddInts(root, "touched_table_indices", options.TouchedTableIndices);
    AddTableCells(root, "touched_table_cells", options.TouchedTableCells);
    AddStrings(root, "allowed_added_content_control_tags", options.AllowedAddedContentControlTags);

    if (root.TryGetProperty("allow_table_dimension_change", out var tableDimEl) && tableDimEl.ValueKind is JsonValueKind.True or JsonValueKind.False)
    {
        options.AllowTableDimensionChange = tableDimEl.GetBoolean();
    }
    if (root.TryGetProperty("allow_paragraph_count_change", out var paraCountEl) && paraCountEl.ValueKind is JsonValueKind.True or JsonValueKind.False)
    {
        options.AllowParagraphCountChange = paraCountEl.GetBoolean();
    }

    return (strict, options);
}

static void AddStrings(JsonElement root, string property, ISet<string> target)
{
    if (!root.TryGetProperty(property, out var el) || el.ValueKind != JsonValueKind.Array)
    {
        return;
    }
    foreach (var item in el.EnumerateArray())
    {
        if (item.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(item.GetString()))
        {
            target.Add(item.GetString()!);
        }
    }
}

static void AddInts(JsonElement root, string property, ISet<int> target)
{
    if (!root.TryGetProperty(property, out var el) || el.ValueKind != JsonValueKind.Array)
    {
        return;
    }
    foreach (var item in el.EnumerateArray())
    {
        if (item.ValueKind == JsonValueKind.Number && item.TryGetInt32(out var value))
        {
            target.Add(value);
        }
    }
}

static void AddTableCells(JsonElement root, string property, ISet<string> target)
{
    if (!root.TryGetProperty(property, out var el) || el.ValueKind != JsonValueKind.Array)
    {
        return;
    }
    foreach (var item in el.EnumerateArray())
    {
        if (item.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(item.GetString()))
        {
            target.Add(item.GetString()!);
        }
        else if (item.ValueKind == JsonValueKind.Array)
        {
            var parts = item.EnumerateArray().Where(x => x.ValueKind == JsonValueKind.Number && x.TryGetInt32(out _)).Select(x => x.GetInt32()).ToList();
            if (parts.Count == 3)
            {
                target.Add($"{parts[0]}:{parts[1]}:{parts[2]}");
            }
        }
    }
}

static void PrintUsage()
{
    Console.Error.WriteLine("Usage:");
    Console.Error.WriteLine("  WordAi.OpenXml inspect <docx>");
    Console.Error.WriteLine("  WordAi.OpenXml assess <source.docx> <patchset.json>");
    Console.Error.WriteLine("  WordAi.OpenXml dry-run <source.docx> <patchset.json> [keep_output]");
    Console.Error.WriteLine("  WordAi.OpenXml apply <source.docx> <patchset.json> <output.docx>");
    Console.Error.WriteLine("  WordAi.OpenXml validate <source.docx> <target.docx> [validation-options.json]");
}
