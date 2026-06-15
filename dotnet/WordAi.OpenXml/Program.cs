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
            Console.WriteLine(JsonSerializer.Serialize(new WordValidator().Validate(args[1], args[2]), JsonSupport.Options));
            return 0;

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

static void PrintUsage()
{
    Console.Error.WriteLine("Usage:");
    Console.Error.WriteLine("  WordAi.OpenXml inspect <docx>");
    Console.Error.WriteLine("  WordAi.OpenXml assess <source.docx> <patchset.json>");
    Console.Error.WriteLine("  WordAi.OpenXml dry-run <source.docx> <patchset.json> [keep_output]");
    Console.Error.WriteLine("  WordAi.OpenXml apply <source.docx> <patchset.json> <output.docx>");
    Console.Error.WriteLine("  WordAi.OpenXml validate <source.docx> <target.docx>");
}
