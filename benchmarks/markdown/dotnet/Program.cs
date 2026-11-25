// Converts DOCX to Markdown, or to plain text for a .txt output, with Aspose.Words or Syncfusion
// DocIO, using each library's defaults.
// MarkdownBenchmark <aspose|syncfusion> <input.docx> <output.md|output.txt>
// MarkdownBenchmark <aspose|syncfusion> --batch <jobs.json>   (prints one time in ms per job, or null)
// MarkdownBenchmark --version
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;

if (args is ["--version"])
{
    Console.WriteLine(JsonSerializer.Serialize(new
    {
        aspose = typeof(Aspose.Words.Document).Assembly.GetName().Version!.ToString(),
        syncfusion = typeof(Syncfusion.DocIO.DLS.WordDocument).Assembly.GetName().Version!.ToString(),
        dotnet = RuntimeInformation.FrameworkDescription,
    }));
    return 0;
}
if (args.Length != 3 || args[0] is not ("aspose" or "syncfusion"))
{
    Console.Error.WriteLine("Usage: MarkdownBenchmark <aspose|syncfusion> <input.docx> <output.md|output.txt> or --batch <jobs.json>");
    return 2;
}

var engine = args[0];
if (engine == "aspose")
{
    var path = Environment.GetEnvironmentVariable("ASPOSE_WORDS_LICENSE");
    if (string.IsNullOrEmpty(path) || !File.Exists(path))
    {
        Console.Error.WriteLine("Set ASPOSE_WORDS_LICENSE to a local Aspose.Words for .NET license.");
        return 2;
    }
    using var license = File.OpenRead(path);
    new Aspose.Words.License().SetLicense(license);
}
else
{
    var key = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
    if (string.IsNullOrEmpty(key))
    {
        Console.Error.WriteLine("Set SYNCFUSION_LICENSE_KEY to a Syncfusion license key.");
        return 2;
    }
    Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(key);
}

if (args[1] == "--batch")
{
    var jobs = JsonSerializer.Deserialize<string[][]>(File.ReadAllText(args[2]))!;
    var times = new List<double?>();
    foreach (var job in jobs)
    {
        var start = Stopwatch.GetTimestamp();
        try
        {
            Convert(engine, job[0], job[1]);
            times.Add(Stopwatch.GetElapsedTime(start).TotalMilliseconds);
        }
        catch
        {
            times.Add(null);
        }
    }
    Console.WriteLine(JsonSerializer.Serialize(times));
}
else
{
    Convert(engine, args[1], args[2]);
}
return 0;

static void Convert(string engine, string input, string output)
{
    bool text = Path.GetExtension(output) == ".txt";
    if (engine == "aspose")
    {
        // Pictures go next to the Markdown file, as Aspose does by default.
        new Aspose.Words.Document(input).Save(output, text ? Aspose.Words.SaveFormat.Text : Aspose.Words.SaveFormat.Markdown);
        return;
    }
    using var source = File.OpenRead(input);
    using var document = new Syncfusion.DocIO.DLS.WordDocument(source, Syncfusion.DocIO.FormatType.Docx);
    using var target = File.Create(output);
    document.Save(target, text ? Syncfusion.DocIO.FormatType.Txt : Syncfusion.DocIO.FormatType.Markdown);
}
