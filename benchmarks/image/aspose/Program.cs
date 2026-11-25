using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Saving;

// Renders each page of a DOCX to <output dir>/page-<n>.png at 150 dpi.
const float Resolution = 150;

if (args is ["--version"])
{
    var version = typeof(Document).Assembly.GetName().Version!.ToString();
    Console.WriteLine(JsonSerializer.Serialize(new { aspose = version, dotnet = RuntimeInformation.FrameworkDescription }));
    return 0;
}

var licensePath = Environment.GetEnvironmentVariable("ASPOSE_WORDS_LICENSE");
if (args.Length != 2 || string.IsNullOrEmpty(licensePath) || !File.Exists(licensePath))
{
    Console.Error.WriteLine("Pass an input DOCX and output directory, or --batch and a job list. Set ASPOSE_WORDS_LICENSE to a local Aspose.Words for .NET license.");
    return 2;
}

using (var license = File.OpenRead(licensePath))
{
    new License().SetLicense(license);
}

if (args is ["--batch", var jobsFile])
{
    var jobs = JsonSerializer.Deserialize<string[][]>(File.ReadAllText(jobsFile))!;
    var times = new List<double>();
    foreach (var job in jobs)
    {
        var start = Stopwatch.GetTimestamp();
        RenderDocument(job[0], job[1]);
        times.Add(Stopwatch.GetElapsedTime(start).TotalMilliseconds);
    }
    Console.WriteLine(JsonSerializer.Serialize(times));
}
else
{
    RenderDocument(args[0], args[1]);
}
return 0;

static void RenderDocument(string inputPath, string outputDirectory)
{
    using var input = File.OpenRead(inputPath);
    var doc = new Document(input);
    Directory.CreateDirectory(outputDirectory);
    // PageCount lays the document out once; each Save reuses that layout for one page.
    for (var page = 0; page < doc.PageCount; page++)
    {
        var options = new ImageSaveOptions(SaveFormat.Png) { Resolution = Resolution, PageSet = new PageSet(page) };
        doc.Save(Path.Combine(outputDirectory, $"page-{page + 1}.png"), options);
    }
}
