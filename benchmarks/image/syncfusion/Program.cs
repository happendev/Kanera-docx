using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Licensing;

// Renders each page of a DOCX to <output dir>/page-<n>.png. DocIO 28 renders page images at a
// fixed 96 dpi; RenderAsImages has no resolution setting.
if (args is ["--version"])
{
    var version = typeof(WordDocument).Assembly.GetName().Version!.ToString();
    Console.WriteLine(JsonSerializer.Serialize(new { syncfusion = version, dotnet = RuntimeInformation.FrameworkDescription }));
    return 0;
}

var licenseKey = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
if (args.Length != 2 || string.IsNullOrEmpty(licenseKey))
{
    Console.Error.WriteLine("Pass an input DOCX and output directory, or --batch and a job list. Set SYNCFUSION_LICENSE_KEY to a Syncfusion license key for the pinned version.");
    return 2;
}

SyncfusionLicenseProvider.RegisterLicense(licenseKey);
// An invalid key renders a trial banner on every page, which the ink score cannot tell apart.
if (!SyncfusionLicenseProvider.ValidateLicense(Platform.FileFormats, out var message))
{
    Console.Error.WriteLine($"SYNCFUSION_LICENSE_KEY is not valid for DocIO {typeof(WordDocument).Assembly.GetName().Version}: {message}");
    return 2;
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
    using var document = new WordDocument(input, FormatType.Automatic);
    using var renderer = new DocIORenderer();
    var pages = document.RenderAsImages();
    Directory.CreateDirectory(outputDirectory);
    for (var i = 0; i < pages.Length; i++)
    {
        using var page = pages[i];
        page.Position = 0;
        using var output = File.Create(Path.Combine(outputDirectory, $"page-{i + 1}.png"));
        page.CopyTo(output);
    }
}
