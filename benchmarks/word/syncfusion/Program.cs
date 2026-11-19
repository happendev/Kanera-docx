using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Licensing;

if (args is ["--version"])
{
    var version = typeof(WordDocument).Assembly.GetName().Version!.ToString();
    Console.WriteLine(JsonSerializer.Serialize(new { syncfusion = version, dotnet = RuntimeInformation.FrameworkDescription }));
    return 0;
}

var licenseKey = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
if (args.Length != 2 || string.IsNullOrEmpty(licenseKey))
{
    Console.Error.WriteLine("Pass an input DOCX and output PDF, or --batch and a job list. Set SYNCFUSION_LICENSE_KEY to a Syncfusion license key for the pinned version.");
    return 2;
}

SyncfusionLicenseProvider.RegisterLicense(licenseKey);

if (args is ["--batch", var jobsFile])
{
    var jobs = JsonSerializer.Deserialize<string[][]>(File.ReadAllText(jobsFile))!;
    var times = new List<double>();
    foreach (var job in jobs)
    {
        var start = Stopwatch.GetTimestamp();
        ConvertDocument(job[0], job[1]);
        times.Add(Stopwatch.GetElapsedTime(start).TotalMilliseconds);
    }
    Console.WriteLine(JsonSerializer.Serialize(times));
}
else
{
    ConvertDocument(args[0], args[1]);
}
return 0;

static void ConvertDocument(string inputPath, string outputPath)
{
    using var input = File.OpenRead(inputPath);
    using var document = new WordDocument(input, FormatType.Automatic);
    using var renderer = new DocIORenderer();
    using var pdf = renderer.ConvertToPDF(document);
    using var output = File.Create(outputPath);
    pdf.Save(output);
}
