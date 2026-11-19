using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json;
using Aspose.Words;
using Aspose.Words.Saving;

if (args is ["--version"])
{
    var version = typeof(Document).Assembly.GetName().Version!.ToString();
    Console.WriteLine(JsonSerializer.Serialize(new { aspose = version, dotnet = RuntimeInformation.FrameworkDescription }));
    return 0;
}

var licensePath = Environment.GetEnvironmentVariable("ASPOSE_WORDS_LICENSE");
if (args.Length != 2 || string.IsNullOrEmpty(licensePath) || !File.Exists(licensePath))
{
    Console.Error.WriteLine("Pass an input DOCX and output PDF, or --batch and a job list. Set ASPOSE_WORDS_LICENSE to a local Aspose.Words for .NET license.");
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
    var doc = new Document(input);
    doc.Save(outputPath, new PdfSaveOptions());
}
