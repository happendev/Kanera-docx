using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using MiniSoftware;

if (args is ["--version"])
{
    var version = Assembly.GetExecutingAssembly().GetCustomAttributes<AssemblyMetadataAttribute>()
        .Single(attribute => attribute.Key == "MiniPdfVersion").Value;
    Console.WriteLine(JsonSerializer.Serialize(new { minipdf = version, dotnet = RuntimeInformation.FrameworkDescription }));
    return 0;
}

var options = new MiniPdfConversionOptions { Compress = true };
if (args is ["--batch", var jobsFile])
{
    var jobs = JsonSerializer.Deserialize<string[][]>(File.ReadAllText(jobsFile))!;
    var times = new List<double?>(jobs.Length);
    foreach (var job in jobs)
    {
        var start = Stopwatch.GetTimestamp();
        if (ConvertDocument(job[0], job[1], options))
            times.Add(Stopwatch.GetElapsedTime(start).TotalMilliseconds);
        else
            times.Add(null);
    }
    Console.WriteLine(JsonSerializer.Serialize(times));
    return 0;
}

if (args is [var input, var output])
    return ConvertDocument(input, output, options) ? 0 : 1;

Console.Error.WriteLine("Pass an input DOCX and output PDF, or --batch and a job list.");
return 2;

static bool ConvertDocument(string input, string output, MiniPdfConversionOptions options)
{
    try
    {
        MiniPdf.ConvertToPdf(input, output, options);
        return true;
    }
    catch (Exception error)
    {
        Console.Error.WriteLine($"{input}: {error.Message}");
        return false;
    }
}
