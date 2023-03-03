using System.Diagnostics;
//convert doc to pdf https://smallpdf.com/word-to-pdf
var gsProgramPath = @"C:\Program Files\gs\gs10.00.0\bin\gswin64.exe";
var currentDirectory = System.IO.Directory.GetCurrentDirectory();
var inputFile = System.IO.Path.Combine(currentDirectory, "input2.pdf");
var outputFile = System.IO.Path.Combine(currentDirectory, "output.pdf");

void CompressPdf(string inputPath, string outputPath)
{
    string[] gsCommand =
    {
        gsProgramPath,
        "-sDEVICE=pdfwrite",
        $"-dPDFSETTINGS=/ebook",
        "-dCompatibilityLevel=1.4",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        "-dDetectDuplicateImages=true",
        "-dCompressFonts=true",
        // "-r150",
        "-dDownsampleColorImages=true",
        "-dColorImageResolution=120",
        "-dMonoImageResolution=140",
        $"-sOutputFile={Path.GetFullPath(outputPath)}",
        Path.GetFullPath(inputPath)
    };
    Process.Start(new ProcessStartInfo
    {
        FileName = gsCommand[0],
        Arguments = string.Join(" ", gsCommand.Skip(1)),
        CreateNoWindow = true,
        UseShellExecute = false,
        RedirectStandardOutput = true,
        RedirectStandardError = true
    }).WaitForExit();
}

CompressPdf(inputFile, outputFile);


Console.WriteLine($"{new FileInfo(inputFile).Length / 1024} KB");
Console.WriteLine($"{new FileInfo(outputFile).Length / 1024} KB");

Process.Start(new ProcessStartInfo
{
    FileName = outputFile,
    UseShellExecute = true
});