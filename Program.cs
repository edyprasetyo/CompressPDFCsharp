using System.Diagnostics;
//convert doc to pdf https://smallpdf.com/word-to-pdf

var currentDirectory = System.IO.Directory.GetCurrentDirectory();
var gsProgramPath = System.IO.Path.Combine(currentDirectory, "GhostScript\\gswin64c.exe");

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
    })!.WaitForExit();
}

var inputFiles = Directory.GetFiles(currentDirectory + "/input", "*.pdf", SearchOption.AllDirectories);
foreach (var inputFile in inputFiles)
{
    var outputFile = Path.Combine(currentDirectory, "output", Path.GetFileNameWithoutExtension(inputFile) + "_compressed" + Path.GetExtension(inputFile));
    Directory.CreateDirectory(Path.GetDirectoryName(outputFile)!);

    CompressPdf(inputFile, outputFile);

    Console.WriteLine($"{Path.GetFileName(inputFile)} {new FileInfo(inputFile).Length / 1024} KB");
    Console.WriteLine($"{Path.GetFileName(outputFile)} {new FileInfo(outputFile).Length / 1024} KB");
    Console.WriteLine();
}





// Process.Start(new ProcessStartInfo
// {
//     FileName = outputFile,
//     UseShellExecute = true
// });