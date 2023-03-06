using System.Data.SqlClient;
using System.Diagnostics;

var currentDirectory = System.IO.Directory.GetCurrentDirectory();
var gsProgramPath = System.IO.Path.Combine(currentDirectory, "GhostScript\\gswin32c.exe");

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
        "-dDownsampleColorImages=true",
        "-dColorImageDownsampleThreshold=1",
        "-dColorImageResolution=120",
        "-dDownsampleMonoImages=true",
        "-dMonoImageResolution=120",
        "-dDownScaleFactor=3",
        "-dUseFlateCompression=true",
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

var startTime = DateTime.Now;

string connectionString = "Persist Security Info=true;server=CSMDBS\\PRODUCTION;database=CRMSArchive;uid=csmapps;pwd=Aud3mars1!;Integrated Security=false; Connection Timeout=160";
using (SqlConnection connection = new SqlConnection(connectionString))
{
    connection.Open();
    string q = "";
    // q += " SELECT TOP 100 [KodeCompany],[KodeCabang],[KodeDokumen],[NoReferensi],[NoReferensi2],[NoReferensi3],[Attachment]";
    q += " SELECT TOP 200 [Attachment]";
    q += " FROM MsAttachment";
    q += " WHERE EkstensiFile = '.pdf'";
    q += " ORDER BY ((DATALENGTH(Attachment) / 1024.00) / 1024.00) DESC";
    SqlCommand command = new SqlCommand(q, connection);
    SqlDataReader reader = command.ExecuteReader();
    List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
    var i = 1;
    while (reader.Read())
    {
        byte[] byteData = (byte[])reader["Attachment"];
        string fileName = i.ToString() + ".pdf";
        list.Add(new Dictionary<string, object>()
        {
            // {"KodeCompany", reader["KodeCompany"]},
            // {"KodeCabang", reader["KodeCabang"]},
            // {"KodeDokumen", reader["KodeDokumen"]},
            // {"NoReferensi", reader["NoReferensi"]},
            // {"NoReferensi2", reader["NoReferensi2"]},
            // {"NoReferensi3", reader["NoReferensi3"]},
            {"Attachment", byteData},
            {"FileName", fileName}
        });
        i++;
    }
    reader.Close();
    string inputPath = Path.Combine(currentDirectory, "input");
    string outputPath = Path.Combine(currentDirectory, "output");
    string ouputLess10Percent = Path.Combine(currentDirectory, "outputLessThan10Percent");
    string outputLess20Percent = Path.Combine(currentDirectory, "outputLessThan20Percent");
    string outputLess30Percent = Path.Combine(currentDirectory, "outputLessThan30Percent");
    string outputLess40Percent = Path.Combine(currentDirectory, "outputLessThan40Percent");
    string outputLess50Percent = Path.Combine(currentDirectory, "outputLessThan50Percent");
    string outputLess60Percent = Path.Combine(currentDirectory, "outputLessThan60Percent");
    string outputLess70Percent = Path.Combine(currentDirectory, "outputLessThan70Percent");
    string outputLess80Percent = Path.Combine(currentDirectory, "outputLessThan80Percent");
    string outputLess90Percent = Path.Combine(currentDirectory, "outputLessThan90Percent");
    string outputLess100Percent = Path.Combine(currentDirectory, "outputLessThan100Percent");

    if (Directory.Exists(inputPath))
    {
        Directory.Delete(inputPath, true);
    }
    Directory.CreateDirectory(inputPath);
    if (Directory.Exists(outputPath))
    {
        Directory.Delete(outputPath, true);
    }
    Directory.CreateDirectory(outputPath);
    if (Directory.Exists(ouputLess10Percent))
    {
        Directory.Delete(ouputLess10Percent, true);
    }
    Directory.CreateDirectory(ouputLess10Percent);
    if (Directory.Exists(outputLess20Percent))
    {
        Directory.Delete(outputLess20Percent, true);
    }
    Directory.CreateDirectory(outputLess20Percent);
    if (Directory.Exists(outputLess30Percent))
    {
        Directory.Delete(outputLess30Percent, true);
    }
    Directory.CreateDirectory(outputLess30Percent);
    if (Directory.Exists(outputLess40Percent))
    {
        Directory.Delete(outputLess40Percent, true);
    }
    Directory.CreateDirectory(outputLess40Percent);
    if (Directory.Exists(outputLess50Percent))
    {
        Directory.Delete(outputLess50Percent, true);
    }
    Directory.CreateDirectory(outputLess50Percent);
    if (Directory.Exists(outputLess60Percent))
    {
        Directory.Delete(outputLess60Percent, true);
    }
    Directory.CreateDirectory(outputLess60Percent);
    if (Directory.Exists(outputLess70Percent))
    {
        Directory.Delete(outputLess70Percent, true);
    }
    Directory.CreateDirectory(outputLess70Percent);
    if (Directory.Exists(outputLess80Percent))
    {
        Directory.Delete(outputLess80Percent, true);
    }
    Directory.CreateDirectory(outputLess80Percent);
    if (Directory.Exists(outputLess90Percent))
    {
        Directory.Delete(outputLess90Percent, true);
    }
    Directory.CreateDirectory(outputLess90Percent);
    if (Directory.Exists(outputLess100Percent))
    {
        Directory.Delete(outputLess100Percent, true);
    }
    Directory.CreateDirectory(outputLess100Percent);

    double averagePercentage = 0;
    double totalMBBefore = 0;
    double totalMBAfter = 0;
    string csvPath = Path.Combine(currentDirectory, "output", "output.csv");
    string csvHeader = "File Name,File Size Before (MB),File Size After (MB),Compression Percentage (%), Saved (MB)";
    File.AppendAllText(csvPath, csvHeader + Environment.NewLine);

    foreach (var file in list)
    {
        string fileName = file["FileName"].ToString()!;
        byte[] byteData = (byte[])file["Attachment"];
        inputPath = Path.Combine(currentDirectory, "input", fileName);
        outputPath = Path.Combine(currentDirectory, "output", fileName.Replace(".pdf", "_compressed.pdf"));
        File.WriteAllBytes(inputPath, byteData);
        CompressPdf(inputPath, outputPath);
        if (new FileInfo(inputPath).Length < new FileInfo(outputPath).Length)
        {
            outputPath = inputPath;

        }

        if (new FileInfo(outputPath).Length < 30 * 1024)
        {
            outputPath = inputPath;

        }

        Console.WriteLine($"File Name: {fileName}");
        Console.WriteLine($"File Size Before: {Math.Round(new FileInfo(inputPath).Length / 1024.0 / 1024.0, 1)} MB");
        Console.WriteLine($"File Size After: {Math.Round(new FileInfo(outputPath).Length / 1024.0 / 1024.0, 1)} MB");
        Console.WriteLine($"Compression Percentage: {Math.Round(100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100, 1)}%");
        Console.WriteLine();
        averagePercentage += 100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100;
        totalMBBefore += new FileInfo(inputPath).Length / 1024.0 / 1024.0;
        totalMBAfter += new FileInfo(outputPath).Length / 1024.0 / 1024.0;

        if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 10)
        {
            File.Copy(outputPath, Path.Combine(ouputLess10Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }
        else if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 20)
        {
            File.Copy(outputPath, Path.Combine(outputLess20Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }
        else if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 30)
        {
            File.Copy(outputPath, Path.Combine(outputLess30Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }
        else if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 40)
        {
            File.Copy(outputPath, Path.Combine(outputLess40Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }
        else if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 50)
        {
            File.Copy(outputPath, Path.Combine(outputLess50Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }
        else if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 60)
        {
            File.Copy(outputPath, Path.Combine(outputLess60Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }
        else if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 70)
        {
            File.Copy(outputPath, Path.Combine(outputLess70Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }
        else if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 80)
        {
            File.Copy(outputPath, Path.Combine(outputLess80Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }
        else if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 90)
        {
            File.Copy(outputPath, Path.Combine(outputLess90Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }
        else if (100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100 < 100)
        {
            File.Copy(outputPath, Path.Combine(outputLess100Percent, fileName.Replace(".pdf", "_compressed.pdf")));
        }

        string csvData = $"{fileName},{Math.Round(new FileInfo(inputPath).Length / 1024.0 / 1024.0, 1)},{Math.Round(new FileInfo(outputPath).Length / 1024.0 / 1024.0, 1)},{Math.Round(100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100, 1)},{Math.Round(new FileInfo(inputPath).Length / 1024.0 / 1024.0, 1) - Math.Round(new FileInfo(outputPath).Length / 1024.0 / 1024.0, 1)}";
        File.AppendAllText(csvPath, csvData + Environment.NewLine);

    }
    averagePercentage = Math.Round(averagePercentage / list.Count, 1);
    Console.WriteLine($"Average Compression Percentage: {averagePercentage}%");
    Console.WriteLine($"Total MB Before: {Math.Round(totalMBBefore, 1)} MB");
    Console.WriteLine($"Total MB After: {Math.Round(totalMBAfter, 1)} MB");
    Console.WriteLine($"Total MB Saved: {Math.Round(totalMBBefore - totalMBAfter, 1)} MB");

    Console.WriteLine();
    Console.WriteLine($"Total Files in {ouputLess10Percent}: {Directory.GetFiles(ouputLess10Percent).Length}");
    Console.WriteLine($"Total Files in {outputLess20Percent}: {Directory.GetFiles(outputLess20Percent).Length}");
    Console.WriteLine($"Total Files in {outputLess30Percent}: {Directory.GetFiles(outputLess30Percent).Length}");
    Console.WriteLine($"Total Files in {outputLess40Percent}: {Directory.GetFiles(outputLess40Percent).Length}");
    Console.WriteLine($"Total Files in {outputLess50Percent}: {Directory.GetFiles(outputLess50Percent).Length}");
    Console.WriteLine($"Total Files in {outputLess60Percent}: {Directory.GetFiles(outputLess60Percent).Length}");
    Console.WriteLine($"Total Files in {outputLess70Percent}: {Directory.GetFiles(outputLess70Percent).Length}");
    Console.WriteLine($"Total Files in {outputLess80Percent}: {Directory.GetFiles(outputLess80Percent).Length}");
    Console.WriteLine($"Total Files in {outputLess90Percent}: {Directory.GetFiles(outputLess90Percent).Length}");
    Console.WriteLine($"Total Files in {outputLess100Percent}: {Directory.GetFiles(outputLess100Percent).Length}");
    Console.WriteLine();
}

Console.WriteLine($"Total Execution Time: {DateTime.Now.Subtract(startTime).ToString(@"hh\:mm\:ss")}");

