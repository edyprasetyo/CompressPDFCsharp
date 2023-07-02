using System.Data.SqlClient;
using System.Diagnostics;
using ClosedXML.Excel;

public class oFunction
{
    public static void Compress(string inputPath, string outputPath)
    {
        var currentDirectory = System.IO.Directory.GetCurrentDirectory();
        var gsProgramPath = System.IO.Path.Combine(currentDirectory, "GhostScript\\gswin32c.exe");

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

    public static List<string> GetAllWorkSheetName(string excelPath)
    {
        var workbook = new XLWorkbook(excelPath);
        var listWorkSheetName = new List<string>();
        foreach (var worksheet in workbook.Worksheets)
        {
            if (int.TryParse(worksheet.Name.Substring(0, 1), out _))
            {
                listWorkSheetName.Add(worksheet.Name);
            }
        }
        return listWorkSheetName;
    }

    public static List<Dictionary<string, object>> GetListExcel(string excelPath, string worksheetName)
    {

        var workbook = new XLWorkbook(excelPath);
        var worksheet = workbook.Worksheet(worksheetName);

        var listExcelData = new List<Dictionary<string, object>>();
        var i = 1;
        foreach (var row in worksheet.RowsUsed())
        {
            if (i > 1)
            {
                listExcelData.Add(new Dictionary<string, object>()
        {
            {"KodeCompany", row.Cell(1).Value},
            {"KodeCabang", row.Cell(2).Value},
            {"KodeDokumen", row.Cell(3).Value},
            {"NoReferensi", row.Cell(4).Value},
            {"NoReferensi2", row.Cell(5).Value},
            {"NoReferensi3", row.Cell(6).Value},
        });
            }
            i++;
        }
        return listExcelData;
    }

    public static List<Dictionary<string, object>> GetAllNullAttachment()
    {
        string conProd = "Persist Security Info=true;server=CSMDBS\\PRODUCTION;database=CRMSArchive;uid=csmapps;pwd=Aud3mars1!;Integrated Security=false; Connection Timeout=160";
        SqlConnection sqlProd = new SqlConnection(conProd);
        sqlProd.Open();
        string q = "";
        q += " SELECT TOP 100000 KodeCompany, KodeCabang, KodeDokumen, NoReferensi, NoReferensi2, NoReferensi3 FROM MsAttachment";
        q += " WHERE Attachment = 0x";
        SqlCommand commandProd = new SqlCommand(q, sqlProd);
        SqlDataReader readerProd = commandProd.ExecuteReader();
        var listNullAttachment = new List<Dictionary<string, object>>();
        if (readerProd.HasRows)
        {
            while (readerProd.Read())
            {
                listNullAttachment.Add(new Dictionary<string, object>()
                {
                    {"KodeCompany", readerProd["KodeCompany"]},
                    {"KodeCabang", readerProd["KodeCabang"]},
                    {"KodeDokumen", readerProd["KodeDokumen"]},
                    {"NoReferensi", readerProd["NoReferensi"]},
                    {"NoReferensi2", readerProd["NoReferensi2"]},
                    {"NoReferensi3", readerProd["NoReferensi3"]},
                });
            }
        }
        sqlProd.Close();
        return listNullAttachment;
    }

    public static bool CheckIfAttachmentNotNull(string kodeCompany, string kodeCabang, string kodeDokumen, string noReferensi, string noReferensi2, string noReferensi3)
    {
        string conProd = "Persist Security Info=true;server=CSMDBS\\PRODUCTION;database=CRMSArchive;uid=csmapps;pwd=Aud3mars1!;Integrated Security=false; Connection Timeout=160";
        SqlConnection sqlProd = new SqlConnection(conProd);
        sqlProd.Open();
        string q = "";
        q += " SELECT Attachment FROM MsAttachment";
        q += " WHERE KodeCompany = '" + kodeCompany + "' AND KodeCabang = '" + kodeCabang +
        "' AND KodeDokumen = '" + kodeDokumen + "' AND NoReferensi = '" + noReferensi +
        "' AND NoReferensi2 = '" + noReferensi2 + "' AND NoReferensi3 = '" + noReferensi3 + "'";
        SqlCommand commandProd = new SqlCommand(q, sqlProd);
        SqlDataReader readerProd = commandProd.ExecuteReader();
        if (readerProd.HasRows)
        {
            while (readerProd.Read())
            {
                if (readerProd["Attachment"] != DBNull.Value)
                {
                    byte[] attachment = (byte[])readerProd["Attachment"];
                    if (attachment.Length > 0)
                    {
                        return true;
                    }
                }
            }
        }
        sqlProd.Close();
        return false;
    }

    public static void UpdateAttachment(byte[] attachment, string kodeCompany, string kodeCabang, string kodeDokumen, string noReferensi, string noReferensi2, string noReferensi3)
    {
        string conProd = "Persist Security Info=true;server=CSMDBS\\PRODUCTION;database=CRMSArchive;uid=csmapps;pwd=Aud3mars1!;Integrated Security=false; Connection Timeout=160";
        SqlConnection sqlProd = new SqlConnection(conProd);
        sqlProd.Open();
        string q = "";
        q += " UPDATE MsAttachment";
        q += " SET Attachment = @Attachment";
        q += " WHERE KodeCompany = '" + kodeCompany + "' AND KodeCabang = '" + kodeCabang +
        "' AND KodeDokumen = '" + kodeDokumen + "' AND NoReferensi = '" + noReferensi +
        "' AND NoReferensi2 = '" + noReferensi2 + "' AND NoReferensi3 = '" + noReferensi3 + "'";
        SqlCommand commandProd = new SqlCommand(q, sqlProd);
        commandProd.Parameters.AddWithValue("@Attachment", attachment);
        commandProd.ExecuteNonQuery();
        sqlProd.Close();
        Console.WriteLine("Update Attachment Success : KodeCompany = " + kodeCompany + ", KodeCabang = " + kodeCabang + ", KodeDokumen = " + kodeDokumen + ", NoReferensi = " + noReferensi + ", NoReferensi2 = " + noReferensi2 + ", NoReferensi3 = " + noReferensi3);
    }

}