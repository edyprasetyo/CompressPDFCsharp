using System.Data.SqlClient;
using System.Diagnostics;
using ClosedXML.Excel;

var currentDirectory = System.IO.Directory.GetCurrentDirectory();
string conDev = "Persist Security Info=true;server=localhost;database=CRMSArchive;uid=sa;pwd=Indorent10!;Integrated Security=false; Connection Timeout=160";

string inputPathFolder = Path.Combine(currentDirectory, "input");
string outputPathFolder = Path.Combine(currentDirectory, "output");
if (Directory.Exists(inputPathFolder))
{
    Directory.Delete(inputPathFolder, true);
}
Directory.CreateDirectory(inputPathFolder);
if (Directory.Exists(outputPathFolder))
{
    Directory.Delete(outputPathFolder, true);
}
Directory.CreateDirectory(outputPathFolder);

var i = 1;
var startTime = DateTime.Now;


var listWorkSheetName = oFunction.GetAllWorkSheetName("Progress Compress.xlsx");
foreach (var workSheet in listWorkSheetName)
{
    var listNullAttachment = oFunction.GetListExcel("Progress Compress.xlsx", workSheet);

    foreach (var dict in listNullAttachment)
    {
        if (oFunction.CheckIfAttachmentNotNull(dict["KodeCompany"].ToString(), dict["KodeCabang"].ToString(), dict["KodeDokumen"].ToString(), dict["NoReferensi"].ToString(), dict["NoReferensi2"].ToString(), dict["NoReferensi3"].ToString()))
        {
            Console.WriteLine("Attachment not null");
            continue;
        }
        using (SqlConnection sqlDev = new SqlConnection(conDev))
        {
            sqlDev.Open();
            string q = "";
            q += " SELECT [KodeCompany],[KodeCabang],[KodeDokumen],[NoReferensi],[NoReferensi2],[NoReferensi3],[Attachment]";
            q += " FROM MsAttachment";
            q += " WHERE KodeCompany = '" + dict["KodeCompany"] + "' AND KodeCabang = '"
            + dict["KodeCabang"] + "' AND KodeDokumen = '" + dict["KodeDokumen"] + "' AND NoReferensi = '"
            + dict["NoReferensi"] + "' AND NoReferensi2 = '" + dict["NoReferensi2"] + "' AND NoReferensi3 = '" + dict["NoReferensi3"] + "'";

            SqlCommand commandDev = new SqlCommand(q, sqlDev);
            SqlDataReader readerDev = commandDev.ExecuteReader();

            string kodeCompany = "";
            string kodeCabang = "";
            string kodeDokumen = "";
            string noReferensi = "";
            string noReferensi2 = "";
            string noReferensi3 = "";
            byte[]? attachment = new byte[0];
            string fileName = "";
            while (readerDev.Read())
            {
                kodeCompany = readerDev["KodeCompany"].ToString();
                kodeCabang = readerDev["KodeCabang"].ToString();
                kodeDokumen = readerDev["KodeDokumen"].ToString();
                noReferensi = readerDev["NoReferensi"].ToString();
                noReferensi2 = readerDev["NoReferensi2"].ToString();
                noReferensi3 = readerDev["NoReferensi3"].ToString();
                attachment = (byte[])readerDev["Attachment"];
                fileName = i.ToString() + ".pdf";
            }
            readerDev.Close();
            sqlDev.Close();

            if (attachment.Length == 0)
            {
                continue;
            }
            i++;
            // var inputPath = Path.Combine(currentDirectory, "input", fileName);
            // var outputPath = Path.Combine(currentDirectory, "output", fileName.Replace(".pdf", "_compressed.pdf"));
            // File.WriteAllBytes(inputPath, attachment);
            // oFunction.Compress(inputPath, outputPath);
            // var deleteInputPath = inputPath;
            // var deleteOutputPath = outputPath;
            // if (new FileInfo(inputPath).Length < new FileInfo(outputPath).Length)
            // {
            //     outputPath = inputPath;
            // }
            // if (new FileInfo(outputPath).Length < 30 * 1024)
            // {
            //     outputPath = inputPath;
            // }
            // Console.WriteLine($"Worksheet: {workSheet}");
            // Console.WriteLine($"File Name: {fileName}");
            // Console.WriteLine($"File Size Before: {Math.Round(new FileInfo(inputPath).Length / 1024.0, 1)} KB");
            // Console.WriteLine($"File Size After: {Math.Round(new FileInfo(outputPath).Length / 1024.0, 1)} KB");
            // Console.WriteLine($"Compression Percentage: {Math.Round(100 - (new FileInfo(outputPath).Length / (double)new FileInfo(inputPath).Length) * 100, 1)}%");
            // Console.WriteLine();

            oFunction.UpdateAttachment(attachment, kodeCompany, kodeCabang, kodeDokumen, noReferensi, noReferensi2, noReferensi3);

            // try
            // {
            //     File.Delete(deleteInputPath);
            // }
            // catch (Exception ex)
            // {
            //     Console.WriteLine(ex.Message);
            // }
            // try
            // {
            //     File.Delete(deleteOutputPath);
            // }
            // catch (Exception ex)
            // {
            //     Console.WriteLine(ex.Message);
            // }

        }
    }

}




