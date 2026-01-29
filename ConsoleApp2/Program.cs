using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using Word = Microsoft.Office.Interop.Word;

class Program
{
    static void Main()
    {
        string excelPath = @"C:\Users\Tzivy\Desktop\נתוני_תלמידים.xlsx";
        string processedExcelPath = @"C:\Users\Tzivy\Desktop\נתוני_תלמידים_מעובדים.xlsx";
        string wordTemplate = @"C:\Users\Tzivy\Desktop\תבנית.docx";
        string pdfFolder = @"C:\Users\Tzivy\Desktop\PDFs\";

        Directory.CreateDirectory(pdfFolder);

        // --------------------------
        // 1️⃣ קריאה וטיוב Excel
        // --------------------------
        var workbook = new XLWorkbook(excelPath);
        var ws = workbook.Worksheet(1);
        var lastRow = ws.LastRowUsed().RowNumber();

        // מציאת עמודות לפי שמות
        var headerRow = ws.Row(1);
        int nameCol = headerRow.Cells().First(c => c.GetString() == "שם").Address.ColumnNumber;
        int theoreticalCol = headerRow.Cells().First(c => c.GetString() == "ציון תאורטי").Address.ColumnNumber;
        int practicalCol = headerRow.Cells().First(c => c.GetString() == "ציון מעשי").Address.ColumnNumber;

        // הוספת עמודת ציון סופי וCustomMessage
        int finalCol = ws.LastColumnUsed().ColumnNumber() + 1;
        ws.Cell(1, finalCol).Value = "ציון סופי";
        int messageCol = finalCol + 1;
        ws.Cell(1, messageCol).Value = "CustomMessage";

        // קריאה לכל השורות והסרת כפילויות לפי שם
        var rows = ws.Rows(2, lastRow).ToList();
        var distinctRows = rows
            .GroupBy(r => r.Cell(nameCol).GetString().Trim().ToLower())
            .Select(g => g.First())
            .ToList();

        int currentRow = 2;
        foreach (var row in distinctRows)
        {
            string name = FormatName(row.Cell(nameCol).GetString());
            ws.Cell(currentRow, nameCol).Value = name;

            double theoretical = 0, practical = 0;
            double.TryParse(row.Cell(theoreticalCol).GetString(), out theoretical);
            double.TryParse(row.Cell(practicalCol).GetString(), out practical);

            double finalScore = practical * 0.6 + theoretical * 0.4;
            ws.Cell(currentRow, theoreticalCol).Value = theoretical;
            ws.Cell(currentRow, practicalCol).Value = practical;
            ws.Cell(currentRow, finalCol).Value = Math.Round(finalScore, 2);

            // יצירת הודעה מותאמת לפי ציון
            string message = finalScore >= 90
                ? $"הרינו להודיעך כי עברת בהצלחה את ההכשרה. הציון הסופי שלך הינו {finalScore}"
                : "הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.";

            ws.Cell(currentRow, messageCol).Value = message;

            currentRow++;
        }

        if (currentRow <= lastRow)
            ws.Rows(currentRow, lastRow).Delete();

        workbook.SaveAs(processedExcelPath);
        workbook.Dispose();

        Console.WriteLine("✔ Excel מעובד ונקי נשמר בהצלחה!");

        Word.Application wordApp = null;
        Word.Document templateDoc = null;
        try
        {
            wordApp = new Word.Application();
            wordApp.Visible = false;

            templateDoc = wordApp.Documents.Open(wordTemplate);
            templateDoc.MailMerge.OpenDataSource(
                Name: processedExcelPath,
                ConfirmConversions: false,
                ReadOnly: true,
                LinkToSource: true,
                AddToRecentFiles: false,
                Revert: false,
                Format: Word.WdOpenFormat.wdOpenFormatAuto
            );

            int recordCount = templateDoc.MailMerge.DataSource.RecordCount;

            for (int i = 1; i <= recordCount; i++)
            {
                templateDoc.MailMerge.DataSource.ActiveRecord = i;

                double finalScore = Convert.ToDouble(templateDoc.MailMerge.DataSource.DataFields["ציון סופי"].Value);
                string studentName = templateDoc.MailMerge.DataSource.DataFields["שם"].Value;

                if (finalScore >= 70)
                {
                    string pdfPath = Path.Combine(pdfFolder, $"{studentName}.pdf");

                    templateDoc.MailMerge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument;
                    templateDoc.MailMerge.Execute();

                    Word.Document mergedDoc = wordApp.ActiveDocument;
                    mergedDoc.ExportAsFixedFormat(pdfPath, Word.WdExportFormat.wdExportFormatPDF);
                    mergedDoc.Close(false);

                    Console.WriteLine($"✔ PDF נוצר עבור: {studentName} (ציון סופי: {finalScore})");
                }
            }

            Console.WriteLine("🎉 כל ה-PDFs נוצרו בהצלחה!");
        }
        finally
        {
            if (templateDoc != null) templateDoc.Close(false);
            if (wordApp != null)
            {
                wordApp.Quit(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    // פונקציה לתיקון שמות
    static string FormatName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "";
        name = Regex.Replace(name.Trim(), @"\s+", " ");
        var parts = name.Split(' ');
        for (int i = 0; i < parts.Length; i++)
            if (parts[i].Length > 0)
                parts[i] = char.ToUpper(parts[i][0]) + parts[i].Substring(1).ToLower();
        return string.Join(" ", parts);
    }
}
