using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

class Program
{
    static bool deleteFiles = false;
    static void Main()
    {
        Console.Write("Delete the word and ppt files when converted (y/n)?");
        string ans = Console.ReadLine();
        deleteFiles = (ans == "y");

        List<string> matchedFiles = FindFiles(Directory.GetCurrentDirectory(), new[] { ".docx", ".pptx", ".xlsx" });

        foreach (string file in matchedFiles)
        {
            Console.WriteLine("processing " + file);
            try
            {
                if (Path.GetExtension(file) == ".docx")
                    WORDtoPDF(file);
                else if (Path.GetExtension(file) == ".pptx")
                    PPTtoPDF(file);
                else if (Path.GetExtension(file) == ".xlsx")
                    ExcelToPDF(file);

                if (deleteFiles)
                    File.Delete(file);
            }
            catch (Exception ex)
            {
                Console.WriteLine("ERROR: " + ex.Message);
            }
        }
        Console.WriteLine("Finished. Press enter to close.");
        Console.ReadLine();
    }

    static List<string> FindFiles(string rootPath, string[] extensions)
    {
        var result = new List<string>();

        try
        {
            var files = Directory.EnumerateFiles(rootPath, "*.*", SearchOption.AllDirectories)
                                 .Where(file => extensions.Contains(Path.GetExtension(file), StringComparer.OrdinalIgnoreCase));

            result.AddRange(files);
        }
        catch (UnauthorizedAccessException ex)
        {
            Console.WriteLine($"Access denied: {ex.Message}");
        }

        return result;
    }

    static dynamic getWord()
    {
        Type wordType = Type.GetTypeFromProgID("Word.Application");
        return Activator.CreateInstance(wordType);
    }
    static dynamic getPPT()
    {
        Type pptType = Type.GetTypeFromProgID("PowerPoint.Application");
        return Activator.CreateInstance(pptType);
    }

    static dynamic getExcel()
    {
        Type excelType = Type.GetTypeFromProgID("Excel.Application");
        dynamic excelApp = Activator.CreateInstance(excelType);
        excelApp.Visible = false;
        excelApp.DisplayAlerts = false;
        return excelApp;
    }

    static void WORDtoPDF(string file)
    {
        //string outputPath = file.Replace("docx", "pdf");
        string outputPath = file + ".pdf";

        var wordApp = getWord(); // new Microsoft.Office.Interop.Word.Application();
        wordApp.Visible = false;

        var doc = wordApp.Documents.Open(file);
        //doc.ExportAsFixedFormat(outputPath, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
        doc.ExportAsFixedFormat(outputPath, 17);
        doc.Close();
        wordApp.Quit();
    }

    static void ExcelToPDF(string file)
    {
        dynamic excelApp = getExcel();
        dynamic workbooks = excelApp.Workbooks;
        dynamic workbook = workbooks.Open(file);
        workbook.ExportAsFixedFormat(0, file + ".pdf");
        workbook.Close(false);
        excelApp.Quit();
        Marshal.ReleaseComObject(workbook);
    }

    static void PPTtoPDF(string file)
    {
        //string outputPath = file.Replace("pptx", "pdf");
        string outputPath = file + ".pdf";

        var app = getPPT(); // new Microsoft.Office.Interop.PowerPoint.Application();
        // app.WindowState = PpWindowState.ppWindowMinimized;
        // app.WindowState = 2;

        //var presentation = app.Presentations.Open(file, WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse);
        var presentation = app.Presentations.Open(file, 0);

        //presentation.SaveAs(outputPath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF);
        presentation.SaveAs(outputPath, 32);
        presentation.Close();
        app.Quit();
    }

}

