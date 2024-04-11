namespace ExcelWorksheetsComparer;

internal class Program
{
    private const bool IsDebug =
#if DEBUG
        true;
#else
        false;
#endif

    private static void Main(string[] args)
    {
        if (!IsDebug && args.Length < 1)
        {
            Console.WriteLine("Usage : Program.exe [SourceExcelPath]");
            Console.WriteLine("  Compare \"new\" and \"old\" worksheets in the workbook.");
            return;
        }

        string sourceExcelPath = IsDebug ? @"D:\data\excel\compare_source.xlsx" : args[0];
        string saveExcelPath = GetSaveExcelPath(sourceExcelPath);
        string compareWorksheetNameNew = "new";
        string compareWorksheetNameOld = "old";

        Console.WriteLine($"Source Excel : \"{sourceExcelPath}\"");
        Console.WriteLine($"Save Excel   : \"{saveExcelPath}\"");
        Console.WriteLine($"Compare New WorkSheet Name : \"{compareWorksheetNameNew}\"");
        Console.WriteLine($"Compare Old WorkSheet Name : \"{compareWorksheetNameOld}\"");
        Console.WriteLine("");

        UsingClosedXML.Core(sourceExcelPath, saveExcelPath, compareWorksheetNameNew, compareWorksheetNameOld);
    }

    private static string GetSaveExcelPath(string sourceExcelPath)
    {
        var saveName = IsDebug ? "_temp.xlsx" : $"_temp_{DateTime.Now:yyMMdd_HHmmss}.xlsx";

        if (Directory.GetParent(sourceExcelPath) is not { } sourceDir)
            return saveName;

        return Path.Combine(sourceDir.FullName, saveName);
    }
}
