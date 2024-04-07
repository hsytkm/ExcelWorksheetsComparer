using ClosedXML.Excel;

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

        Core(sourceExcelPath, saveExcelPath, compareWorksheetNameNew, compareWorksheetNameOld);
    }

    private static string GetSaveExcelPath(string sourceExcelPath)
    {
        var saveName = IsDebug ? "_temp.xlsx" : $"_temp_{DateTime.Now:yyMMdd_HHmmss}.xlsx";

        if (Directory.GetParent(sourceExcelPath) is not { } sourceDir)
            return saveName;

        return Path.Combine(sourceDir.FullName, saveName);
    }

    private static void Core(
        string sourceExcelPath, string saveExcelPath,
        string compareWorksheetNameNew, string compareWorksheetNameOld,
        string? newSheetName = null)
    {
        ArgumentOutOfRangeException.ThrowIfEqual(compareWorksheetNameNew, compareWorksheetNameOld);

        using XLWorkbook workbook = new(sourceExcelPath);

        IXLWorksheet? compareWsNew = null;
        try
        {
            compareWsNew = workbook.Worksheet(compareWorksheetNameNew);
        }
        catch(ArgumentException ex)
        {
            Console.Error.WriteLine(ex);
        }
        if (compareWsNew is null) throw new NullReferenceException(nameof(compareWsNew));

        IXLWorksheet? compareWsOld = null;
        try
        {
            compareWsOld = workbook.Worksheet(compareWorksheetNameOld);
        }
        catch (ArgumentException ex)
        {
            Console.Error.WriteLine(ex);
        }
        if (compareWsOld is null) throw new NullReferenceException(nameof(compareWsOld));

        Console.WriteLine($"Ws1 : Name={compareWsNew.Name}, LastCell={compareWsNew.LastCellUsed()}");
        Console.WriteLine($"Ws2 : Name={compareWsOld.Name}, LastCell={compareWsOld.LastCellUsed()}");

        // 差分セルのリストを行でまとめます
        // （入力されたセルを順にチェックするので、情報が多い新しいワークシートを基準にする必要があります）
        var groupedDifferentRows = compareWsNew
            .EnumerateDifferentCellAddress(compareWsOld)
            .GroupBy(x => x.RowNumber);

        //foreach (var rowCells in groupedDifferentRows)
        //    Console.WriteLine($"DifferentRow={rowCells.Key} ({string.Join(", ", rowCells)})");

        var newSheet = workbook.Worksheets.Add(newSheetName ?? $"_compare_{DateTime.Now:yyMMdd}");
        workbook.ActiveSheet(newSheet);

        // 新シートに差分行をコピーします
        int destRowIndex = 2;   // 先頭は空行
        foreach (int sourceRowIndex in groupedDifferentRows.Select(x => x.Key))
        {
            compareWsOld.CopyRow(newSheet, sourceRowIndex, destRowIndex);
            compareWsNew.CopyRow(newSheet, sourceRowIndex, destRowIndex + 1);
            destRowIndex += 3;  // 空行含む
        }

        // 変化点がないセルを抑制します
        IXLCell lastCell = newSheet.LastCellUsed();
        for (int rowIndex = 1; rowIndex <= lastCell.Address.RowNumber; rowIndex++)
        {
            for (int columnIndex = 1; columnIndex <= lastCell.Address.ColumnNumber; columnIndex++)
            {
                var cell1 = newSheet.Cell(rowIndex, columnIndex);
                var cell2 = newSheet.Cell(rowIndex + 1, columnIndex);
                if (cell1.IsSameCellValue(cell2))
                    cell1.Style.Font.FontColor = cell2.Style.Font.FontColor = XLColor.LightGray;
            }
        }

        try
        {
            workbook.SaveAs(saveExcelPath);
        }
        catch (IOException ex)
        {
            Console.Error.WriteLine(ex);
        }
    }
}
