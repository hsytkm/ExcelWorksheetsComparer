using ClosedXML.Excel;
using ExcelWorksheetsComparer.Extensions.ClosedXML;

namespace ExcelWorksheetsComparer;

internal static class UsingClosedXML
{
    internal static void Core(
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
        catch (ArgumentException ex)
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
