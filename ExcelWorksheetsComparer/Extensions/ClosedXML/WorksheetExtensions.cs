using ClosedXML.Excel;

namespace ExcelWorksheetsComparer.Extensions.ClosedXML;

internal static class WorksheetExtensions
{
    internal static IEnumerable<IXLAddress> EnumerateDifferentCellAddress(this IXLWorksheet ws1, IXLWorksheet ws2)
    {
#if false
        // ws1.CellsUsed() だと ws.LastCellUsed() まで列挙されないケースがあるので無効化します
        // 右下のセルが空白だった場合に列挙されてない気がします。ClosedXML="0.104.0-preview2"
        foreach (IXLCell cell1 in ws1.CellsUsed())
        {
            IXLCell cell2 = ws2.Cell(cell1.Address);
            if (!cell1.IsSameCellValue(cell2))
                yield return cell1.Address;
        }
#else
        var lastCell = ws1.LastCellUsed();
        int rowEnd = lastCell.Address.RowNumber;
        int colEnd = lastCell.Address.ColumnNumber;
        for (int row = 1; row <= rowEnd; row++)
        {
            for (int col = 1; col <= colEnd; col++)
            {
                var cell1 = ws1.Cell(row, col);
                var cell2 = ws2.Cell(row, col);
                if (!cell1.IsSameCellValue(cell2))
                    yield return cell1.Address;
            }
        }
#endif
    }

    internal static IOrderedEnumerable<int> EnumerateDifferentRowNumber(this IXLWorksheet ws1, IXLWorksheet ws2)
    {
        // 先頭行は 1 になります
        return ws1.EnumerateDifferentCellAddress(ws2)
            .Select(x => x.RowNumber)
            .Distinct().Order();
    }

    internal static void CopyRow(this IXLWorksheet sourceWs, IXLWorksheet destWs, int sourceRowIndex, int destRowIndex)
    {
        // 先頭行は 1 です
        ArgumentOutOfRangeException.ThrowIfLessThan(sourceRowIndex, 0);
        ArgumentOutOfRangeException.ThrowIfLessThan(destRowIndex, 0);

        foreach (IXLCell sourceCell in sourceWs.Row(sourceRowIndex).CellsUsed())
        {
            var destCell = destWs.Cell(destRowIndex, sourceCell.WorksheetColumn().ColumnNumber());
            destCell.Value = sourceCell.Value;
            destCell.Style = sourceCell.Style;
        }
    }
}
