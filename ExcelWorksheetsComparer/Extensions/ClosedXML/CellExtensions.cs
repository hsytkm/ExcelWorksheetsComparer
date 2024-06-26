﻿using ClosedXML.Excel;

namespace ExcelWorksheetsComparer.Extensions.ClosedXML;

internal static class CellExtensions
{
    internal static bool IsSameCellValue(this IXLCell cell1, IXLCell cell2)
    {
        return cell1.Value.Equals(cell2.Value);
    }
}
