using ClosedXML.Excel;

internal static class WorkbookExtensions
{
    // 指定シートを選択状態にします
    internal static void ActiveSheet(this IXLWorkbook workbook, IXLWorksheet worksheet)
    {
        // 以下テキトーに生み出しました。よりよい手段があると思いますが動けばよい
        foreach (var ws in workbook.Worksheets)
        {
            ws.SetTabSelected(false);
            ws.SetTabActive(false);
        }
        worksheet.SetTabSelected(true);
        worksheet.SetTabActive(true);
    }
}
