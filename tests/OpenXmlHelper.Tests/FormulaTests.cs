using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlHelper.Tests;
using static OpenXmlHelper.OpenXmlHelper;
using Xunit;

namespace OpenXmlHelper.Tests;

/// <summary>
/// 验证删除操作后公式引用的自动重排：
/// 区域收缩、行/列位移、相对/绝对引用、被删引用变为 #REF!。
/// </summary>
public class FormulaTests
{
    private static string BuildSheet()
    {
        var te = TestExcel.Create();
        for (int i = 1; i <= 10; i++)
            te.SetNumber(i, 1, i);                 // A1..A10 = 1..10
        te.Save();
        var path = te.FilePath;
        te.Dispose();
        return path;
    }

    [Fact]
    public void DeleteRow_ShrinksSumRange()
    {
        var path = BuildSheet();
        // 在已存在文件中追加公式单元格
        AppendFormula(path, "B1", "SUM(A1:A10)");

        var r = DeleteRows(path, "Sheet1", 5, 5);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Equal("SUM(A1:A9)", TestExcel.GetFormula(path, "Sheet1", "B1"));
        // A6(原) -> A5
        Assert.Equal("6", TestExcel.GetValue(path, "Sheet1", "A5"));
        Assert.Equal("10", TestExcel.GetValue(path, "Sheet1", "A9"));
        Assert.Null(TestExcel.GetValue(path, "Sheet1", "A10"));   // 末行已移除
    }

    [Fact]
    public void DeleteRows_ShrinksRangeAcrossMultipleRows()
    {
        var path = BuildSheet();
        AppendFormula(path, "B1", "SUM(A1:A10)");

        var r = DeleteRows(path, "Sheet1", 3, 4);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Equal("SUM(A1:A8)", TestExcel.GetFormula(path, "Sheet1", "B1"));
    }

    [Fact]
    public void DeleteRow_AdjustsAbsoluteAndRelativeReferences()
    {
        var te = TestExcel.Create();
        te.SetNumber(1, 1, 1);  // A1
        te.SetNumber(2, 1, 2);  // A2
        te.SetNumber(3, 1, 3);  // A3
        te.SetFormula(1, 3, "$A$1+A3"); // C1
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 2, 2);
        Assert.True(r.Success, r.ErrorMessage);

        // 绝对引用 $A$1 行1不变；相对引用 A3 -> A2
        Assert.Equal("$A$1+A2", TestExcel.GetFormula(path, "Sheet1", "C1"));
    }

    [Fact]
    public void DeleteRow_DeletedReferenceBecomesRefError()
    {
        var te = TestExcel.Create();
        te.SetNumber(2, 1, 5);   // A2
        te.SetFormula(1, 2, "A2+B2"); // B1
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 2, 2);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Equal("#REF!+#REF!", TestExcel.GetFormula(path, "Sheet1", "B1"));
    }

    [Fact]
    public void DeleteRow_RangeFullyDeletedBecomesRefError()
    {
        var te = TestExcel.Create();
        te.SetNumber(2, 1, 5);   // A2
        te.SetFormula(1, 2, "SUM(A2:A2)"); // B1
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 2, 2);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Equal("SUM(#REF!)", TestExcel.GetFormula(path, "Sheet1", "B1"));
    }

    [Fact]
    public void DeleteColumn_AdjustsColumnReferences()
    {
        var te = TestExcel.Create();
        te.SetNumber(1, 1, 1); // A1
        te.SetNumber(1, 2, 2); // B1
        te.SetNumber(1, 3, 3); // C1
        te.SetFormula(1, 4, "SUM(A1:C1)"); // D1
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteColumns(path, "Sheet1", 2, 2);
        Assert.True(r.Success, r.ErrorMessage);

        // B 列删除：A1 不变，C1 -> B1
        Assert.Equal("SUM(A1:B1)", TestExcel.GetFormula(path, "Sheet1", "C1"));
    }

    [Fact]
    public void DeleteRow_ReferenceBelowShifts_SingleCell()
    {
        var te = TestExcel.Create();
        te.SetNumber(1, 1, 1);
        te.SetNumber(3, 1, 30);  // A3
        te.SetFormula(2, 2, "A3"); // B2 引用 A3（B2 不在被删行内）
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 1, 1);
        Assert.True(r.Success, r.ErrorMessage);

        // 删除第1行后 B2 -> B1，引用 A3 -> A2
        Assert.Equal("A2", TestExcel.GetFormula(path, "Sheet1", "B1"));
    }

    /// <summary>在已保存文件中追加一个带公式的单元格。</summary>
    private static void AppendFormula(string path, string cellRef, string formula)
    {
        using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(path, true);
        var wb = doc.WorkbookPart!;
        var sheetEl = wb.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name?.Value == "Sheet1");
        var ws = (DocumentFormat.OpenXml.Packaging.WorksheetPart)wb.GetPartById(sheetEl.Id!.Value);
        var sd = ws.Worksheet.GetFirstChild<SheetData>()!;
        int rowIdx = ParseRow(cellRef);
        Row? row = null;
        foreach (var r in sd.Elements<Row>())
            if (r.RowIndex?.Value == (uint)rowIdx) { row = r; break; }
        row ??= sd.AppendChild(new Row { RowIndex = (uint)rowIdx });

        var cell = new Cell { CellReference = cellRef };
        cell.AppendChild(new CellFormula(formula));
        row.AppendChild(cell);

        ws.Worksheet.Save();
        wb.Workbook.Save();
    }

    private static int ParseRow(string cellRef)
    {
        int i = 0;
        while (i < cellRef.Length && cellRef[i] is '$' or (>= 'A' and <= 'Z') or (>= 'a' and <= 'z')) i++;
        return int.Parse(cellRef[i..], System.Globalization.CultureInfo.InvariantCulture);
    }
}
