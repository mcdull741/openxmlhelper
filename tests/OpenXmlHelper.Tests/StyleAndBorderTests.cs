using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlHelper.Helpers;
using static OpenXmlHelper.OpenXmlHelper;
using Xunit;

namespace OpenXmlHelper.Tests;

/// <summary>
/// 验证删除后：相邻单元格原有格式（字体/颜色）保留，以及边框冲突处理。
/// </summary>
public class StyleAndBorderTests
{
    [Fact]
    public void DeleteRow_PreservesFontAndColorOfShiftedCell()
    {
        var te = TestExcel.Create();
        var fontId = te.Font(bold: true, rgb: "FFFF0000"); // 粗体红色
        var xf = te.CellFormat(fontId: fontId, applyFont: true);
        te.SetNumberWithStyle(3, 1, 42, xf);   // A3：粗体红
        te.SetNumber(1, 1, 1);                  // A1 占位以便删除首行
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 1, 1);
        Assert.True(r.Success, r.ErrorMessage);

        // A3 上移到 A2，字体应保持粗体红色
        var (bold, rgb) = TestExcel.GetFont(path, "Sheet1", "A2");
        Assert.True(bold);
        Assert.Equal("FFFF0000", rgb);
        Assert.Equal("42", TestExcel.GetValue(path, "Sheet1", "A2"));
    }

    [Fact]
    public void DeleteRow_TransfersDeletedTopBorderToAboveNeighbor()
    {
        // A4 无边框，A5 有上边框(thin)。删除第5行后，A4 应获得下边框 = thin。
        var te = TestExcel.Create();
        var borderId = te.Border(Edge.Top, BorderStyleValues.Thin, rgb: "FF000000");
        var xfA5 = te.CellFormat(borderId: borderId, applyBorder: true);
        te.SetNumber(4, 1, 4);           // A4：默认无边框
        te.SetNumberWithStyle(5, 1, 5, xfA5); // A5：上边框 thin
        te.SetNumber(6, 1, 6);           // A6：占位
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 5, 5);
        Assert.True(r.Success, r.ErrorMessage);

        var edge = TestExcel.GetEdge(path, "Sheet1", "A4", Edge.Bottom);
        Assert.True(edge.HasValue);
        Assert.Equal(BorderStyleValues.Thin, edge!.Value.Style);

        // 下方存活单元格（A6 -> A5）原本无上边框，A5(被删)的下边框也为空，故不应新增上边框
        Assert.Null(TestExcel.GetEdge(path, "Sheet1", "A5", Edge.Top));
    }

    [Fact]
    public void DeleteRow_KeepsExistingNeighborBorderWhenNotEmpty()
    {
        // A4 已有下边框(medium)，A5 有上边框(thin)。删除第5行后 A4 下边框应保持 medium。
        var te = TestExcel.Create();
        var borderMedium = te.Border(Edge.Bottom, BorderStyleValues.Medium, rgb: "FF000000");
        var borderThin = te.Border(Edge.Top, BorderStyleValues.Thin, rgb: "FF000000");
        var xfA4 = te.CellFormat(borderId: borderMedium, applyBorder: true);
        var xfA5 = te.CellFormat(borderId: borderThin, applyBorder: true);
        te.SetNumberWithStyle(4, 1, 4, xfA4);
        te.SetNumberWithStyle(5, 1, 5, xfA5);
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 5, 5);
        Assert.True(r.Success, r.ErrorMessage);

        var edge = TestExcel.GetEdge(path, "Sheet1", "A4", Edge.Bottom);
        Assert.True(edge.HasValue);
        Assert.Equal(BorderStyleValues.Medium, edge!.Value.Style);
    }

    [Fact]
    public void DeleteColumn_TransfersDeletedLeftBorderToLeftNeighbor()
    {
        // D1 无边框，E1 有左边框(thin)。删除 E 列后，D1 应获得右边框 = thin。
        var te = TestExcel.Create();
        var borderId = te.Border(Edge.Left, BorderStyleValues.Thin, rgb: "FF000000");
        var xfE = te.CellFormat(borderId: borderId, applyBorder: true);
        te.SetNumber(1, 4, 40);                 // D1：无边框
        te.SetNumberWithStyle(1, 5, 50, xfE);   // E1：左边框
        te.SetNumber(1, 6, 60);                 // F1 占位
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteColumns(path, "Sheet1", 5, 5);
        Assert.True(r.Success, r.ErrorMessage);

        var edge = TestExcel.GetEdge(path, "Sheet1", "D1", Edge.Right);
        Assert.True(edge.HasValue);
        Assert.Equal(BorderStyleValues.Thin, edge!.Value.Style);
    }
}
