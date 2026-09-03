using static OpenXmlHelper.OpenXmlHelper;
using Xunit;

namespace OpenXmlHelper.Tests;

/// <summary>验证包含合并单元格的区域删除场景。</summary>
public class MergeCellsTests
{
    [Fact]
    public void DeleteRow_ShrinksMergeSpanningDeletedRow()
    {
        var te = TestExcel.Create();
        for (int i = 1; i <= 6; i++) te.SetNumber(i, 1, i);
        te.Merge("A1:A5");
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 3, 3);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Equal(new[] { "A1:A4" }, TestExcel.GetMerges(path, "Sheet1"));
    }

    [Fact]
    public void DeleteRows_ShrinksMergeAcrossMultipleRows()
    {
        var te = TestExcel.Create();
        for (int i = 1; i <= 12; i++) te.SetNumber(i, 1, i);
        te.Merge("A1:A10");
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 3, 4);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Equal(new[] { "A1:A8" }, TestExcel.GetMerges(path, "Sheet1"));
    }

    [Fact]
    public void DeleteRow_RemovesMergeWhenAnchorDeleted()
    {
        var te = TestExcel.Create();
        for (int i = 1; i <= 6; i++) te.SetNumber(i, 1, i);
        te.Merge("A3:A5");   // 锚点 A3
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 3, 3);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Empty(TestExcel.GetMerges(path, "Sheet1"));
    }

    [Fact]
    public void DeleteColumn_ShrinksMergeSpanningDeletedColumn()
    {
        var te = TestExcel.Create();
        te.SetNumber(1, 1, 1);
        te.SetNumber(1, 2, 2);
        te.SetNumber(1, 3, 3);
        te.SetNumber(1, 4, 4);
        te.Merge("A1:C1");
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteColumns(path, "Sheet1", 2, 2);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Equal(new[] { "A1:B1" }, TestExcel.GetMerges(path, "Sheet1"));
    }

    [Fact]
    public void DeleteRange_RemovesMergeFullyInsideDeletedRect()
    {
        var te = TestExcel.Create();
        for (int i = 1; i <= 6; i++) te.SetNumber(i, 3, i); // C1:C6
        te.Merge("C2:C3");
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRange(path, "Sheet1", "C2", "C3",
            ShiftDirection.Up);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Empty(TestExcel.GetMerges(path, "Sheet1"));
        // C4 -> C2, C5 -> C3
        Assert.Equal("4", TestExcel.GetValue(path, "Sheet1", "C2"));
        Assert.Equal("5", TestExcel.GetValue(path, "Sheet1", "C3"));
    }
}
