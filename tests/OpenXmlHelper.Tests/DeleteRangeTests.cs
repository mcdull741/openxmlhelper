using static OpenXmlHelper.OpenXmlHelper;
using Xunit;

namespace OpenXmlHelper.Tests;

/// <summary>验证自定义矩形区域的删除与位移。</summary>
public class DeleteRangeTests
{
    [Fact]
    public void DeleteRange_ShiftUp_MovesBelowCellsUpAndUpdatesFormula()
    {
        var te = TestExcel.Create();
        for (int i = 1; i <= 5; i++) te.SetNumber(i, 3, i); // C1=1..C5=5
        te.SetFormula(1, 4, "SUM(C1:C5)");                  // D1
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRange(path, "Sheet1", "C2", "C3",
            ShiftDirection.Up);
        Assert.True(r.Success, r.ErrorMessage);

        // C2,C3 被删除，C4->C2, C5->C3
        Assert.Equal("4", TestExcel.GetValue(path, "Sheet1", "C2"));
        Assert.Equal("5", TestExcel.GetValue(path, "Sheet1", "C3"));
        Assert.Null(TestExcel.GetValue(path, "Sheet1", "C4"));
        Assert.Null(TestExcel.GetValue(path, "Sheet1", "C5"));
        // 公式区域收缩
        Assert.Equal("SUM(C1:C3)", TestExcel.GetFormula(path, "Sheet1", "D1"));
    }

    [Fact]
    public void DeleteRange_ShiftLeft_MovesRightCellsLeftAndUpdatesFormula()
    {
        var te = TestExcel.Create();
        // 第1行：A=1 B=2 C=3 D=4 E=5
        for (int c = 1; c <= 5; c++) te.SetNumber(1, c, c);
        te.SetFormula(2, 1, "SUM(A1:E1)"); // A2
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRange(path, "Sheet1", "B1", "C1",
            ShiftDirection.Left);
        Assert.True(r.Success, r.ErrorMessage);

        // B,C 被删除，D->B, E->C
        Assert.Equal("4", TestExcel.GetValue(path, "Sheet1", "B1"));
        Assert.Equal("5", TestExcel.GetValue(path, "Sheet1", "C1"));
        Assert.Null(TestExcel.GetValue(path, "Sheet1", "D1"));
        Assert.Equal("SUM(A1:C1)", TestExcel.GetFormula(path, "Sheet1", "A2"));
    }

    [Fact]
    public void DeleteRange_ShiftUp_DeletedReferenceBecomesRefError()
    {
        var te = TestExcel.Create();
        te.SetNumber(2, 3, 9); // C2
        te.SetFormula(1, 1, "C2"); // A1 引用 C2
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRange(path, "Sheet1", "C2", "C2",
            ShiftDirection.Up);
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Equal("#REF!", TestExcel.GetFormula(path, "Sheet1", "A1"));
    }

    [Fact]
    public void DeleteRange_DefaultDirectionIsUp()
    {
        var te = TestExcel.Create();
        te.SetNumber(2, 3, 7);
        te.SetNumber(4, 3, 14);
        te.Save(); var path = te.FilePath; te.Dispose();

        // 不传方向参数，默认 Up
        var r = DeleteRange(path, "Sheet1", "C2", "C3");
        Assert.True(r.Success, r.ErrorMessage);

        Assert.Equal("14", TestExcel.GetValue(path, "Sheet1", "C2"));
    }
}
