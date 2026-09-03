using static OpenXmlHelper.OpenXmlHelper;
using Xunit;

namespace OpenXmlHelper.Tests;

/// <summary>验证异常处理：文件/工作表/范围无效等情况返回失败状态及错误信息。</summary>
public class ErrorHandlingTests
{
    [Fact]
    public void DeleteRows_FileNotExist_ReturnsFailure()
    {
        var r = DeleteRows("Z:\\no_such_file.xlsx", "Sheet1", 1, 1);
        Assert.False(r.Success);
        Assert.Contains("文件不存在", r.ErrorMessage);
    }

    [Fact]
    public void DeleteRows_SheetNotExist_ReturnsFailure()
    {
        var path = CreateMinimalFile();
        var r = DeleteRows(path, "NotASheet", 1, 1);
        Assert.False(r.Success);
        Assert.Contains("工作表", r.ErrorMessage!);
    }

    [Fact]
    public void DeleteRows_InvalidRange_ReturnsFailure()
    {
        var path = CreateMinimalFile();
        var r = DeleteRows(path, "Sheet1", 0, 1);
        Assert.False(r.Success);
        Assert.Contains("无效的行范围", r.ErrorMessage);
    }

    [Fact]
    public void DeleteColumns_InvalidRange_ReturnsFailure()
    {
        var path = CreateMinimalFile();
        var r = DeleteColumns(path, "Sheet1", 5, 1);
        Assert.False(r.Success);
        Assert.Contains("无效的列范围", r.ErrorMessage);
    }

    [Fact]
    public void DeleteRange_InvalidCell_ReturnsFailure()
    {
        var path = CreateMinimalFile();
        var r = DeleteRange(path, "Sheet1", "A0", "B2");
        Assert.False(r.Success);
        Assert.Contains("无效", r.ErrorMessage);
    }

    [Fact]
    public void DeleteRange_EmptyArgs_ReturnsFailure()
    {
        var path = CreateMinimalFile();
        var r = DeleteRange(path, "Sheet1", "", "B2");
        Assert.False(r.Success);
        Assert.Contains("不能为空", r.ErrorMessage);
    }

    [Fact]
    public void DeleteRows_OnValidFile_ReturnsSuccess()
    {
        var te = TestExcel.Create();
        for (int i = 1; i <= 3; i++) te.SetNumber(i, 1, i);
        te.Save(); var path = te.FilePath; te.Dispose();

        var r = DeleteRows(path, "Sheet1", 2, 2);
        Assert.True(r.Success, r.ErrorMessage);
    }

    private static string CreateMinimalFile()
    {
        var te = TestExcel.Create();
        te.SetNumber(1, 1, 1);
        te.SetNumber(2, 1, 2);
        te.Save(); var path = te.FilePath; te.Dispose();
        return path;
    }
}
