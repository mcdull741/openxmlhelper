using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlHelper.Helpers;

namespace OpenXmlHelper.Tests;

/// <summary>
/// 测试用：以最小结构手工构造 .xlsx 文件，并按需添加带样式的单元格/公式/合并区域。
/// 同时提供读取断言辅助方法。
/// </summary>
public sealed class TestExcel : IDisposable
{
    private readonly string _path;
    private readonly SpreadsheetDocument _doc;
    private readonly WorkbookPart _wbPart;
    private readonly WorksheetPart _wsPart;
    private readonly SheetData _sheetData;
    private readonly Stylesheet _ss;
    private readonly Dictionary<string, uint> _fonts = new();
    private readonly Dictionary<string, uint> _borders = new();
    private readonly Dictionary<string, uint> _xfs = new();

    private TestExcel(string path, SpreadsheetDocument doc, WorkbookPart wb, WorksheetPart ws,
        SheetData sd, Stylesheet ss)
    {
        _path = path; _doc = doc; _wbPart = wb; _wsPart = ws; _sheetData = sd; _ss = ss;
    }

    public string FilePath => _path;
    public string SheetName { get; private set; } = "Sheet1";

    public static TestExcel Create(string sheetName = "Sheet1")
    {
        var path = Path.Combine(Path.GetTempPath(), $"oxh_{Guid.NewGuid():N}.xlsx");
        var doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        wsPart.Worksheet = new Worksheet(sheetData);

        var sheets = wbPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet { Id = wbPart.GetIdOfPart(wsPart), SheetId = 1, Name = sheetName });

        var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
        var ss = new Stylesheet();
        // 默认字体/填充/边框/格式（索引0）
        ss.AppendChild(new Fonts(new Font()));
        ss.AppendChild(new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 })));
        ss.AppendChild(new Borders(new Border()));
        ss.AppendChild(new CellStyleFormats(new CellFormat()));
        ss.AppendChild(new CellFormats(new CellFormat()));
        stylesPart.Stylesheet = ss;

        var te = new TestExcel(path, doc, wbPart, wsPart, sheetData, ss) { SheetName = sheetName };
        return te;
    }

    // ---------- 写入辅助 ----------

    public void SetNumber(int row, int col, double value, uint? styleIndex = null)
        => SetCell(row, col, new CellValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture)), null, styleIndex);

    public void SetFormula(int row, int col, string formula, uint? styleIndex = null)
        => SetCell(row, col, null, new CellFormula(formula), styleIndex);

    public void SetNumberWithStyle(int row, int col, double value, uint styleIndex)
        => SetCell(row, col, new CellValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture)), null, styleIndex);

    private void SetCell(int row, int col, CellValue? value, CellFormula? formula, uint? styleIndex)
    {
        var rowEl = GetOrCreateRow(row);
        var cell = new Cell { CellReference = CellReference.ToColumnLetters(col) + row };
        if (value is not null) cell.AppendChild(value);
        if (formula is not null) cell.AppendChild(formula);
        if (styleIndex.HasValue) cell.StyleIndex = styleIndex.Value;
        rowEl.AppendChild(cell);
    }

    private Row GetOrCreateRow(int row)
    {
        foreach (var r in _sheetData.Elements<Row>())
            if (r.RowIndex?.Value == (uint)row) return r;
        var newRow = new Row { RowIndex = (uint)row };
        _sheetData.AppendChild(newRow);
        return newRow;
    }

    /// <summary>创建/复用字体，返回字体索引（0 基）。</summary>
    public uint Font(bool bold = false, string? rgb = null)
    {
        string key = $"b{bold}c{rgb}";
        if (_fonts.TryGetValue(key, out uint f)) return f;
        var font = new Font();
        if (bold) font.AppendChild(new Bold());
        if (rgb is not null) font.AppendChild(new Color { Rgb = rgb });
        var fonts = _ss.Fonts!;
        f = (uint)fonts.ChildElements.Count;
        fonts.AppendChild(font);
        fonts.Count = (uint)fonts.ChildElements.Count;
        _fonts[key] = f;
        return f;
    }

    /// <summary>创建/复用边框，返回边框索引（0 基）。</summary>
    public uint Border(Edge edge, BorderStyleValues style, string? rgb = null)
    {
        string key = $"{edge}:{style}:{rgb}";
        if (_borders.TryGetValue(key, out uint b)) return b;
        var border = new Border();
        BorderPropertiesType e = edge switch
        {
            Edge.Left => new LeftBorder(),
            Edge.Top => new TopBorder(),
            Edge.Right => new RightBorder(),
            Edge.Bottom => new BottomBorder(),
            _ => throw new ArgumentOutOfRangeException(nameof(edge))
        };
        e.Style = style;
        if (rgb is not null) e.Color = new Color { Rgb = rgb };
        switch (edge)
        {
            case Edge.Left: border.LeftBorder = (LeftBorder)e; break;
            case Edge.Right: border.RightBorder = (RightBorder)e; break;
            case Edge.Top: border.TopBorder = (TopBorder)e; break;
            case Edge.Bottom: border.BottomBorder = (BottomBorder)e; break;
        }
        var borders = _ss.Borders!;
        b = (uint)borders.ChildElements.Count;
        borders.AppendChild(border);
        borders.Count = (uint)borders.ChildElements.Count;
        _borders[key] = b;
        return b;
    }

    /// <summary>创建/复用单元格格式（xf），返回其索引（0 基）。</summary>
    public uint CellFormat(uint? fontId = null, uint? borderId = null, bool applyFont = false, bool applyBorder = false)
    {
        string key = $"f{fontId}b{borderId}af{applyFont}ab{applyBorder}";
        if (_xfs.TryGetValue(key, out uint x)) return x;
        var xf = new CellFormat();
        if (fontId.HasValue) { xf.FontId = fontId.Value; xf.ApplyFont = applyFont; }
        if (borderId.HasValue) { xf.BorderId = borderId.Value; xf.ApplyBorder = applyBorder; }
        var cellFormats = _ss.CellFormats!;
        x = (uint)cellFormats.ChildElements.Count;
        cellFormats.AppendChild(xf);
        cellFormats.Count = (uint)cellFormats.ChildElements.Count;
        _xfs[key] = x;
        return x;
    }

    public void Merge(string range)
    {
        var ws = _wsPart.Worksheet;
        var mc = ws.GetFirstChild<MergeCells>();
        if (mc is null)
        {
            mc = new MergeCells();
            // MergeCells 必须在 SheetData 之后、特定位置；直接追加通常可行
            ws.AppendChild(mc);
        }
        mc.AppendChild(new MergeCell { Reference = range });
        mc.Count = (uint)mc.ChildElements.Count;
    }

    /// <summary>保存并关闭，返回文件路径。</summary>
    public string Save()
    {
        _ss.Fonts!.Count = (uint)_ss.Fonts!.ChildElements.Count;
        _ss.Fills!.Count = (uint)_ss.Fills!.ChildElements.Count;
        _ss.Borders!.Count = (uint)_ss.Borders!.ChildElements.Count;
        _ss.CellFormats!.Count = (uint)_ss.CellFormats!.ChildElements.Count;
        _wbPart.Workbook.Save();
        _wsPart.Worksheet.Save();
        _wbPart.WorkbookStylesPart!.Stylesheet.Save();
        return _path;
    }

    public void Dispose()
    {
        _doc.Dispose();
    }

    // ---------- 读取辅助（静态，针对已保存文件） ----------

    public static Cell? GetCell(string path, string sheet, string cellRef)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var wb = doc.WorkbookPart!;
        var sheetEl = wb.Workbook.Sheets!.Elements<Sheet>().FirstOrDefault(s => s.Name?.Value == sheet)
            ?? throw new InvalidOperationException($"工作表 {sheet} 不存在");
        var ws = (WorksheetPart)wb.GetPartById(sheetEl.Id!.Value);
        foreach (var row in ws.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>())
            foreach (var c in row.Elements<Cell>())
                if (c.CellReference?.Value == cellRef) return (Cell)c.CloneNode(true);
        return null;
    }

    public static string? GetFormula(string path, string sheet, string cellRef)
        => GetCell(path, sheet, cellRef)?.CellFormula?.Text;

    public static string? GetValue(string path, string sheet, string cellRef)
        => GetCell(path, sheet, cellRef)?.CellValue?.Text;

    public static BorderEdge? GetEdge(string path, string sheet, string cellRef, Edge edge)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var wb = doc.WorkbookPart!;
        var sheetEl = wb.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name?.Value == sheet);
        var ws = (WorksheetPart)wb.GetPartById(sheetEl.Id!.Value);
        Cell? target = null;
        foreach (var row in ws.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>())
            foreach (var c in row.Elements<Cell>())
                if (c.CellReference?.Value == cellRef) target = c;
        if (target is null) return null;
        var ss = wb.WorkbookStylesPart?.Stylesheet;
        if (ss is null) return null;
        return new BorderConflictResolver(ss).GetBorderEdge(target, edge);
    }

    public static (bool bold, string? rgb) GetFont(string path, string sheet, string cellRef)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var wb = doc.WorkbookPart!;
        var sheetEl = wb.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name?.Value == sheet);
        var ws = (WorksheetPart)wb.GetPartById(sheetEl.Id!.Value);
        Cell? target = null;
        foreach (var row in ws.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>())
            foreach (var c in row.Elements<Cell>())
                if (c.CellReference?.Value == cellRef) target = c;
        if (target?.StyleIndex is null) return (false, null);
        var ss = wb.WorkbookStylesPart?.Stylesheet!;
        var xf = ss.CellFormats!.ChildElements[(int)target.StyleIndex.Value] as CellFormat;
        if (xf?.FontId is null) return (false, null);
        var font = ss.Fonts!.ChildElements[(int)xf.FontId.Value] as Font;
        return (font?.Bold != null, font?.Color?.Rgb?.Value);
    }

    public static List<string> GetMerges(string path, string sheet)
    {
        using var doc = SpreadsheetDocument.Open(path, false);
        var wb = doc.WorkbookPart!;
        var sheetEl = wb.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name?.Value == sheet);
        var ws = (WorksheetPart)wb.GetPartById(sheetEl.Id!.Value);
        var mc = ws.Worksheet.GetFirstChild<MergeCells>();
        if (mc is null) return new List<string>();
        return mc.Elements<MergeCell>().Select(m => m.Reference?.Value ?? "").ToList();
    }
}
