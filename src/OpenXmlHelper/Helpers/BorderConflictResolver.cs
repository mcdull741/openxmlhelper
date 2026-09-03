using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlHelper.Helpers;

namespace OpenXmlHelper.Helpers;

/// <summary>
/// 边框方向。
/// </summary>
public enum Edge { Left, Top, Right, Bottom }

/// <summary>
/// 一条边框边的快照（样式 + 颜色），与单元格元素解耦，便于在删除前后传递。
/// </summary>
public readonly record struct BorderEdge
{
    public BorderStyleValues Style { get; init; }
    public Color? Color { get; init; }

    public bool IsNone => Style == BorderStyleValues.None;

    /// <summary>从一条边框边元素构造快照。</summary>
    public static BorderEdge? From(BorderPropertiesType? edge)
    {
        if (edge is null) return null;
        var style = edge.Style;
        if (style is null || !style.HasValue || style.Value == BorderStyleValues.None) return null;
        return new BorderEdge { Style = style.Value, Color = edge.Color == null ? null : (Color)edge.Color.CloneNode(true) };
    }

    /// <summary>构造一条带颜色的边框边元素（类型由方向决定）。</summary>
    public BorderPropertiesType ToElement(Edge edge)
    {
        BorderPropertiesType e = edge switch
        {
            Edge.Left => new LeftBorder(),
            Edge.Top => new TopBorder(),
            Edge.Right => new RightBorder(),
            Edge.Bottom => new BottomBorder(),
            _ => throw new ArgumentOutOfRangeException(nameof(edge))
        };
        e.Style = Style;
        if (Color is not null) e.Color = (Color)Color.CloneNode(true);
        return e;
    }
}

/// <summary>
/// 处理删除区域后的边框冲突：
/// 当被删除单元格在朝向相邻存活单元格的边上存在边框，而该存活单元格对应边没有边框时，
/// 将被删除单元格的边框“继承”到存活单元格，从而保留原有的边框样式。
///
/// 样式保留说明：字体、颜色、对齐等格式随单元格位移自动保留（单元格的 s 属性指向不变的 xf）。
/// 唯一需要特殊处理的是“被删除单元格独有、否则会丢失”的边框，因此本类专注于边框。
/// </summary>
public sealed class BorderConflictResolver
{
    private readonly Stylesheet _stylesheet;
    private readonly Dictionary<string, uint> _borderCache = new();
    private readonly Dictionary<string, uint> _xfCache = new();

    public BorderConflictResolver(Stylesheet stylesheet)
    {
        _stylesheet = stylesheet ?? throw new ArgumentNullException(nameof(stylesheet));
        EnsureStyleCollections();
    }

    /// <summary>读取单元格指定方向的边框（null 表示无边框）。</summary>
    public BorderEdge? GetBorderEdge(Cell? cell, Edge edge)
    {
        var border = GetCellBorder(cell);
        if (border is null) return null;
        return BorderEdge.From(GetEdgeElement(border, edge));
    }

    /// <summary>
    /// 若单元格指定方向无边框，则将 <paramref name="source"/> 写入该方向。
    /// 用于把被删除单元格的边框转移到相邻存活单元格。
    /// </summary>
    public void SetBorderEdgeIfAbsent(Cell? cell, Edge edge, BorderEdge? source)
    {
        if (cell is null) return;
        if (source is null || source.Value.IsNone) return;

        // 该方向已有边框则不覆盖（保留存活单元格原有边框）
        if (GetBorderEdge(cell, edge) is not null) return;

        var newBorderId = BuildBorderWithEdge(cell, edge, source.Value);
        var newXfId = BuildCellFormatWithBorder(cell, newBorderId);
        cell.StyleIndex = newXfId;
    }

    // ---------- 内部实现 ----------

    private Border? GetCellBorder(Cell? cell)
    {
        if (cell is null || cell.StyleIndex is null) return null;
        var cellFormats = _stylesheet.CellFormats;
        if (cellFormats is null) return null;
        var xf = cellFormats.ChildElements[(int)cell.StyleIndex.Value] as CellFormat;
        if (xf is null) return null;
        var borders = _stylesheet.Borders;
        if (borders is null || xf.BorderId is null) return null;
        return borders.ChildElements[(int)xf.BorderId.Value] as Border;
    }

    private uint BuildBorderWithEdge(Cell cell, Edge edge, BorderEdge source)
    {
        // 以单元格当前边框为蓝本克隆，再设置目标边
        var currentBorder = GetCellBorder(cell);
        Border newBorder;
        if (currentBorder is not null)
            newBorder = (Border)currentBorder.CloneNode(true);
        else
            newBorder = new Border();

        // 移除目标方向已有的元素（若有），再写入新元素以保证子元素顺序
        RemoveEdge(newBorder, edge);
        InsertEdge(newBorder, edge, source.ToElement(edge));

        string sig = Signature(newBorder);
        if (_borderCache.TryGetValue(sig, out uint existing)) return existing;

        var borders = _stylesheet.Borders!;
        uint id = (uint)borders.ChildElements.Count;
        borders.AppendChild(newBorder);
        _borderCache[sig] = id;
        return id;
    }

    private uint BuildCellFormatWithBorder(Cell cell, uint borderId)
    {
        var cellFormats = _stylesheet.CellFormats!;
        CellFormat newXF;
        if (cell.StyleIndex is not null
            && cellFormats.ChildElements[(int)cell.StyleIndex.Value] is CellFormat current)
        {
            newXF = (CellFormat)current.CloneNode(true);
        }
        else
        {
            // 无样式的单元格：以默认 xf0 为蓝本
            newXF = cellFormats.ChildElements[0] is CellFormat def
                ? (CellFormat)def.CloneNode(true)
                : new CellFormat();
        }
        newXF.BorderId = borderId;
        newXF.ApplyBorder = true;

        string sig = Signature(newXF);
        if (_xfCache.TryGetValue(sig, out uint existing)) return existing;

        uint id = (uint)cellFormats.ChildElements.Count;
        cellFormats.AppendChild(newXF);
        _xfCache[sig] = id;
        return id;
    }

    private static BorderPropertiesType? GetEdgeElement(Border border, Edge edge) => edge switch
    {
        Edge.Left => border.LeftBorder,
        Edge.Top => border.TopBorder,
        Edge.Right => border.RightBorder,
        Edge.Bottom => border.BottomBorder,
        _ => null
    };

    private static void RemoveEdge(Border border, Edge edge)
    {
        var el = GetEdgeElement(border, edge);
        if (el is not null) border.RemoveChild(el);
    }

    // 按规范顺序插入边框边元素：left, right, top, bottom, diagonal
    private static void InsertEdge(Border border, Edge edge, BorderPropertiesType element)
    {
        switch (edge)
        {
            case Edge.Left: border.LeftBorder = (LeftBorder)element; break;
            case Edge.Right: border.RightBorder = (RightBorder)element; break;
            case Edge.Top: border.TopBorder = (TopBorder)element; break;
            case Edge.Bottom: border.BottomBorder = (BottomBorder)element; break;
        }
    }

    private static string Signature(Border border) => border.OuterXml;

    private static string Signature(CellFormat xf) => xf.OuterXml;

    private void EnsureStyleCollections()
    {
        _stylesheet.Borders ??= new Borders();
        if (_stylesheet.Borders.Count() == 0)
            _stylesheet.Borders.AppendChild(new Border());

        _stylesheet.CellFormats ??= new CellFormats();
        if (_stylesheet.CellFormats.Count() == 0)
            _stylesheet.CellFormats.AppendChild(new CellFormat());
    }
}
