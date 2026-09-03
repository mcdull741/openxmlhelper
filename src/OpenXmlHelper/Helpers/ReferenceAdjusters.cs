using OpenXmlHelper.Helpers;

namespace OpenXmlHelper.Helpers;

/// <summary>
/// 引用调整器的抽象基类。给定一个引用范围（c1,r1,c2,r2，均为 1 基），
/// 返回删除/位移操作后该引用应变换为的范围；若整个引用被删除则标记为 #REF!。
///
/// 设计说明：
/// - 删除行/列时，相对引用与绝对引用（带 $）都会按相同规则位移，
///   这与 Excel 的删除行为一致（$ 仅在复制/填充时锁定，不影响插入/删除的调整）。
/// - 对于范围引用，当删除区域位于范围内部时范围会收缩，而非整体变 #REF!。
///   仅当范围完全被删除区域覆盖时才变为 #REF!。
/// </summary>
public abstract class ReferenceAdjuster
{
    /// <summary>最大行数（用于将“整列”引用表示为 (1, 1, col, MaxRows)）。</summary>
    public const int MaxRows = 1048576;
    /// <summary>最大列数（用于将“整行”引用表示为 (1, row, MaxCols, row)）。</summary>
    public const int MaxColumns = 16384;

    /// <summary>
    /// 调整一个引用范围。返回 false 表示引用整体被删除（应替换为 #REF!）。
    /// </summary>
    public abstract bool Adjust(int c1, int r1, int c2, int r2,
        out int nc1, out int nr1, out int nc2, out int nr2);

    /// <summary>构造“保持不变”的结果。</summary>
    protected static bool Keep(int c1, int r1, int c2, int r2,
        out int nc1, out int nr1, out int nc2, out int nr2)
    {
        nc1 = c1; nr1 = r1; nc2 = c2; nr2 = r2;
        return true;
    }

    /// <summary>构造“已删除”的结果（#REF!）。</summary>
    protected static bool Deleted(out int nc1, out int nr1, out int nc2, out int nr2)
    {
        nc1 = nr1 = nc2 = nr2 = 0;
        return false;
    }

    /// <summary>
    /// 对单一轴向（行或列）应用删除规则，返回端点调整后的新值。
    /// dStart/dEnd 为被删除区间的起止（1 基，含端点），count 为删除数量。
    /// isStart 指示该端点是范围的起点（影响夹紧方向）。
    /// </summary>
    protected static int AdjustEndpoint(int endpoint, int dStart, int dEnd, int count, bool isStart)
    {
        if (endpoint < dStart) return endpoint;            // 在删除区之前：不变
        if (endpoint > dEnd) return endpoint - count;      // 在删除区之后：前移
        // 端点落在删除区内：夹紧到删除区边界
        // 起点夹紧到 dStart（删除区下方首个存活单元格位移后的位置）；
        // 终点夹紧到 dStart-1（删除区上方最后一个存活单元格）。
        return isStart ? dStart : dStart - 1;
    }

    /// <summary>判断 [r1,r2] 是否被 [dStart,dEnd] 完全覆盖（整体删除）。</summary>
    protected static bool FullyCovered(int r1, int r2, int dStart, int dEnd)
        => dStart <= r1 && r2 <= dEnd;
}

/// <summary>整行删除调整器：删除 [startRow, endRow] 行，下方行上移。</summary>
public sealed class RowDeletionAdjuster : ReferenceAdjuster
{
    private readonly int _startRow, _endRow, _count;
    public RowDeletionAdjuster(int startRow, int endRow)
    {
        if (startRow < 1 || endRow < startRow) throw new ArgumentException("无效的行删除范围。");
        _startRow = startRow; _endRow = endRow; _count = endRow - startRow + 1;
    }

    public override bool Adjust(int c1, int r1, int c2, int r2,
        out int nc1, out int nr1, out int nc2, out int nr2)
    {
        // 列不变；仅行受影响
        if (FullyCovered(r1, r2, _startRow, _endRow)) return Deleted(out nc1, out nr1, out nc2, out nr2);
        nr1 = AdjustEndpoint(r1, _startRow, _endRow, _count, isStart: true);
        nr2 = AdjustEndpoint(r2, _startRow, _endRow, _count, isStart: false);
        return Keep(c1, nr1, c2, nr2, out nc1, out _, out nc2, out _);
    }
}

/// <summary>整列删除调整器：删除 [startCol, endCol] 列，右侧列左移。</summary>
public sealed class ColumnDeletionAdjuster : ReferenceAdjuster
{
    private readonly int _startCol, _endCol, _count;
    public ColumnDeletionAdjuster(int startCol, int endCol)
    {
        if (startCol < 1 || endCol < startCol) throw new ArgumentException("无效的列删除范围。");
        _startCol = startCol; _endCol = endCol; _count = endCol - startCol + 1;
    }

    public override bool Adjust(int c1, int r1, int c2, int r2,
        out int nc1, out int nr1, out int nc2, out int nr2)
    {
        // 行不变；仅列受影响
        if (FullyCovered(c1, c2, _startCol, _endCol)) return Deleted(out nc1, out nr1, out nc2, out nr2);
        nc1 = AdjustEndpoint(c1, _startCol, _endCol, _count, isStart: true);
        nc2 = AdjustEndpoint(c2, _startCol, _endCol, _count, isStart: false);
        return Keep(nc1, r1, nc2, r2, out _, out nr1, out _, out nr2);
    }
}

/// <summary>
/// 矩形区域“上移”删除调整器：在列区间 [colStart, colEnd]、行区间 [rowStart, rowEnd]
/// 内删除单元格，并将同列区间内下方的单元格上移。
/// 仅影响列落入 [colStart, colEnd] 的引用；列部分重叠的引用保持不变（避免错误拆分）。
/// </summary>
public sealed class RangeShiftUpAdjuster : ReferenceAdjuster
{
    private readonly int _colStart, _colEnd, _rowStart, _rowEnd, _count;
    public RangeShiftUpAdjuster(int colStart, int colEnd, int rowStart, int rowEnd)
    {
        if (colStart < 1 || colEnd < colStart || rowStart < 1 || rowEnd < rowStart)
            throw new ArgumentException("无效的删除区域范围。");
        _colStart = colStart; _colEnd = colEnd;
        _rowStart = rowStart; _rowEnd = rowEnd;
        _count = rowEnd - rowStart + 1;
    }

    public override bool Adjust(int c1, int r1, int c2, int r2,
        out int nc1, out int nr1, out int nc2, out int nr2)
    {
        // 列完全在位移列区间外：不变
        if (c2 < _colStart || c1 > _colEnd)
            return Keep(c1, r1, c2, r2, out nc1, out nr1, out nc2, out nr2);

        // 列完全在位移列区间内：应用行位移规则
        if (c1 >= _colStart && c2 <= _colEnd)
        {
            if (FullyCovered(r1, r2, _rowStart, _rowEnd))
                return Deleted(out nc1, out nr1, out nc2, out nr2);
            nr1 = AdjustEndpoint(r1, _rowStart, _rowEnd, _count, isStart: true);
            nr2 = AdjustEndpoint(r2, _rowStart, _rowEnd, _count, isStart: false);
            return Keep(c1, nr1, c2, nr2, out nc1, out _, out nc2, out _);
        }

        // 列部分重叠：保持不变（不安全拆分）
        return Keep(c1, r1, c2, r2, out nc1, out nr1, out nc2, out nr2);
    }
}

/// <summary>
/// 矩形区域“左移”删除调整器：在行区间 [rowStart, rowEnd]、列区间 [colStart, colEnd]
/// 内删除单元格，并将同行区间内右侧的单元格左移。
/// 仅影响行落入 [rowStart, rowEnd] 的引用；行部分重叠的引用保持不变。
/// </summary>
public sealed class RangeShiftLeftAdjuster : ReferenceAdjuster
{
    private readonly int _colStart, _colEnd, _rowStart, _rowEnd, _count;
    public RangeShiftLeftAdjuster(int colStart, int colEnd, int rowStart, int rowEnd)
    {
        if (colStart < 1 || colEnd < colStart || rowStart < 1 || rowEnd < rowStart)
            throw new ArgumentException("无效的删除区域范围。");
        _colStart = colStart; _colEnd = colEnd;
        _rowStart = rowStart; _rowEnd = rowEnd;
        _count = colEnd - colStart + 1;
    }

    public override bool Adjust(int c1, int r1, int c2, int r2,
        out int nc1, out int nr1, out int nc2, out int nr2)
    {
        // 行完全在区域外：不变
        if (r2 < _rowStart || r1 > _rowEnd)
            return Keep(c1, r1, c2, r2, out nc1, out nr1, out nc2, out nr2);

        // 行完全在区域内：应用列位移规则
        if (r1 >= _rowStart && r2 <= _rowEnd)
        {
            if (FullyCovered(c1, c2, _colStart, _colEnd))
                return Deleted(out nc1, out nr1, out nc2, out nr2);
            nc1 = AdjustEndpoint(c1, _colStart, _colEnd, _count, isStart: true);
            nc2 = AdjustEndpoint(c2, _colStart, _colEnd, _count, isStart: false);
            return Keep(nc1, r1, nc2, r2, out _, out nr1, out _, out nr2);
        }

        // 行部分重叠：保持不变
        return Keep(c1, r1, c2, r2, out nc1, out nr1, out nc2, out nr2);
    }
}
