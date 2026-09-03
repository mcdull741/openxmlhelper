using System.Globalization;
using System.Text;

namespace OpenXmlHelper.Helpers;

/// <summary>
/// 解析与格式化 A1 表示法的单元格引用。
/// 处理列字母与列号互转、绝对引用标记（$）的解析。
/// </summary>
public static class CellReference
{
    private const int MaxColumns = 16384;   // XFD
    private const int MaxRows = 1048576;

    /// <summary>将 1 基列号转换为列字母（1 -> A, 27 -> AA）。</summary>
    public static string ToColumnLetters(int columnIndex)
    {
        if (columnIndex < 1 || columnIndex > MaxColumns)
            throw new ArgumentOutOfRangeException(nameof(columnIndex));

        var sb = new StringBuilder();
        while (columnIndex > 0)
        {
            columnIndex--; // 转为 0 基
            sb.Insert(0, (char)('A' + columnIndex % 26));
            columnIndex /= 26;
        }
        return sb.ToString();
    }

    /// <summary>将列字母转换为 1 基列号（A -> 1, AA -> 27）。</summary>
    public static int ToColumnIndex(ReadOnlySpan<char> columnLetters)
    {
        int index = 0;
        foreach (var ch in columnLetters)
        {
            if (ch < 'A' || ch > 'Z')
                throw new FormatException($"无效的列字母: {ch}");
            index = index * 26 + (ch - 'A' + 1);
        }
        return index;
    }

    /// <summary>尝试解析 A1 引用，返回列号、行号及绝对标记。解析失败返回 false。</summary>
    public static bool TryParse(string reference, out int columnIndex, out int rowIndex,
        out bool columnAbsolute, out bool rowAbsolute)
    {
        columnIndex = rowIndex = 0;
        columnAbsolute = rowAbsolute = false;

        if (string.IsNullOrWhiteSpace(reference))
            return false;

        var span = reference.AsSpan().Trim();
        int i = 0;

        if (i < span.Length && span[i] == '$') { columnAbsolute = true; i++; }

        int colStart = i;
        while (i < span.Length && span[i] >= 'A' && span[i] <= 'Z') i++;
        if (i == colStart) return false;

        int col = ToColumnIndex(span.Slice(colStart, i - colStart).ToString().ToUpperInvariant());
        if (col < 1 || col > MaxColumns) return false;

        if (i < span.Length && span[i] == '$') { rowAbsolute = true; i++; }

        int rowStart = i;
        while (i < span.Length && span[i] >= '0' && span[i] <= '9') i++;
        if (i == rowStart) return false;

        if (int.TryParse(span.Slice(rowStart, i - rowStart), NumberStyles.None, CultureInfo.InvariantCulture, out int row)
            && row >= 1 && row <= MaxRows && i == span.Length)
        {
            columnIndex = col;
            rowIndex = row;
            return true;
        }
        return false;
    }

    /// <summary>组合列号、行号及绝对标记为 A1 引用字符串。</summary>
    public static string ToString(int columnIndex, int rowIndex, bool columnAbsolute, bool rowAbsolute)
        => $"{(columnAbsolute ? "$" : "")}{ToColumnLetters(columnIndex)}{(rowAbsolute ? "$" : "")}{rowIndex}";
}
