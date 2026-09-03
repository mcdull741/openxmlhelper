using System.Globalization;
using System.Text.RegularExpressions;

namespace OpenXmlHelper.Helpers;

/// <summary>
/// 使用正则扫描公式中的 A1 引用，并依据 <see cref="ReferenceAdjuster"/> 调整之。
///
/// 支持的引用形式：
/// - 单元格：A1、$A$1
/// - 单元格区域：A1:B10、Sheet1!A1:B10、'My Sheet'!$A$1:$B$10
/// - 整列：A:A、A:C、Sheet1!A:A
/// - 整行：1:1、1:10、Sheet1!1:10
///
/// 限制（已在注释中说明）：
/// - 不解析定义名称（Defined Name）。形如 "Data1" 的名称若恰好符合单元格引用格式（1-3 字母+数字且列号≤16384）
///   会被误判为引用。这是无名称表时的固有歧义；常见公式（SUM(A1:A10) 等）不受影响。
/// - 删除整区域后的 #REF! 会保留原工作表前缀（如 Sheet1!#REF!），与 Excel 表现一致。
/// </summary>
public sealed class FormulaReferenceUpdater
{
    private readonly ReferenceAdjuster _adjuster;

    // 查找引用 token。顺序：单元格区域 > 整列区域 > 整行区域 > 单格。
    // (?<![A-Za-z0-9_]) / (?![A-Za-z0-9_]) 避免匹配到标识符内部。
    private const string Pattern = @"
(?<![A-Za-z0-9_])
(?<sheet>'[^']*'!|[A-Za-z][A-Za-z0-9_.]*!)?
(?:
    (?<c1>\$?[A-Z]{1,3})(?<r1>\$?\d{1,7}):(?<c2>\$?[A-Z]{1,3})(?<r2>\$?\d{1,7})
  | (?<fc1>\$?[A-Z]{1,3}):(?<fc2>\$?[A-Z]{1,3})
  | (?<fr1>\$?\d{1,7}):(?<fr2>\$?\d{1,7})
  | (?<c1>\$?[A-Z]{1,3})(?<r1>\$?\d{1,7})
)
(?![A-Za-z0-9_])";

    private static readonly Regex RefRegex = new(
        Pattern,
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace | RegexOptions.ExplicitCapture,
        TimeSpan.FromSeconds(5));

    public FormulaReferenceUpdater(ReferenceAdjuster adjuster) => _adjuster = adjuster;

    /// <summary>更新公式文本。返回更新后的公式；若原公式为空则原样返回。</summary>
    public string UpdateFormula(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula)) return formula ?? string.Empty;
        return RefRegex.Replace(formula, Evaluate);
    }

    /// <summary>更新共享/数组公式的 ref 区域属性。返回是否成功更新（非删除时为 true）。</summary>
    public bool TryUpdateRangeRef(string? refAttr, out string newRef)
    {
        newRef = refAttr ?? string.Empty;
        if (string.IsNullOrWhiteSpace(refAttr)) return false;
        var updated = RefRegex.Replace(refAttr, Evaluate);
        if (updated.Contains("#REF!")) return false; // 整个区域被删除，ref 失效
        newRef = updated;
        return true;
    }

    private string Evaluate(Match m)
    {
        var sheet = m.Groups["sheet"].Value;
        string result;

        if (m.Groups["c2"].Success) // 单元格区域
        {
            if (!TryParseCellPart(m.Groups["c1"].Value, out int c1abs, out int c1) ||
                !TryParseRowPart(m.Groups["r1"].Value, out int r1abs, out int r1) ||
                !TryParseCellPart(m.Groups["c2"].Value, out int c2abs, out int c2) ||
                !TryParseRowPart(m.Groups["r2"].Value, out int r2abs, out int r2))
                return m.Value;

            Normalize(ref c1, ref c2, ref r1, ref r2, ref c1abs, ref c2abs, ref r1abs, ref r2abs);

            if (!_adjuster.Adjust(c1, r1, c2, r2, out int nc1, out int nr1, out int nc2, out int nr2))
                return sheet + "#REF!";

            result = ToCell(nc1, nr1, c1abs, r1abs) + ":" + ToCell(nc2, nr2, c2abs, r2abs);
        }
        else if (m.Groups["fc1"].Success) // 整列区域
        {
            if (!TryParseCellPart(m.Groups["fc1"].Value, out int c1abs, out int c1) ||
                !TryParseCellPart(m.Groups["fc2"].Value, out int c2abs, out int c2))
                return m.Value;
            NormalizeCols(ref c1, ref c2, ref c1abs, ref c2abs);

            // 整列：行视为 [1, MaxRows]
            if (!_adjuster.Adjust(c1, 1, c2, ReferenceAdjuster.MaxRows, out int nc1, out _, out int nc2, out _))
                return sheet + "#REF!";
            result = ToCol(nc1, c1abs) + ":" + ToCol(nc2, c2abs);
        }
        else if (m.Groups["fr1"].Success) // 整行区域
        {
            if (!TryParseRowPart(m.Groups["fr1"].Value, out int r1abs, out int r1) ||
                !TryParseRowPart(m.Groups["fr2"].Value, out int r2abs, out int r2))
                return m.Value;
            NormalizeRows(ref r1, ref r2, ref r1abs, ref r2abs);

            if (!_adjuster.Adjust(1, r1, ReferenceAdjuster.MaxColumns, r2, out _, out int nr1, out _, out int nr2))
                return sheet + "#REF!";
            result = ToRow(nr1, r1abs) + ":" + ToRow(nr2, r2abs);
        }
        else if (m.Groups["c1"].Success) // 单格
        {
            if (!TryParseCellPart(m.Groups["c1"].Value, out int c1abs, out int c1) ||
                !TryParseRowPart(m.Groups["r1"].Value, out int r1abs, out int r1))
                return m.Value;

            if (!_adjuster.Adjust(c1, r1, c1, r1, out int nc1, out int nr1, out _, out _))
                return sheet + "#REF!";
            result = ToCell(nc1, nr1, c1abs, r1abs);
        }
        else
        {
            return m.Value;
        }

        return sheet + result;
    }

    private static bool TryParseCellPart(string text, out int abs, out int column)
    {
        abs = 0; column = 0;
        if (string.IsNullOrEmpty(text)) return false;
        int i = 0;
        if (text[0] == '$') { abs = 1; i++; }
        int col = CellReference.ToColumnIndex(text.AsSpan(i).ToString().ToUpperInvariant());
        if (col < 1 || col > 16384) return false;
        column = col;
        return true;
    }

    private static bool TryParseRowPart(string text, out int abs, out int row)
    {
        abs = 0; row = 0;
        if (string.IsNullOrEmpty(text)) return false;
        int i = 0;
        if (text[0] == '$') { abs = 1; i++; }
        if (!int.TryParse(text.AsSpan(i), NumberStyles.None, CultureInfo.InvariantCulture, out int r) || r < 1 || r > 1048576)
            return false;
        row = r;
        return true;
    }

    private static string ToCell(int col, int row, int colAbs, int rowAbs)
        => (colAbs == 1 ? "$" : "") + CellReference.ToColumnLetters(col) + (rowAbs == 1 ? "$" : "") + row;

    private static string ToCol(int col, int colAbs) => (colAbs == 1 ? "$" : "") + CellReference.ToColumnLetters(col);
    private static string ToRow(int row, int rowAbs) => (rowAbs == 1 ? "$" : "") + row;

    private static void Normalize(ref int c1, ref int c2, ref int r1, ref int r2, ref int c1a, ref int c2a, ref int r1a, ref int r2a)
    {
        if (c1 > c2) { (c1, c2) = (c2, c1); (c1a, c2a) = (c2a, c1a); }
        if (r1 > r2) { (r1, r2) = (r2, r1); (r1a, r2a) = (r2a, r1a); }
    }
    private static void NormalizeCols(ref int c1, ref int c2, ref int c1a, ref int c2a)
    {
        if (c1 > c2) { (c1, c2) = (c2, c1); (c1a, c2a) = (c2a, c1a); }
    }
    private static void NormalizeRows(ref int r1, ref int r2, ref int r1a, ref int r2a)
    {
        if (r1 > r2) { (r1, r2) = (r2, r1); (r1a, r2a) = (r2a, r1a); }
    }
}
