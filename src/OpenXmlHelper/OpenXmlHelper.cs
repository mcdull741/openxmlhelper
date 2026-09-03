using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenXmlHelper.Helpers;

namespace OpenXmlHelper;

/// <summary>
/// 基于 OpenXML SDK 的 Excel 区域删除帮助类（静态）。
///
/// 公共 API：
/// <list type="bullet">
/// <item><see cref="DeleteRows"/>：删除指定行区间，下方单元格上移。</item>
/// <item><see cref="DeleteColumns"/>：删除指定列区间，右侧单元格左移。</item>
/// <item><see cref="DeleteRange"/>：删除矩形区域并按指定方向位移相邻单元格。</item>
/// </list>
///
/// 实现要点：
/// <list type="bullet">
/// <item>样式保留：单元格位移时其样式索引（s）不变，字体/填充/对齐/边框随之保留。</item>
/// <item>边框冲突：由 <see cref="BorderConflictResolver"/> 将被删除单元格独有的边框转移到相邻存活单元格。</item>
/// <item>公式重排：由 <see cref="FormulaReferenceUpdater"/> 调整 A1 引用（含相对/绝对、区域收缩、#REF!）。</item>
/// <item>合并单元格：删除区域内或锚点被删除时取消合并，其余按位移调整。</item>
/// <item>异常处理：文件/工作表/范围校验失败返回带错误信息的 <see cref="OperationResult"/>。</item>
/// </list>
///
/// 性能说明：使用 DOM 模式将工作表载入内存，适用于常规文件；
/// 对于极大工作表（数十万行）建议改用 OpenXmlReader/Writer 流式处理。
/// </summary>
public static class OpenXmlHelper
{
    /// <summary>区域删除后的位移方向。</summary>
    public enum ShiftDirection { Up, Left }

    // ---------------- 公共 API ----------------

    /// <summary>
    /// 删除指定工作表中 [startRow, endRow] 的整行，下方行上移。
    /// </summary>
    public static OperationResult DeleteRows(string filePath, string sheetName, int startRow, int endRow)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return OperationResult.Fail("文件路径不能为空。");
        if (!File.Exists(filePath)) return OperationResult.Fail($"文件不存在：{filePath}");
        if (string.IsNullOrWhiteSpace(sheetName)) return OperationResult.Fail("工作表名称不能为空。");
        if (startRow < 1 || endRow < startRow) return OperationResult.Fail($"无效的行范围：{startRow}-{endRow}。");

        try
        {
            using var doc = SpreadsheetDocument.Open(filePath, true);
            var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("工作簿部件缺失。");
            var (wsPart, sheet) = GetWorksheet(wbPart, sheetName);
            if (wsPart is null) return OperationResult.Fail($"工作表 '{sheetName}' 不存在。");

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()
                ?? throw new InvalidOperationException("工作表数据(SheetData)缺失。");
            MaterializeReferences(sheetData);

            var styles = wbPart.WorkbookStylesPart?.Stylesheet;
            var borderResolver = styles is null ? null : new BorderConflictResolver(styles);

            int count = endRow - startRow + 1;

            // 1) 捕获被删除行边界边框（删除前）
            var topEdges = new Dictionary<int, BorderEdge?>();   // 列 -> 删除区顶部单元格的上边框
            var bottomEdges = new Dictionary<int, BorderEdge?>();
            if (borderResolver is not null)
            {
                foreach (var (col, cell) in CellsInRow(sheetData, startRow))
                    topEdges[col] = borderResolver.GetBorderEdge(cell, Edge.Top);
                foreach (var (col, cell) in CellsInRow(sheetData, endRow))
                    bottomEdges[col] = borderResolver.GetBorderEdge(cell, Edge.Bottom);
            }

            // 2) 删除目标行；上移下方行
            var toRemove = new List<Row>();
            foreach (var row in sheetData.Elements<Row>())
            {
                int r = RowIndexOf(row) ?? 0;
                if (r == 0) continue;
                if (r >= startRow && r <= endRow) { toRemove.Add(row); continue; }
                if (r > endRow) ShiftRowDown(row, -count);
            }
            foreach (var row in toRemove) sheetData.RemoveChild(row);

            // 3) 公式重排（全表）
            var updater = new FormulaReferenceUpdater(new RowDeletionAdjuster(startRow, endRow));
            UpdateAllFormulas(sheetData, updater);

            // 4) 合并单元格
            AdjustMergesForRowDeletion(wsPart.Worksheet, startRow, endRow, count);

            // 5) 边框转移（删除后）
            if (borderResolver is not null)
            {
                foreach (var (col, topSrc) in topEdges)
                {
                    var above = GetCell(sheetData, startRow - 1, col);
                    borderResolver.SetBorderEdgeIfAbsent(above, Edge.Bottom, topSrc);
                }
                foreach (var (col, bottomSrc) in bottomEdges)
                {
                    var below = GetCell(sheetData, startRow, col); // 原 endRow+1 已上移至此
                    borderResolver.SetBorderEdgeIfAbsent(below, Edge.Top, bottomSrc);
                }
            }

            // 6) 维度与重算
            RecalculateDimension(wsPart.Worksheet);
            MarkFullRecalc(wbPart);
            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
            return OperationResult.Ok();
        }
        catch (Exception ex)
        {
            return OperationResult.Fail($"删除行失败：{ex.Message}");
        }
    }

    /// <summary>
    /// 删除指定工作表中 [startColumn, endColumn] 的整列，右侧列左移。
    /// </summary>
    public static OperationResult DeleteColumns(string filePath, string sheetName, int startColumn, int endColumn)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return OperationResult.Fail("文件路径不能为空。");
        if (!File.Exists(filePath)) return OperationResult.Fail($"文件不存在：{filePath}");
        if (string.IsNullOrWhiteSpace(sheetName)) return OperationResult.Fail("工作表名称不能为空。");
        if (startColumn < 1 || endColumn < startColumn) return OperationResult.Fail($"无效的列范围：{startColumn}-{endColumn}。");

        try
        {
            using var doc = SpreadsheetDocument.Open(filePath, true);
            var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("工作簿部件缺失。");
            var (wsPart, sheet) = GetWorksheet(wbPart, sheetName);
            if (wsPart is null) return OperationResult.Fail($"工作表 '{sheetName}' 不存在。");

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()
                ?? throw new InvalidOperationException("工作表数据(SheetData)缺失。");
            MaterializeReferences(sheetData);

            var styles = wbPart.WorkbookStylesPart?.Stylesheet;
            var borderResolver = styles is null ? null : new BorderConflictResolver(styles);

            int count = endColumn - startColumn + 1;

            // 1) 捕获被删除列边界边框
            var leftEdges = new Dictionary<int, BorderEdge?>();
            var rightEdges = new Dictionary<int, BorderEdge?>();
            if (borderResolver is not null)
            {
                foreach (var (row, cell) in CellsInColumn(sheetData, startColumn))
                    leftEdges[row] = borderResolver.GetBorderEdge(cell, Edge.Left);
                foreach (var (row, cell) in CellsInColumn(sheetData, endColumn))
                    rightEdges[row] = borderResolver.GetBorderEdge(cell, Edge.Right);
            }

            // 2) 删除目标列单元格；左移右侧单元格
            foreach (var row in sheetData.Elements<Row>().ToList())
            {
                var cellsToRemove = new List<Cell>();
                foreach (var cell in row.Elements<Cell>())
                {
                    var (col, _) = ParseCellRef(cell.CellReference);
                    if (col == 0) continue;
                    if (col >= startColumn && col <= endColumn) { cellsToRemove.Add(cell); continue; }
                    if (col > endColumn) ShiftCellColumn(cell, -count);
                }
                foreach (var cell in cellsToRemove) row.RemoveChild(cell);
            }

            // 3) 公式重排
            var updater = new FormulaReferenceUpdater(new ColumnDeletionAdjuster(startColumn, endColumn));
            UpdateAllFormulas(sheetData, updater);

            // 4) 合并单元格
            AdjustMergesForColumnDeletion(wsPart.Worksheet, startColumn, endColumn, count);

            // 5) 边框转移
            if (borderResolver is not null)
            {
                foreach (var (row, leftSrc) in leftEdges)
                {
                    var left = GetCell(sheetData, row, startColumn - 1);
                    borderResolver.SetBorderEdgeIfAbsent(left, Edge.Right, leftSrc);
                }
                foreach (var (row, rightSrc) in rightEdges)
                {
                    var right = GetCell(sheetData, row, startColumn); // 原 endColumn+1 已左移至此
                    borderResolver.SetBorderEdgeIfAbsent(right, Edge.Left, rightSrc);
                }
            }

            RecalculateDimension(wsPart.Worksheet);
            MarkFullRecalc(wbPart);
            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
            return OperationResult.Ok();
        }
        catch (Exception ex)
        {
            return OperationResult.Fail($"删除列失败：{ex.Message}");
        }
    }

    /// <summary>
    /// 删除矩形区域 [startCell:endCell]，并按 <paramref name="direction"/> 位移相邻单元格。
    /// 仅支持 Up（下方上移）与 Left（右方左移）。
    /// </summary>
    public static OperationResult DeleteRange(string filePath, string sheetName,
        string startCell, string endCell, ShiftDirection direction = ShiftDirection.Up)
    {
        if (string.IsNullOrWhiteSpace(filePath)) return OperationResult.Fail("文件路径不能为空。");
        if (!File.Exists(filePath)) return OperationResult.Fail($"文件不存在：{filePath}");
        if (string.IsNullOrWhiteSpace(sheetName)) return OperationResult.Fail("工作表名称不能为空。");
        if (string.IsNullOrWhiteSpace(startCell) || string.IsNullOrWhiteSpace(endCell))
            return OperationResult.Fail("区域起止单元格不能为空。");

        if (!CellReference.TryParse(startCell, out int c1, out int r1, out _, out _))
            return OperationResult.Fail($"无效的起始单元格：{startCell}");
        if (!CellReference.TryParse(endCell, out int c2, out int r2, out _, out _))
            return OperationResult.Fail($"无效的结束单元格：{endCell}");
        if (c1 > c2) (c1, c2) = (c2, c1);
        if (r1 > r2) (r1, r2) = (r2, r1);

        try
        {
            using var doc = SpreadsheetDocument.Open(filePath, true);
            var wbPart = doc.WorkbookPart ?? throw new InvalidOperationException("工作簿部件缺失。");
            var (wsPart, sheet) = GetWorksheet(wbPart, sheetName);
            if (wsPart is null) return OperationResult.Fail($"工作表 '{sheetName}' 不存在。");

            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>()
                ?? throw new InvalidOperationException("工作表数据(SheetData)缺失。");
            MaterializeReferences(sheetData);

            var styles = wbPart.WorkbookStylesPart?.Stylesheet;
            var borderResolver = styles is null ? null : new BorderConflictResolver(styles);

            if (direction == ShiftDirection.Up)
                DeleteRangeShiftUp(sheetData, wsPart.Worksheet, borderResolver, c1, c2, r1, r2);
            else
                DeleteRangeShiftLeft(sheetData, wsPart.Worksheet, borderResolver, c1, c2, r1, r2);

            RecalculateDimension(wsPart.Worksheet);
            MarkFullRecalc(wbPart);
            wsPart.Worksheet.Save();
            wbPart.Workbook.Save();
            return OperationResult.Ok();
        }
        catch (Exception ex)
        {
            return OperationResult.Fail($"删除区域失败：{ex.Message}");
        }
    }

    // ---------------- DeleteRange 实现 ----------------

    private static void DeleteRangeShiftUp(SheetData sheetData, Worksheet ws,
        BorderConflictResolver? borderResolver, int c1, int c2, int r1, int r2)
    {
        int count = r2 - r1 + 1;

        // 1) 捕获边界边框
        var topEdges = new Dictionary<int, BorderEdge?>();
        var bottomEdges = new Dictionary<int, BorderEdge?>();
        if (borderResolver is not null)
        {
            for (int col = c1; col <= c2; col++)
            {
                topEdges[col] = borderResolver.GetBorderEdge(GetCell(sheetData, r1, col), Edge.Top);
                bottomEdges[col] = borderResolver.GetBorderEdge(GetCell(sheetData, r2, col), Edge.Bottom);
            }
        }

        // 2) 删除区域内单元格；上移区域内列下方的单元格
        var movedCells = new List<(Cell cell, int newRow)>();
        foreach (var row in sheetData.Elements<Row>().ToList())
        {
            int r = RowIndexOf(row) ?? 0;
            if (r == 0) continue;
            var toRemove = new List<Cell>();
            foreach (var cell in row.Elements<Cell>())
            {
                var (col, _) = ParseCellRef(cell.CellReference);
                if (col < c1 || col > c2) continue;          // 仅处理区域列
                if (r >= r1 && r <= r2) { toRemove.Add(cell); continue; }
                if (r > r2) movedCells.Add((cell, r - count));
            }
            foreach (var cell in toRemove) row.RemoveChild(cell);
        }

        // 执行上移（跨行迁移单元格）
        foreach (var (cell, newRow) in movedCells)
        {
            var (col, _) = ParseCellRef(cell.CellReference);
            var srcRow = (Row)cell.Parent!;
            srcRow.RemoveChild(cell);
            cell.CellReference = CellReference.ToColumnLetters(col) + newRow;
            var targetRow = GetOrCreateRow(sheetData, newRow);
            targetRow.AppendChild(cell);
        }

        ReorderSheetData(sheetData);

        // 3) 公式重排
        var updater = new FormulaReferenceUpdater(new RangeShiftUpAdjuster(c1, c2, r1, r2));
        UpdateAllFormulas(sheetData, updater);

        // 4) 合并单元格
        AdjustMergesForRangeShiftUp(ws, c1, c2, r1, r2, count);

        // 5) 边框转移
        if (borderResolver is not null)
        {
            foreach (var (col, topSrc) in topEdges)
            {
                var above = GetCell(sheetData, r1 - 1, col);
                borderResolver.SetBorderEdgeIfAbsent(above, Edge.Bottom, topSrc);
            }
            foreach (var (col, bottomSrc) in bottomEdges)
            {
                var below = GetCell(sheetData, r1, col); // 原 r2+1 已上移至 r1
                borderResolver.SetBorderEdgeIfAbsent(below, Edge.Top, bottomSrc);
            }
        }
    }

    private static void DeleteRangeShiftLeft(SheetData sheetData, Worksheet ws,
        BorderConflictResolver? borderResolver, int c1, int c2, int r1, int r2)
    {
        int count = c2 - c1 + 1;

        // 1) 捕获边界边框
        var leftEdges = new Dictionary<int, BorderEdge?>();
        var rightEdges = new Dictionary<int, BorderEdge?>();
        if (borderResolver is not null)
        {
            for (int row = r1; row <= r2; row++)
            {
                leftEdges[row] = borderResolver.GetBorderEdge(GetCell(sheetData, row, c1), Edge.Left);
                rightEdges[row] = borderResolver.GetBorderEdge(GetCell(sheetData, row, c2), Edge.Right);
            }
        }

        // 2) 删除区域内单元格；左移区域内行右侧的单元格
        foreach (var row in sheetData.Elements<Row>().ToList())
        {
            int r = RowIndexOf(row) ?? 0;
            if (r < r1 || r > r2) continue;                  // 仅处理区域行
            var toRemove = new List<Cell>();
            foreach (var cell in row.Elements<Cell>())
            {
                var (col, _) = ParseCellRef(cell.CellReference);
                if (col == 0) continue;
                if (col >= c1 && col <= c2) { toRemove.Add(cell); continue; }
                if (col > c2) ShiftCellColumn(cell, -count);
            }
            foreach (var cell in toRemove) row.RemoveChild(cell);
        }

        // 3) 公式重排
        var updater = new FormulaReferenceUpdater(new RangeShiftLeftAdjuster(c1, c2, r1, r2));
        UpdateAllFormulas(sheetData, updater);

        // 4) 合并单元格
        AdjustMergesForRangeShiftLeft(ws, c1, c2, r1, r2, count);

        // 5) 边框转移
        if (borderResolver is not null)
        {
            foreach (var (row, leftSrc) in leftEdges)
            {
                var left = GetCell(sheetData, row, c1 - 1);
                borderResolver.SetBorderEdgeIfAbsent(left, Edge.Right, leftSrc);
            }
            foreach (var (row, rightSrc) in rightEdges)
            {
                var right = GetCell(sheetData, row, c1); // 原 c2+1 已左移至此
                borderResolver.SetBorderEdgeIfAbsent(right, Edge.Left, rightSrc);
            }
        }
    }

    // ---------------- 公式更新 ----------------

    private static void UpdateAllFormulas(SheetData sheetData, FormulaReferenceUpdater updater)
    {
        foreach (var cell in sheetData.Descendants<Cell>())
        {
            var f = cell.CellFormula;
            if (f is null) continue;

            // 公式文本
            if (!string.IsNullOrEmpty(f.Text))
                f.Text = updater.UpdateFormula(f.Text);

            // 共享/数组公式主单元格的 ref 属性
            var refAttr = f.Reference;
            if (refAttr is not null && refAttr.HasValue)
            {
                if (updater.TryUpdateRangeRef(refAttr.Value, out string newRef))
                    f.Reference = newRef;
            }
        }
    }

    // ---------------- 合并单元格 ----------------

    private static void AdjustMergesForRowDeletion(Worksheet ws, int startRow, int endRow, int count)
    {
        var mc = ws.GetFirstChild<MergeCells>();
        if (mc is null) return;
        var adjuster = new RowDeletionAdjuster(startRow, endRow);
        var toRemove = new List<MergeCell>();

        foreach (var merge in mc.Elements<MergeCell>())
        {
            var (c1, r1, c2, r2) = ParseRange(merge.Reference);
            // 锚点（左上角）被删除 → 取消合并
            if (r1 >= startRow && r1 <= endRow) { toRemove.Add(merge); continue; }
            if (!adjuster.Adjust(c1, r1, c2, r2, out int nc1, out int nr1, out int nc2, out int nr2))
            { toRemove.Add(merge); continue; }
            merge.Reference = BuildRange(nc1, nr1, nc2, nr2);
        }
        foreach (var m in toRemove) mc.RemoveChild(m);
    }

    private static void AdjustMergesForColumnDeletion(Worksheet ws, int startCol, int endCol, int count)
    {
        var mc = ws.GetFirstChild<MergeCells>();
        if (mc is null) return;
        var adjuster = new ColumnDeletionAdjuster(startCol, endCol);
        var toRemove = new List<MergeCell>();

        foreach (var merge in mc.Elements<MergeCell>())
        {
            var (c1, r1, c2, r2) = ParseRange(merge.Reference);
            if (c1 >= startCol && c1 <= endCol) { toRemove.Add(merge); continue; }
            if (!adjuster.Adjust(c1, r1, c2, r2, out int nc1, out int nr1, out int nc2, out int nr2))
            { toRemove.Add(merge); continue; }
            merge.Reference = BuildRange(nc1, nr1, nc2, nr2);
        }
        foreach (var m in toRemove) mc.RemoveChild(m);
    }

    private static void AdjustMergesForRangeShiftUp(Worksheet ws, int c1, int c2, int r1, int r2, int count)
    {
        var mc = ws.GetFirstChild<MergeCells>();
        if (mc is null) return;
        var toRemove = new List<MergeCell>();

        foreach (var merge in mc.Elements<MergeCell>())
        {
            var (mc1, mr1, mc2, mr2) = ParseRange(merge.Reference);
            bool colInside = mc1 >= c1 && mc2 <= c2;          // 列完全在区域内
            bool rowInside = mr1 >= r1 && mr2 <= r2;           // 行完全在区域内
            bool colOutside = mc2 < c1 || mc1 > c2;           // 列完全在区域外
            bool rowBelow = mr1 > r2;                          // 行完全在区域下方

            if (colInside && rowInside) { toRemove.Add(merge); continue; }      // 整块被删除
            if (colOutside) continue;                                           // 列不涉及：不变
            if (colInside && rowBelow) { merge.Reference = BuildRange(mc1, mr1 - count, mc2, mr2 - count); continue; }
            // 部分重叠 → 取消合并（安全）
            toRemove.Add(merge);
        }
        foreach (var m in toRemove) mc.RemoveChild(m);
    }

    private static void AdjustMergesForRangeShiftLeft(Worksheet ws, int c1, int c2, int r1, int r2, int count)
    {
        var mc = ws.GetFirstChild<MergeCells>();
        if (mc is null) return;
        var toRemove = new List<MergeCell>();

        foreach (var merge in mc.Elements<MergeCell>())
        {
            var (mc1, mr1, mc2, mr2) = ParseRange(merge.Reference);
            bool rowInside = mr1 >= r1 && mr2 <= r2;
            bool colInside = mc1 >= c1 && mc2 <= c2;
            bool rowOutside = mr2 < r1 || mr1 > r2;
            bool colRight = mc1 > c2;

            if (rowInside && colInside) { toRemove.Add(merge); continue; }
            if (rowOutside) continue;
            if (rowInside && colRight) { merge.Reference = BuildRange(mc1 - count, mr1, mc2 - count, mr2); continue; }
            toRemove.Add(merge);
        }
        foreach (var m in toRemove) mc.RemoveChild(m);
    }

    // ---------------- 单元格/行 辅助 ----------------

    /// <summary>确保所有单元格都有显式 CellReference（r），以便后续按坐标操作。</summary>
    private static void MaterializeReferences(SheetData sheetData)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            int nextCol = 1;
            foreach (var cell in row.Elements<Cell>())
            {
                if (!string.IsNullOrEmpty(cell.CellReference))
                {
                    var (col, _) = ParseCellRef(cell.CellReference);
                    if (col > 0) { nextCol = col + 1; continue; }
                }
                cell.CellReference = CellReference.ToColumnLetters(nextCol) + (RowIndexOf(row) ?? 1);
                nextCol++;
            }
        }
    }

    private static void ShiftRowDown(Row row, int delta)
    {
        if (row.RowIndex?.HasValue == true)
            row.RowIndex = (uint)(row.RowIndex.Value + delta);
        else if (delta != 0)
            row.RowIndex = (uint)((RowIndexOf(row) ?? 0) + delta);
        foreach (var cell in row.Elements<Cell>())
        {
            var (col, r) = ParseCellRef(cell.CellReference);
            if (col == 0) continue;
            cell.CellReference = CellReference.ToColumnLetters(col) + (r + delta);
        }
    }

    private static void ShiftCellColumn(Cell cell, int delta)
    {
        var (col, r) = ParseCellRef(cell.CellReference);
        if (col == 0) return;
        cell.CellReference = CellReference.ToColumnLetters(col + delta) + r;
    }

    private static Cell? GetCell(SheetData sheetData, int rowIndex, int columnIndex)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            if (RowIndexOf(row) != rowIndex) continue;
            foreach (var cell in row.Elements<Cell>())
            {
                var (col, _) = ParseCellRef(cell.CellReference);
                if (col == columnIndex) return cell;
            }
        }
        return null;
    }

    private static Row GetOrCreateRow(SheetData sheetData, int rowIndex)
    {
        foreach (var row in sheetData.Elements<Row>())
            if (RowIndexOf(row) == rowIndex) return row;

        var newRow = new Row { RowIndex = (uint)rowIndex };
        // 保持行升序
        Row? inserted = null;
        foreach (var row in sheetData.Elements<Row>())
        {
            if (RowIndexOf(row) > rowIndex)
            {
                sheetData.InsertBefore(newRow, row);
                inserted = newRow;
                break;
            }
        }
        if (inserted is null) sheetData.AppendChild(newRow);
        return newRow;
    }

    /// <summary>对工作表行按行号升序、行内单元格按列升序重排（DeleteRange 跨行迁移后调用）。</summary>
    private static void ReorderSheetData(SheetData sheetData)
    {
        var rows = sheetData.Elements<Row>().OrderBy(r => RowIndexOf(r) ?? 0).ToList();
        foreach (var row in rows)
        {
            var cells = row.Elements<Cell>().ToList();
            cells.Sort((a, b) =>
            {
                var (ca, _) = ParseCellRef(a.CellReference);
                var (cb, _) = ParseCellRef(b.CellReference);
                return ca.CompareTo(cb);
            });
            foreach (var c in cells) row.RemoveChild(c);
            foreach (var c in cells) row.AppendChild(c);
        }
        foreach (var row in rows) sheetData.RemoveChild(row);
        foreach (var row in rows) sheetData.AppendChild(row);
    }

    private static IEnumerable<(int col, Cell cell)> CellsInRow(SheetData sheetData, int rowIndex)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            if (RowIndexOf(row) != rowIndex) continue;
            foreach (var cell in row.Elements<Cell>())
            {
                var (col, _) = ParseCellRef(cell.CellReference);
                if (col > 0) yield return (col, cell);
            }
        }
    }

    private static IEnumerable<(int row, Cell cell)> CellsInColumn(SheetData sheetData, int columnIndex)
    {
        foreach (var row in sheetData.Elements<Row>())
        foreach (var cell in row.Elements<Cell>())
        {
            var (col, _) = ParseCellRef(cell.CellReference);
            if (col == columnIndex) yield return (RowIndexOf(row) ?? 0, cell);
        }
    }

    // ---------------- 引用解析 ----------------

    private static (int col, int row) ParseCellRef(StringValue? reference)
    {
        if (string.IsNullOrEmpty(reference))
            return (0, 0);
        if (CellReference.TryParse(reference!, out int col, out int row, out _, out _))
            return (col, row);
        return (0, 0);
    }

    /// <summary>从 Row.RowIndex (UInt32Value) 安全取出行号。</summary>
    private static int? RowIndexOf(Row row)
    {
        var idx = row.RowIndex;
        return idx is not null && idx.HasValue ? (int?)idx.Value : null;
    }

    private static (int c1, int r1, int c2, int r2) ParseRange(string? range)
    {
        if (string.IsNullOrEmpty(range)) return (0, 0, 0, 0);
        var parts = range.Split(':');
        if (parts.Length != 2) return (0, 0, 0, 0);
        CellReference.TryParse(parts[0], out int c1, out int r1, out _, out _);
        CellReference.TryParse(parts[1], out int c2, out int r2, out _, out _);
        if (c1 > c2) (c1, c2) = (c2, c1);
        if (r1 > r2) (r1, r2) = (r2, r1);
        return (c1, r1, c2, r2);
    }

    private static string BuildRange(int c1, int r1, int c2, int r2)
        => $"{CellReference.ToColumnLetters(c1)}{r1}:{CellReference.ToColumnLetters(c2)}{r2}";

    // ---------------- 工作簿辅助 ----------------

    private static (WorksheetPart? part, Sheet? sheet) GetWorksheet(WorkbookPart wbPart, string sheetName)
    {
        var sheet = wbPart.Workbook.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => s.Name?.Value == sheetName);
        if (sheet is null) return (null, null);
        var wsPart = (WorksheetPart?)wbPart.GetPartById(sheet.Id!.Value);
        return (wsPart, sheet);
    }

    private static void RecalculateDimension(Worksheet ws)
    {
        var sheetData = ws.GetFirstChild<SheetData>();
        if (sheetData is null) return;
        int maxRow = 0, maxCol = 0;
        foreach (var row in sheetData.Elements<Row>())
        {
            int r = RowIndexOf(row) ?? 0;
            if (r > maxRow) maxRow = r;
            foreach (var cell in row.Elements<Cell>())
            {
                var (col, _) = ParseCellRef(cell.CellReference);
                if (col > maxCol) maxCol = col;
            }
        }
        var dim = ws.GetFirstChild<SheetDimension>();
        if (maxRow == 0 || maxCol == 0)
        {
            if (dim is not null) dim.Remove();
            return;
        }
        var newRef = $"A1:{CellReference.ToColumnLetters(maxCol)}{maxRow}";
        if (dim is null)
            ws.InsertAt(new SheetDimension { Reference = newRef }, 0);
        else
            dim.Reference = newRef;
    }

    private static void MarkFullRecalc(WorkbookPart wbPart)
    {
        wbPart.Workbook.CalculationProperties ??= new CalculationProperties();
        wbPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
    }
}
