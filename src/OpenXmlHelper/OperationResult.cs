namespace OpenXmlHelper;

/// <summary>
/// 表示一次 Excel 区域删除操作的结果。
/// </summary>
public sealed class OperationResult
{
    /// <summary>操作是否成功。</summary>
    public bool Success { get; init; }

    /// <summary>操作失败时的错误信息（成功时为 null）。</summary>
    public string? ErrorMessage { get; init; }

    /// <summary>获取成功结果。</summary>
    public static OperationResult Ok() => new() { Success = true };

    /// <summary>获取失败结果并附带错误信息。</summary>
    public static OperationResult Fail(string message) => new() { Success = false, ErrorMessage = message };

    public override string ToString() => Success ? "Success" : $"Failed: {ErrorMessage}";
}
