namespace Marimo.DocumentAsData;

/// <summary>
/// 文書とオブジェクトの対応付けに失敗したことを表します。
/// </summary>
public sealed class DocumentMappingException : Exception
{
    /// <summary>
    /// 詳細メッセージを指定せず、対応付けの失敗を表します。
    /// </summary>
    public DocumentMappingException()
    {
    }

    /// <summary>
    /// 指定したメッセージで、対応付けに失敗した理由を示します。
    /// </summary>
    /// <param name="message">対応付けに失敗した理由。</param>
    public DocumentMappingException(string message)
        : base(message)
    {
    }
}
