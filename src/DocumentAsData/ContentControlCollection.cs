using System.Collections;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のContent Controlを取得するコレクションを表します。
/// </summary>
public class ContentControlCollection : IEnumerable<ContentControl>
{
    /// <summary>
    /// 指定したTagのContent Controlを取得します。
    /// </summary>
    /// <param name="tag">取得するContent ControlのTag。</param>
    /// <returns>指定したTagのContent Control。</returns>
    public ContentControl this[string tag] =>
        throw new NotImplementedException();

    /// <summary>
    /// Content Controlを列挙する列挙子を返します。
    /// </summary>
    /// <returns>Content Controlを列挙する列挙子。</returns>
    public IEnumerator<ContentControl> GetEnumerator() =>
        throw new NotImplementedException();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
