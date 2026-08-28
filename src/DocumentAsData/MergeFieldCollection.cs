using System.Collections;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを取得するコレクションを表します。
/// </summary>
public class MergeFieldCollection : IEnumerable<MergeField>
{
    /// <summary>
    /// 指定した名前のMERGEFIELDを取得します。
    /// </summary>
    /// <param name="name">取得するMERGEFIELDの名前。</param>
    /// <returns>指定した名前のMERGEFIELD。</returns>
    public MergeField this[string name] =>
        throw new NotImplementedException();

    /// <summary>
    /// MERGEFIELDを列挙する列挙子を返します。
    /// </summary>
    /// <returns>MERGEFIELDを列挙する列挙子。</returns>
    public IEnumerator<MergeField> GetEnumerator() =>
        throw new NotImplementedException();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();
}
