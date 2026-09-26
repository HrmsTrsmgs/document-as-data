using System.Collections;

namespace Marimo.DocumentAsData;

/// <summary>
/// 明細を指定した型で列挙し、置き換えられる繰り返しセクションを表します。
/// </summary>
/// <typeparam name="T">明細1件分のデータを保持する型。</typeparam>
/// <param name="source">読み書きする繰り返しセクション。</param>
public class RepeatingSection<T>(RepeatingSection source) : RepeatingSection(source), IEnumerable<T>
{
    /// <summary>
    /// 各明細を指定した型へ読み込み、文書順に列挙します。
    /// </summary>
    /// <returns>明細データの列挙子。</returns>
    public IEnumerator<T> GetEnumerator() =>
        Read<T>().GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();

    /// <summary>
    /// 明細の内容と件数を、指定した型付きデータで置き換えます。
    /// </summary>
    /// <param name="items">1件以上の、置き換え後の明細データ。</param>
    public void Replace(IEnumerable<T> items) =>
        base.Replace(items);
}
