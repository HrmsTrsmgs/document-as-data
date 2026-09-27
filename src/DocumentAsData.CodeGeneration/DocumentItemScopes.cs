namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// 通常項目のプロパティ生成と名前診断で、同じ所属範囲を選びます。
/// 文書全体を調べる必須項目検査や汎用Read/Replaceには適用しません。
/// </summary>
static class DocumentItemScopes
{
    /// <summary>
    /// 文書直下の生成プロパティに含めるMERGEFIELDを選びます。
    /// 明細内の同名項目を、通常項目として扱わないための選別です。
    /// </summary>
    /// <param name="items">文書全体のMERGEFIELD。</param>
    /// <returns>明細に属さないMERGEFIELD。</returns>
    internal static IEnumerable<MergeField> TopLevel(IEnumerable<MergeField> items) =>
        from item in items
        where !item.IsInRepeatingSection
        select item;

    /// <summary>
    /// 文書直下の生成プロパティに含めるContent Controlを選びます。
    /// 種類ごとに同じ所属条件を繰り返し記述しないための選別です。
    /// </summary>
    /// <typeparam name="T">生成対象のContent Controlの種類。</typeparam>
    /// <param name="items">文書全体のContent Control。</param>
    /// <returns>明細に属さないContent Control。</returns>
    internal static IEnumerable<T> TopLevel<T>(IEnumerable<T> items) where T : ContentControl =>
        from item in items
        where !item.IsInRepeatingSection
        select item;
}
