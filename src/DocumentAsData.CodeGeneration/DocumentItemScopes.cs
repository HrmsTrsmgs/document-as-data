namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// プロパティ生成、名前診断、必須項目検査で、通常項目と明細の生成対象を宣言順に選びます。
/// </summary>
static class DocumentItemScopes
{
    /// <summary>
    /// 生成するプロパティと必須項目検査を選ぶための、文書上の項目の種類です。
    /// </summary>
    internal enum ItemKind
    {
        MergeField,
        TextContentControl,
        CheckBox,
        DatePicker
    }

    /// <summary>
    /// 検査を種類ごとに生成できるよう、同じ種類の元名を文書内の順序で保持します。
    /// </summary>
    /// <param name="Kind">文書上の項目の種類。</param>
    /// <param name="Names">プロパティ生成と検査に使う、変換前の項目名。</param>
    internal readonly record struct ItemGroup(ItemKind Kind, IEnumerable<string> Names);

    /// <summary>
    /// 文書直下の通常項目を、文書型の宣言順に取得します。
    /// </summary>
    internal static ItemGroup[] TopLevelItems(Document document) =>
        GroupItems(
            TopLevel(document.MergeFields),
            TopLevel(document.ContentControls),
            TopLevel(document.CheckBoxes),
            TopLevel(document.DatePickers),
            mergeFieldsFirst: true);

    /// <summary>
    /// 生成元の先頭明細から、明細型の宣言順に項目を取得します。
    /// </summary>
    internal static ItemGroup[] DetailItems(RepeatingSection section)
    {
        var templateItem = section.Items[0];
        return GroupItems(
            templateItem.MergeFields,
            templateItem.ContentControls,
            templateItem.CheckBoxes,
            templateItem.DatePickers,
            mergeFieldsFirst: false);
    }

    /// <summary>
    /// 項目の種類と元名の取り出し方を対応付けます。
    /// 文書直下と明細では、最初の2種類の文字列項目だけが逆順になります。
    /// </summary>
    static ItemGroup[] GroupItems(
        IEnumerable<MergeField> mergeFields,
        IEnumerable<TextContentControl> contentControls,
        IEnumerable<CheckBox> checkBoxes,
        IEnumerable<DatePicker> datePickers,
        bool mergeFieldsFirst)
    {
        ItemGroup fields = new(ItemKind.MergeField, from item in mergeFields select item.Name);
        ItemGroup textControls = new(ItemKind.TextContentControl, from item in contentControls select item.Tag);
        var (first, second) = mergeFieldsFirst ? (fields, textControls) : (textControls, fields);

        return
        [
            first,
            second,
            new(ItemKind.CheckBox, from item in checkBoxes select item.Tag),
            new(ItemKind.DatePicker, from item in datePickers select item.Tag)
        ];
    }

    /// <summary>
    /// 文書直下の生成プロパティに含めるMERGEFIELDを選びます。
    /// 明細内の同名項目を、通常項目として扱わないための選別です。
    /// </summary>
    /// <param name="items">文書全体のMERGEFIELD。</param>
    /// <returns>明細に属さないMERGEFIELD。</returns>
    static IEnumerable<MergeField> TopLevel(IEnumerable<MergeField> items) =>
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
    static IEnumerable<T> TopLevel<T>(IEnumerable<T> items) where T : ContentControl =>
        from item in items
        where !item.IsInRepeatingSection
        select item;
}
