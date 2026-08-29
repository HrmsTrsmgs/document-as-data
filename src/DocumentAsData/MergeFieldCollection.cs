using System.Collections;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを取得するコレクションを表します。
/// </summary>
public class MergeFieldCollection : IEnumerable<MergeField>
{
    readonly Document document;
    readonly Dictionary<Wordprocessing.SimpleField, MergeField> cache = [];

    internal MergeFieldCollection(Document document)
    {
        this.document = document;
    }

    /// <summary>
    /// 指定した名前のMERGEFIELDを取得します。
    /// </summary>
    /// <param name="name">取得するMERGEFIELDの名前。</param>
    /// <returns>指定した名前のMERGEFIELD。</returns>
    public MergeField this[string name] =>
        (
            from field in this
            where field.Name == name
            select field
        ).First();

    /// <summary>
    /// MERGEFIELDを列挙する列挙子を返します。
    /// </summary>
    /// <returns>MERGEFIELDを列挙する列挙子。</returns>
    public IEnumerator<MergeField> GetEnumerator() =>
        (
            from field in document.SimpleFields
            let name = MergeFieldName(field)
            where name is not null
            select GetMergeField(field, name)
        ).GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();

    MergeField GetMergeField(Wordprocessing.SimpleField field, string name)
    {
        if (cache.TryGetValue(field, out var mergeField))
        {
            return mergeField;
        }

        mergeField = new(name);
        cache.Add(field, mergeField);
        return mergeField;
    }

    static string? MergeFieldName(Wordprocessing.SimpleField field)
    {
        const string fieldType = "MERGEFIELD";

        var instruction = field.Instruction?.Value?.Trim();
        if (instruction is null ||
            !instruction.StartsWith(fieldType, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return instruction[fieldType.Length..].Trim().Trim('"');
    }
}
