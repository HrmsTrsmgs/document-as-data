using System.Collections;
using DocumentFormat.OpenXml;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData;

/// <summary>
/// 文書内のMERGEFIELDを取得するコレクションを表します。
/// </summary>
public class MergeFieldCollection : IEnumerable<MergeField>
{
    readonly Document document;
    readonly Dictionary<OpenXmlElement, MergeField> cache = [];

    internal MergeFieldCollection(Document document)
    {
        this.document = document;
    }

    /// <summary>
    /// 指定した名前のMERGEFIELDを取得します。
    /// </summary>
    /// <param name="name">取得するMERGEFIELDの名前。</param>
    /// <returns>指定した名前のMERGEFIELD。</returns>
    /// <exception cref="KeyNotFoundException">指定した名前のMERGEFIELDが存在しない場合。</exception>
    public MergeField this[string name] =>
        (
            from field in this
            where field.Name == name
            select field
        ).FirstOrDefault() ?? throw new KeyNotFoundException();

    /// <summary>
    /// MERGEFIELDを列挙する列挙子を返します。
    /// </summary>
    /// <returns>MERGEFIELDを列挙する列挙子。</returns>
    public IEnumerator<MergeField> GetEnumerator() =>
        MergeFields().GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() =>
        GetEnumerator();

    IEnumerable<MergeField> MergeFields()
    {
        Wordprocessing.FieldChar? complexField = null;
        Wordprocessing.FieldCode? instruction = null;
        List<Wordprocessing.Text>? valueTexts = null;

        foreach (var element in document.Elements)
        {
            if (element is Wordprocessing.SimpleField simpleField)
            {
                var simpleFieldName = MergeFieldName(simpleField.Instruction?.Value);
                if (simpleFieldName is not null)
                {
                    yield return GetMergeField(
                        simpleField,
                        () => new(document, simpleField, simpleFieldName));
                }

                continue;
            }

            if (element is Wordprocessing.FieldChar fieldChar)
            {
                if (fieldChar.FieldCharType?.Value == Wordprocessing.FieldCharValues.Begin)
                {
                    complexField = fieldChar;
                    instruction = null;
                    valueTexts = null;
                }
                else if (
                    complexField is not null &&
                    fieldChar.FieldCharType?.Value == Wordprocessing.FieldCharValues.Separate)
                {
                    valueTexts = [];
                }
                else if (
                    complexField is not null &&
                    fieldChar.FieldCharType?.Value == Wordprocessing.FieldCharValues.End)
                {
                    var complexFieldName = MergeFieldName(instruction?.Text);
                    if (complexFieldName is not null && valueTexts is not null)
                    {
                        yield return GetMergeField(
                            complexField,
                            () => new(document, complexFieldName, valueTexts));
                    }

                    complexField = null;
                    instruction = null;
                    valueTexts = null;
                }

                continue;
            }

            if (complexField is null)
            {
                continue;
            }

            if (valueTexts is null &&
                instruction is null &&
                element is Wordprocessing.FieldCode fieldCode)
            {
                instruction = fieldCode;
            }
            else if (valueTexts is not null && element is Wordprocessing.Text text)
            {
                valueTexts.Add(text);
            }
        }
    }

    MergeField GetMergeField(OpenXmlElement field, Func<MergeField> create)
    {
        if (cache.TryGetValue(field, out var mergeField))
        {
            return mergeField;
        }

        mergeField = create();
        cache.Add(field, mergeField);
        return mergeField;
    }

    static string? MergeFieldName(string? fieldInstruction)
    {
        const string fieldType = "MERGEFIELD";

        var instruction = fieldInstruction?.Trim();
        if (instruction is null ||
            !instruction.StartsWith(fieldType, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return instruction[fieldType.Length..].Trim().Trim('"');
    }
}
