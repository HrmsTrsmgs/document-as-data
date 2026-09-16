namespace Marimo.DocumentAsData;

/// <summary>
/// プロパティに対応する文書項目の名前を指定します。
/// </summary>
/// <param name="name">文書項目の名前。</param>
[AttributeUsage(AttributeTargets.Property)]
public sealed class DocumentItemNameAttribute(string name) : Attribute
{
    /// <summary>
    /// 文書項目の名前を取得します。
    /// </summary>
    public string Name { get; } = name;
}
