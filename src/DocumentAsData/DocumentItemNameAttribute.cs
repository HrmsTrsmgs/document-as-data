namespace Marimo.DocumentAsData;

/// <summary>
/// プロパティに対応する文書項目の名前を指定します。
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public sealed class DocumentItemNameAttribute : Attribute
{
    /// <summary>
    /// 指定した文書項目の名前で属性を初期化します。
    /// </summary>
    /// <param name="name">文書項目の名前。</param>
    public DocumentItemNameAttribute(string name)
    {
        Name = name;
    }

    /// <summary>
    /// 文書項目の名前を取得します。
    /// </summary>
    public string Name { get; }
}
