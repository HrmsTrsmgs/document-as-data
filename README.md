# document-as-data
Word文書を構造化データとして扱うための.NETライブラリ

現在は v0.1 の設計・開発初期段階です。公開API案と対象範囲は
[v0.1 設計案](docs/design-v0.1.md)を参照してください。

## 日付選択Content Control

Tagを名前として、Wordの日付選択Content Controlに記録された日時を
`DateTimeOffset`で読み書きできます。

```csharp
using var document = Document.Open("template.docx");

var deliveryDate = document.DatePickers["DeliveryDate"].Value;
document.DatePickers["DeliveryDate"].Value = new DateTimeOffset(
    2027,
    1,
    2,
    0,
    0,
    0,
    TimeSpan.FromHours(9));

document.SaveAs("output.docx");
```

`Read<T>()`と`Replace<T>()`では、`DateTimeOffset`プロパティを同じ名前の
日付選択Content Controlに対応付けます。表示文字列はテンプレートの表示形式と
表示言語に従って更新します。

## 現在対応していないもの

* `w:dataBinding` によりCustom XMLへデータバインドされたContent Control

データバインドされたContent Controlでは、表示内容とは別にCustom XML側にも値が保持されます。
DocumentAsDataは現在Custom XMLとの同期を行わないため、読み書きの対象には使用しないでください。
