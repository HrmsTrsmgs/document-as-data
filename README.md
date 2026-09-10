# DocumentAsData

DocumentAsDataは、Word文書（DOCX）を構造化データとして読み書きする.NETライブラリです。
Open XML SDKを内部で使用し、利用側は文書内の名前付き項目を通じてデータを扱います。

## 最小の読み書き例

まずWordで、`template.docx` の本文に次の項目を用意します。

* 名前が `CustomerName` の差し込みフィールド（MERGEFIELD）
* Tagが `Address` の文字列Content Control（コンテンツコントロール）

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open("template.docx");

// Content Controlに入力された文字列を読み取ります。
var address = document.ContentControls["Address"].Text;

// 名前やTagで指定した項目へ文字列を書き込みます。
document.MergeFields["CustomerName"].Text = "株式会社○○";
document.ContentControls["Address"].Text = "東京都…";

document.SaveAs("output.docx");
```

`MergeFields` はMERGEFIELDの名前、`ContentControls` はTagで検索します。
それぞれのコレクション内で、指定した名前やTagに一致する項目は一つだけにしてください。
該当項目がない場合や複数ある場合は例外になります。

ファイルパスを渡す `Open(string)` では、変更は元のファイルへ自動保存されません。
変更を残すには、`SaveAs` で別ファイルへ保存します。`using` の終了時にも元のファイルは変更しません。
MERGEFIELDやContent Controlの構造は残したまま保存します。

データとして扱いたい場所には、Word側で明示的に名前付きの構造を付けます。
自由入力された文章から氏名や住所を推測して取り出す機能ではありません。

## オブジェクトとして読み書きする

`Read<T>()` で文書の値をオブジェクトへ読み込み、`Replace<T>()` でオブジェクトの値を文書へ書き込めます。
前の例と同じ `template.docx` を、次のように扱えます。

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open("template.docx");

var data = document.Read<CustomerData>();
data.CustomerName = "株式会社○○";
data.Address = "東京都…";

document.Replace(data);
document.SaveAs("output.docx");

public sealed class CustomerData
{
    public string CustomerName { get; set; } = "";

    public string Address { get; set; } = "";
}
```

文字列プロパティは、プロパティ名と同じ名前のMERGEFIELD、または同じTagの文字列Content Controlへ対応付けます。
この例では `CustomerName` はMERGEFIELD、`Address` はContent Controlですが、DTO側で種類を指定する必要はありません。
両方に同名の項目がある場合は、対応先を一つに決められないため例外になります。

`Read<T>()` が返すのは文書と連動するオブジェクトではありません。
プロパティを変更した後は `Replace(data)` で文書へ反映し、`SaveAs` で別ファイルへ保存します。
`Replace(data)` の型引数は、渡したオブジェクトから推論されます。

### 文書の名前とプロパティ名を変える

プロパティ名とは別の名前で文書項目を指定するには、DTOのプロパティに `DocumentItemName` 属性を付けます。
前の例の `CustomerData` を次の定義に置き換えると、`Name` が文書内の `CustomerName` に対応します。
呼び出し側の代入も `data.CustomerName = ...` から `data.Name = ...` へ変更します。

```csharp
public sealed class CustomerData
{
    [DocumentItemName("CustomerName")]
    public string Name { get; set; } = "";

    public string Address { get; set; } = "";
}
```

`Read<CustomerData>()` は `CustomerName` の値を `Name` へ読み込み、
`Replace(data)` は `data.Name` の値を `CustomerName` へ書き込みます。
属性のない `Address` は、引き続きプロパティ名で対応付けます。
文書内の名前やTag自体を変更する機能ではありません。

## 日付選択Content Control

Tagを名前として、Wordの日付選択Content Controlに記録された日時を
`DateTimeOffset`で読み書きできます。

```csharp
using var document = Document.Open("template.docx");

var deliveryDate = document.DatePickers["DeliveryDate"].SelectedDateTime;
document.DatePickers["DeliveryDate"].SelectedDateTime = new DateTimeOffset(
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
