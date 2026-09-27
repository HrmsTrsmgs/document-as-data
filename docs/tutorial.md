# チュートリアル：型付きコードから名前指定のAPIまで

[README](../README.md) / [機能リファレンス](reference.md)

まずREADMEの手順で、生成した `TemplateDocument` を使ってDOCXを保存してください。
ここでは同じサンプルを起点に、必要に応じて低レイヤーのAPIへ進みます。
すべての使い方を覚えてから利用を始める必要はありません。

以下のStream出力は `0.4.0` の機能です。

## 1. 自分のテンプレートを使う

READMEの `template.docx` には、`CustomerName` と `Address` という名前のMERGEFIELDがあります。
名前がC#のプロパティになり、文書の値を直接読み書きできます。

```csharp
Console.WriteLine(document.CustomerName);
document.CustomerName = "株式会社サンプル";
```

自分のDOCXでも、対象を `DocumentAsData` 項目としてプロジェクトへ追加し、
通常のビルドを行えば、その文書に合わせた型が生成されます。
たとえば `application.docx` なら `ApplicationDocument` と `ApplicationData` です。

文書の項目名を変更・追加した後も、もう一度ビルドします。
名前の衝突でビルドできない場合は、文書の名前を直すか、
[名前変換辞書](reference.md#名前変換辞書)で生成する名前を指定します。
生成された `.DocumentAsData.g.cs` は直接編集しません。

### 人に入力してもらう文書ではContent Controlを使う

人がWord上で入力する位置を示したい場合は、文字列のContent Controlが使えます。
Windows版Wordでは、開発タブからプレーンテキストまたはリッチテキストのContent Controlを挿入し、
そのコントロールのプロパティを設定します。
操作の詳細は[Microsoftのフォーム作成ガイド](https://support.microsoft.com/ja-jp/word/create-a-form-in-word-that-users-can-complete-or-print)を参照してください。

DocumentAsDataで使うIDは、タイトルや表示文字列ではなく **Tag** です。
たとえば、名前の入力欄のTagを `CustomerName`、住所欄を `Address` にします。
プレースホルダーによる入力案内はWord側で設定できます。
単に本文へ「CustomerName」と書くだけでは、名前付き項目にはなりません。

試す場合は、READMEで使ったDOCXの代わりに次の既存ファイルを `template.docx` としてコピーできます。

* `tests/DocumentAsData.Tests/TestData/複数のContent Control.docx`
* Tag：`CustomerName`（山田太郎）、`Address`（東京都）

再ビルドすると、こちらも `document.CustomerName` と `document.Address` で扱えます。
MERGEFIELDと文字列Content Controlの違いを、値を読み書きするコードで毎回指定する必要はありません。
チェックボックスと日付選択を追加した場合は、それぞれ `bool` と `DateTimeOffset` のプロパティになります。

### 同じ型で開ける文書

生成型の `Open` は、生成元にあった項目が開く文書にもあるか検査します。
不足している場合は `DocumentMappingException` になります。
たとえば `CustomerName` を必要とする型で、名前の入力欄がないDOCXを開くことはできません。

文字列項目は同名のMERGEFIELDと文字列Content Controlの間で互換性があります。
両方がある場合の扱いなど、詳細は[生成型で開く文書の条件](reference.md#生成型で開く文書の条件)を参照してください。

## 2. 生成したデータ型でまとめて受け渡す

画面や別の処理へ値を渡すときは、`Read()` で通常のオブジェクトとして取り出せます。
READMEのサンプルの `Program.cs` を、次の内容に置き換えて試せます。

```csharp
using DocumentAsDataDemo;

using var document = TemplateDocument.Open(
    Path.Combine(AppContext.BaseDirectory, "template.docx"));

// 戻り値は、テンプレートから生成されたTemplateDataです。
var data = document.Read();
Console.WriteLine(data.CustomerName);

// ここではデータだけを変更しています。
data.CustomerName = "株式会社サンプル";
data.Address = "大阪府";

// 変更したデータを、開いている文書へ反映します。
document.Replace(data);
document.SaveAs(Path.Combine(AppContext.BaseDirectory, "output.docx"));
```

`document.CustomerName` は開いている文書の値、
`data.CustomerName` は読み取った時点の値を保持する別のオブジェクトです。
データを変更しただけでは文書は変わらず、`Replace(data)` が必要です。
さらに、ファイルへ保存するには `SaveAs` または `Save` を呼びます。

新しくデータを組み立てて、`document.Replace(new TemplateData { ... })` と渡すこともできます。
生成されたData型は、Open XML SDKの型を保持するものではありません。

## 3. 自分で定義したクラスを使う

すでにアプリ側にデータ型がある場合は、コード生成を使わずに読み書きできます。
この使い方だけなら `Marimo.DocumentAsData.Core` パッケージの参照で足ります。
READMEのサンプルではBuildパッケージの依存関係に含まれるため、そのまま試せます。

次は、同じ `CustomerName` と `Address` を持つDOCXを読む、完結した `Program.cs` の例です。

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open(
    Path.Combine(AppContext.BaseDirectory, "template.docx"));

var data = document.Read<CustomerData>();
Console.WriteLine(data.Name);

data.Name = "株式会社サンプル";
data.Address = "大阪府";
document.Replace(data);

document.SaveAs(Path.Combine(AppContext.BaseDirectory, "output.docx"));

public sealed class CustomerData
{
    [DocumentItemName("CustomerName")]
    public string Name { get; set; } = "";

    public string Address { get; set; } = "";
}
```

`Address` はプロパティ名で文書の項目に対応します。
`Name` は文書側の名前と異なるため、`DocumentItemName` 属性で `CustomerName` を指定しています。
属性がない場合も、文書の `customerName` を `CustomerName` に対応させるなどの名前変換を行います。

対応する値の型は `string`、`bool`、`DateTimeOffset` です。
数値を自動的に文字列化したり、日付を日本標準時へ自動変換したりするAPIではありません。
また、同名のMERGEFIELDと文字列Content Controlが両方あると、
`Read<T>()`・`Replace(data)` は文字列の対応先を決められず例外になります。
詳しくは[オブジェクトとの対応付け](reference.md#オブジェクトとの対応付け)を参照してください。

## 4. 名前を指定して一項目ずつ扱う

扱う項目が実行時に決まる場合や、テンプレートにどんな項目があるか調べたい場合は、
`Document` のコレクションを直接使います。
READMEのMERGEFIELDのサンプルを使った `Program.cs` は、次のようになります。
第1節でContent Control版へ差し替えた場合は、`template.docx` をREADMEの
`単純形式のMERGEFIELD.docx` に戻して、再ビルドしてから進んでください。

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open(
    Path.Combine(AppContext.BaseDirectory, "template.docx"));

foreach (var field in document.MergeFields)
{
    Console.WriteLine($"{field.Name}: {field.Text}");
}

document.MergeFields["CustomerName"].Text = "株式会社サンプル";
document.MergeFields["Address"].Text = "大阪府";
document.SaveAs(Path.Combine(AppContext.BaseDirectory, "output.docx"));
```

このレイヤーでは、項目の種類ごとにコレクションと値のプロパティが分かれます。
次の例はそれぞれのTagを持つContent Controlがある文書で使用します。
READMEのMERGEFIELDだけの文書には、これらのコントロールはありません。

```csharp
document.ContentControls["CustomerName"].Text = "山田太郎";
document.CheckBoxes["Agreement"].IsChecked = true;
Console.WriteLine(document.DatePickers["ApplicationDate"].SelectedDateTime);
```

`ContentControls` は文字列用のContent Controlだけを扱います。
同じコレクション内で、指定した名前・Tagに一致する項目が0件または複数件なら例外になります。
「同名を全部更新する」という動作はしません。

## 5. 開き方・保存先を選ぶ

ここまでの例はファイルパスで開き、`SaveAs` で別のファイルへ保存していました。

| 操作 | 意味 |
| --- | --- |
| プロパティへの代入、`Replace(data)` | 開いている文書を変更する。保存元にはまだ書き込まない |
| `SaveAs("output.docx")` | 指定したファイルへ保存する。保存元は切り替わらない |
| `Save()` | ファイルパスで開いた保存元を上書きする |
| `Dispose()`、`using` の終了 | リソースを解放する。暗黙には保存しない |

### Streamから開く

ファイルパス以外で取得したDOCXも、通常の `Stream` から開けます。
たとえば `Program.cs` を次のようにできます。

```csharp
using Marimo.DocumentAsData;

using var stream = File.OpenRead(
    Path.Combine(AppContext.BaseDirectory, "template.docx"));
using var document = Document.Open(stream);

document.MergeFields["CustomerName"].Text = "株式会社サンプル";
document.SaveAs(Path.Combine(AppContext.BaseDirectory, "output.docx"));
```

生成型も `TemplateDocument.Open(stream)` で開けます。
読み取り専用・固定容量のMemoryStream・シークできない入力Streamにも対応します。
入力Streamの内容は変更せず、Documentの破棄時にも入力Streamを閉じません。
所有者である呼び出し側が破棄してください。読み取りによって位置は進む場合があります。

Streamで開いた文書に `Save()` を呼ぶと `NotSupportedException` になります。
容量を拡張できるStreamであっても、入力先へ書き戻すAPIとしては使えません。
別のStreamへ出力する場合は `SaveAs(Stream)` を使います。生成型も同じAPIを利用できます。

```csharp
using var output = new MemoryStream();
document.SaveAs(output);
byte[] bytes = output.ToArray();
```

出力先には同期書き込み・シーク・長さ変更が必要です。現在位置にかかわらず先頭から全体を置き換え、古い末尾も削除します。
出力先は閉じません。Streamとして読み直す場合は `output.Position = 0` で先頭へ戻してください。
保存中に出力先への書き込みが失敗すると、出力先は途中まで変更される場合があります。元文書とは別のStreamを渡してください。

Blazorでユーザーが選択したファイルは、まず呼び出し側で非同期にメモリへ読み込みます。
`IBrowserFile.OpenReadStream()` のStreamは同期読み取りができないため、そのまま同期APIの `Document.Open` には渡しません。
次は `browserFile`、許容サイズの `maxAllowedSize`、書き込む `data` が利用側にある場合の例です。

```csharp
using var input = new MemoryStream();
await using var upload = browserFile.OpenReadStream(maxAllowedSize);
await upload.CopyToAsync(input);
input.Position = 0;

using var document = Document.Open(input);
document.Replace(data);
using var output = new MemoryStream();
document.SaveAs(output);
byte[] bytes = output.ToArray(); // 利用側のダウンロード処理へ渡します。
```

ファイルパスや一時ファイルは不要です。ただし文書全体と出力をメモリに持つので、ファイルサイズの上限を設けてください。
ライブラリはBlazor固有の型には依存しません。ブラウザーでのダウンロード処理は利用側の責務です。
Blazor WebAssemblyでの発行後の実動作は未検証です。

### 構造を検証して開く

OOXMLの構造検証が必要なら、低レイヤーの `Document.Open` で指定できます。

```csharp
using var document = Document.Open("template.docx", validate: true);
```

Stream版にも同じオプションがあります。違反が見つかると `InvalidDataException` になります。
これはWord上の見た目や入力内容の正しさを保証する検査ではありません。
生成型の必須項目検査とも別の機能です。

## 6. 繰り返し明細を扱う

明細を扱うには、Word側で繰り返しセクションを設定します。明細が1件でも同じコレクション用のAPIを使います。
同じTagを持つ入力欄を並べただけでは繰り返しセクションにはなりません。

まず既存のサンプルで試します。READMEで作ったプロジェクトの `template.docx` を、
`tests/DocumentAsData.Tests/TestData/繰り返しセクションに2件の明細.docx` のコピーに差し替えてください。
リポジトリを取得していない場合は、プロジェクトのフォルダーで次を実行してサンプルを取得できます。既存の `template.docx` は置き換わります。

```powershell
Invoke-WebRequest "https://raw.githubusercontent.com/HrmsTrsmgs/document-as-data/main/tests/DocumentAsData.Tests/TestData/繰り返しセクションに2件の明細.docx" -OutFile template.docx
```

この文書にはTagが `Items` のセクションと、`ItemName` が商品A・商品Bの明細があります。
コピー先の名前を `template.docx` にするので、生成型名は `TemplateDocument` のままです。

`Program.cs` を次の内容に置き換え、通常のビルドと実行を行います。

```csharp
using DocumentAsDataDemo;

using var document = TemplateDocument.Open(
    Path.Combine(AppContext.BaseDirectory, "template.docx"));

foreach (var item in document.Items)
{
    Console.WriteLine(item.ItemName);
}

// 明細はデータとして読み取ります。変更しただけでは文書は変わりません。
var items = document.Items.ToArray();
items[0].ItemName = "商品C";
document.Items.Replace(items);

// 1件のデータを渡すと、文書の明細も1件になります。
document.Items.Replace([new() { ItemName = "商品D" }]);
document.SaveAs(Path.Combine(AppContext.BaseDirectory, "output.docx"));
```

コンソールには商品A・商品Bが表示され、保存した文書には商品Dの明細が1件残ります。
セクションのTagから `Items` プロパティと、明細1件分の `TemplateDocument.ItemsData` 型が生成されます。
上の `new()` はその明細データ型です。文書全体の `Read()`／`Replace(data)` へ明細リストを渡す方式ではありません。

コード生成を使わない場合は、同じテンプレートを次のように扱えます。

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open(
    Path.Combine(AppContext.BaseDirectory, "template.docx"));
var section = document.RepeatingSections["Items"];

section.Items[0].ContentControls["ItemName"].Text = "商品C";
var items = section.Read<ItemData>().ToArray();
items[1].ItemName = "商品D";
section.Replace(items);
document.SaveAs(Path.Combine(AppContext.BaseDirectory, "output.docx"));

public sealed class ItemData
{
    public string ItemName { get; set; } = "";
}
```

文書全体の `document.ContentControls["ItemName"]` では同名が複数になり、一意に取得できません。
`section.Items[0].ContentControls["ItemName"]` と指定すると、最初の明細内だけを検索できます。

この機能は1件以上の明細の基本対応です。空のデータでの置き換えや入れ子などには対応していません。
失敗時の部分変更や生成型の検査範囲についても、[繰り返しセクションの制約](reference.md#繰り返しセクションの制約)を確認してください。

## 次に調べる

各APIの詳細、日付の扱い、Content Controlの制限、生成設定は[機能リファレンス](reference.md)へ進んでください。
PDFが必要な場合は、保存したDOCXをWordやLibreOfficeなどでPDF化する処理を利用側で用意します。
