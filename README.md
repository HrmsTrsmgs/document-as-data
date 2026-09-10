# DocumentAsData

DocumentAsDataは、Word文書（DOCX）を構造化データとして読み書きする.NETライブラリです。
Open XML SDKを内部で使用し、利用側は文書内の名前付き項目を通じてデータを扱います。

現在の対象はDOCXです。ライブラリの実行にWordのインストールは必要ありません。
文書のレイアウトを編集するのではなく、あらかじめ名前を付けた場所とプログラムの間でデータを交換します。

## 利用環境と参照

現在のプロジェクトは `.NET 10`（`net10.0`）を対象にしています。
このREADMEではNuGetへの公開済みバージョンを前提にせず、ソースからビルドする構成を説明します。
利用側のプロジェクトから `src/DocumentAsData/DocumentAsData.csproj` を参照してください。
名前空間は `Marimo.DocumentAsData` です。

| 文書内の項目 | コレクション | 識別方法 | 値のプロパティと型 |
| --- | --- | --- | --- |
| MERGEFIELD | `MergeFields` | フィールド名 | `Text`：`string` |
| 文字列Content Control | `ContentControls` | Tag | `Text`：`string` |
| チェックボックス | `CheckBoxes` | Tag | `IsChecked`：`bool` |
| 日付選択Content Control | `DatePickers` | Tag | `SelectedDateTime`：`DateTimeOffset` |

各コレクションは列挙できます。例えば、`foreach (var field in document.MergeFields)` で
`field.Name` と `field.Text` を取得できます。
`ContentControls` はすべての種類のContent Controlを返すコレクションではありません。

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

## チェックボックス

WordのチェックボックスContent ControlにTag `Agreement` を設定した文書を用意します。
チェック状態は表示文字の解析ではなく、`bool` として読み書きします。

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open("agreement.docx");

var agreed = document.CheckBoxes["Agreement"].IsChecked;
document.CheckBoxes["Agreement"].IsChecked = true;
document.SaveAs("agreement-output.docx");
```

`Read<T>()`・`Replace<T>()` では、DTOの `bool` プロパティを同名のTagを持つCheckBoxへ対応付けます。
`DocumentItemName` 属性で別のTagも指定できます。

## 日付選択Content Control

Tagを名前として、Wordの日付選択Content Controlに記録された日時を
`DateTimeOffset`で読み書きできます。

```csharp
using System;
using Marimo.DocumentAsData;

// TagがDeliveryDateの日付選択Content Controlを用意します。
using var document = Document.Open("delivery.docx");

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

`SelectedDateTime` はOOXMLに記録された日時を扱います。`DateOnly` へ自動変換したり、
日本標準時へ一律に変換したりはしません。誕生日・期日などの業務上の意味付けは利用側で行ってください。

## Streamから開く

`Open(Stream)` も使用できます。ファイルパス版との重要な違いは保存先です。

| 開き方 | 編集内容の反映先 | Dispose時の入力の扱い |
| --- | --- | --- |
| `Open(string)` | 作業コピー。`SaveAs` で別ファイルへ保存 | 元ファイルを変更せず解放 |
| `Open(Stream)` | 渡したStreamへ書き戻す | 文書を閉じるがStream自体は閉じない |

Streamは呼び出し側が所有し、使用後にDisposeします。
入力を保持したい場合は、編集用の書き込み・シーク可能な `MemoryStream` へコピーします。

```csharp
using System.IO;
using Marimo.DocumentAsData;

using var source = File.OpenRead("template.docx");
using var working = new MemoryStream();
source.CopyTo(working);
working.Position = 0;

using (var document = Document.Open(working))
{
    document.MergeFields["CustomerName"].Text = "株式会社○○";
}

// 文書をDisposeした後に、完成したDOCXのバイト列を取り出します。
var bytes = working.ToArray();
File.WriteAllBytes("output.docx", bytes);
```

固定長のバッファで編集領域が不足しないよう、この例では拡張可能な `MemoryStream` を作成しています。
任意のStreamがそのまま使えることを保証するものではありません。
ブラウザから取得した読み取り専用・非シーク可能なStreamも、呼び出し側で編集用Streamへコピーしてください。
本体は `IBrowserFile` に依存しません。ブラウザでのダウンロードやHTTPレスポンスの生成は利用側の責務です。
Blazor WebAssembly上での実動作は、このリポジトリのテストでは確認していません。

## 開くときにOOXMLを検証する

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open("template.docx", validate: true);
```

Stream版にも `validate` 引数があります。省略時は検証を実行しません。
`true` の場合はOpen XML SDKのValidatorで検証し、検証エラーがあれば `InvalidDataException` を投げます。
検証は文書の表示・レイアウトの正しさや、業務上のデータの妥当性まで保証するものではありません。

## 文書から型付きコードを生成する

`src/DocumentAsData.CodeGeneration/DocumentAsData.CodeGeneration.csproj` を参照すると、
文書内の項目に対応するC#ソースを生成できます。

```csharp
using System.IO;
using Marimo.DocumentAsData.CodeGeneration;

var sources = DocumentWrapperGenerator.GenerateSources(
    "template.docx",
    options => options.Namespace = "MyDocuments");

// 現在は一つの文書から一つのソースが返ります。
File.WriteAllText("Template.DocumentAsData.g.cs", sources[0]);
```

生成ソースをDocumentAsDataを参照する利用側プロジェクトへ追加してコンパイルすると、
`TemplateDocument` と `TemplateData` を使用できます。
次の例は、先ほどの `CustomerName` と `Address` を持つテンプレートから生成した場合です。

```csharp
using MyDocuments;

using var document = TemplateDocument.Open("template.docx");

// 文書を直接読み書きするプロパティです。
document.CustomerName = "株式会社○○";

// 型引数なしのReadで、生成されたTemplateDataを取得します。
var data = document.Read();
data.Address = "東京都…";
document.Replace(data);
document.SaveAs("output.docx");
```

生成Documentのプロパティは、文字列・bool・DateTimeOffsetの値を直接公開します。
MERGEFIELDやContent Controlのオブジェクトを取得するプロパティではありません。
生成Data型は通常のデータ保持用オブジェクトで、変更後に `Replace` が必要です。
生成Documentは `Document` を継承し、ファイルパスとStreamの `Open` を公開します。
生成型はいずれも `partial` なので、追記は別ファイルで行い、生成ファイルは直接編集しないでください。

### 生成名を指定する

```csharp
var sources = DocumentWrapperGenerator.GenerateSources(
    "template.docx",
    options =>
    {
        options.Namespace = "MyDocuments";
        options.NameMappings["CustomerName"] = "ClientName";
    });
```

既定の名前空間は `Generated` です。
文書ファイル名や項目名はC#識別子へ変換され、型名には `Document`・`Data` が付きます。
`NameMappings` は自動変換より優先します。辞書には生成後の名前ではなく、文書内の元名を指定します。
同じ元名への設定は、項目の種類を区別せず適用されます。
生成Dataプロパティ名が元名と異なる場合は `DocumentItemName` 属性が生成され、読み書き先を保持します。

### 名前の衝突を診断する

`DocumentWrapperGenerator.GenerateDiagnostics(filePath, configure)` は、生成名の重複や
空の生成文書名を診断します。結果には `IsError`、`GeneratedName`、`SourceNames`、`InvalidSourceName` が含まれます。
名前の変換で異なる項目が同じC#名になる場合は、テンプレートまたは `NameMappings` を調整してください。
診断と生成には同じ設定を渡します。

`GenerateSources` は診断を自動実行しません。また診断は、すべてのC#コンパイルエラーを検出する機能ではありません。
生成後は利用側のプロジェクトでビルドしてください。

## MSBuild連携の現在の範囲

`DocumentAsData.Build` には `Marimo.DocumentAsData.Build.GenerateDocumentAsData` タスクがあります。
入力は `DocumentFiles`（`ITaskItem[]`）と `ProjectDirectory`、出力は `GeneratedFiles` です。
文書の隣に `<文書ファイル名の拡張子を除いた部分>.DocumentAsData.g.cs` を作成し、再実行時には上書きします。
出力項目には `DependentUpon` と `DesignTimeSharedInput` メタデータを設定します。

名前変換辞書は、例えば `template.docx` に対して `template.documentasdata.json` を置きます。

```json
{
  "CustomerName": "ClientName"
}
```

文書と同じディレクトリを優先し、見つからなければ `ProjectDirectory` 直下を探します。
両方ある場合に辞書をマージする動作ではありません。タスクが生成する名前空間は現在 `Generated` です。

`src/DocumentAsData.Build/buildTransitive/Marimo.DocumentAsData.Build.targets` を明示的にimportし、
`DocumentAsDataTaskAssembly` にビルド済みタスクDLLのパスを指定すると、
`DocumentAsData` 項目から `GenerateDocumentAsDataSources` ターゲットで生成できます。
PowerShellから `dotnet msbuild` でこのターゲットを呼び出す動作をテストしています。

`DesignTimeBuild=true` で評価すると、対象DOCXに対応する生成済みソースを `Compile` 項目へ登録します。
これは登録の処理であり、生成ファイルは事前に作成しておく必要があります。
NuGet経由の自動import、ビルドアクションの登録、通常ビルドへの自動組み込み、
SDKの既定Compile項目との重複対策はまだ整備していません。
このテストはVisual Studio全体の操作を保証するものではありません。

## 現在対応していないもの

* `w:dataBinding` によりCustom XMLへデータバインドされたContent Control

データバインドされたContent Controlでは、表示内容とは別にCustom XML側にも値が保持されます。
DocumentAsDataは現在Custom XMLとの同期を行わないため、読み書きの対象には使用しないでください。

## 対象範囲と注意点

* 対象となる項目の探索範囲は文書の本文です。ヘッダー・フッターの読み書きは対象にしていません。
* MERGEFIELDは単純フィールドと複合フィールドに対応します。Word標準の差し込み印刷やCSV連携は実行しません。
* MERGEFIELDの `Text` は保存された表示結果を読み書きします。フィールドの計算エンジンではありません。
* 文字列Content Controlはプレーンテキストとリッチテキストを扱いますが、書式編集APIではありません。
* プレースホルダー表示中の文字列Content Controlは、案内文ではなく空文字列として読み取ります。
* 一つの項目内で文字列が複数のrunへ分かれていても、値をまとめて扱います。部分ごとの書式を指定することはできません。
* コンボボックス、ドロップダウン、画像、繰り返しセクションなどを文字列Content Controlとして扱いません。繰り返し行の追加・削除APIもありません。
* チェックボックスはWordのContent Controlが対象です。自由入力のチェック記号を判定する機能ではありません。
* 保存時にContent Controlを除去したり、MERGEFIELDを通常テキストへ変換したりはしません。
* ODT、PDFへのレンダリング、ページ・行・レイアウトの計算、OCR、自然言語解析は対象外です。

PDFが必要な場合は、生成したDOCXをWordやLibreOfficeなどで変換してください。
変換処理は利用側が行います。PDF化のためだけにContent Controlを除去する機能は提供しません。

## 開発とテスト

`global.json` は .NET SDK `10.0.302` を指定しています（同系列の最新パッチを許容）。
MSBuild連携テストではPowerShell 7の `pwsh` をPATH上から起動します。
リポジトリのルートで実行します。

```powershell
dotnet restore DocumentAsData.slnx
dotnet build DocumentAsData.slnx --no-restore
dotnet test DocumentAsData.slnx --no-build --no-restore
```

主な構成は次のとおりです。

* `src/DocumentAsData`：DOCXの読み書きとオブジェクトの対応付け
* `src/DocumentAsData.CodeGeneration`：型付きコードの生成と名前の診断
* `src/DocumentAsData.Build`：MSBuildタスク
* `tests/DocumentAsData.Tests`：文書操作のテストと固定DOCXデータ
* `tests/DocumentAsData.CodeGeneration.Tests`：生成コードとMSBuildタスクのテスト

テストでは固定DOCXの読み書き、保存後の再読込、OOXML構造、生成C#のコンパイルなどを確認します。
自動テストの成功は、すべてのWord文書や表示・印刷結果を保証するものではありません。
READMEのサンプルはそれぞれに記載した項目を持つ文書を必要とし、全コードブロックを一つのプログラムへ連結する想定ではありません。

## ライセンス

[MIT License](LICENSE)
