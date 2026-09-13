# 機能の概要とリファレンス

[README](../README.md) / [チュートリアル](tutorial.md)

DocumentAsDataは、DOCXの本文に名前を付けた項目とプログラムの間でデータを交換します。
このページは機能と制約を確認するための参照用です。最初の読み書きはチュートリアルから始めてください。

## 使うAPIを選ぶ

| やりたいこと | 使うもの | 特徴 |
| --- | --- | --- |
| テンプレートに合わせたプロパティで読み書きする | 生成された `TemplateDocument` などの型 | 名前の文字列を利用側コードに繰り返さず、補完を使える |
| フォームなどへデータをまとめて渡す | 生成型の `Read()`・`Replace(data)` | 生成されたData型で入出力する |
| 自分で定義したデータ型を使う | `Document.Read<T>()`・`Replace(data)` | プロパティ名や属性で文書項目と対応付ける |
| 実行時に決まる名前で読み書きする | `MergeFields` などのコレクション | 名前やTagで項目を指定する |
| 文書にある項目を調べる | 各コレクションの列挙 | 名前・Tagと現在の値を取り出す |
| ファイルパスを使わず入力する | `Document.Open(Stream)` | 読み取り専用・非シーク入力も扱う。出力は現在ファイルパスのみ |
| OOXMLの構造を検証して開く | `Document.Open(..., validate: true)` | 名前の有無や業務データの検査とは別の検証 |

### パッケージ

現在は `.NET 10`（`net10.0`）が対象で、名前空間は `Marimo.DocumentAsData` です。
ライブラリの実行にWordのインストールは必要ありません。
初版候補は `0.1.0` で、現在は公開前です。ローカルNuGetパッケージまたはプロジェクト参照で利用します。

| パッケージ | 用途 |
| --- | --- |
| `Marimo.DocumentAsData.Build` | ビルド時のコード自動生成。以下の2パッケージを依存関係に含む |
| `Marimo.DocumentAsData.CodeGeneration` | プログラムから生成器を呼ぶ。Coreを依存関係に含む |
| `Marimo.DocumentAsData.Core` | DOCXの読み書き。コード生成を使わない場合はこれだけを参照 |

梱包・ローカルフィードへの登録・確認済みのビルド環境は、[ビルドとリリースの手引き](build-and-release.md) を参照してください。

## 型付きコード生成

### 生成される型

例えば `template.docx` からは、`TemplateDocument` と `TemplateData` が生成されます。
生成名を変更することもできます。

| 生成物 | 役割 |
| --- | --- |
| `TemplateDocument` | `Document` を継承し、文書を直接読み書きするプロパティを公開する |
| `TemplateData` | 文書のデータを保持する通常のオブジェクト。文書とは自動連動しない |

生成Documentのプロパティは、文字列・`bool`・`DateTimeOffset` の値を直接公開します。
MERGEFIELDやContent Controlのオブジェクトを取得するためのプロパティではありません。

```csharp
using MyDocuments;

using var document = TemplateDocument.Open("template.docx");

// 文書を直接変更します。
document.CustomerName = "株式会社○○";

// Read()の戻り値はTemplateDataです。
var data = document.Read();
data.Address = "東京都…";
document.Replace(data);

document.SaveAs("output.docx");
```

この例では、生成元テンプレートに `CustomerName` と `Address` の文字列項目が必要です。
両方の生成型は `partial` です。追記するコードは別ファイルに置き、生成ファイルは直接編集しないでください。

### 生成型で開く文書の条件

生成Documentには、ファイルパスとStreamをそれぞれ一つ受け取る `Open` が生成されます。
生成元にあった項目が開く文書にも存在するかを検査し、不足していれば `DocumentMappingException` を投げます。
例外メッセージには不足した名前と種類を含みます。この必須項目検査を無効にするオプションはありません。

文字列項目については、同名のMERGEFIELDと文字列Content Controlを互換として扱います。
生成プロパティは生成元と同じ種類を優先し、存在しない場合だけもう一方を使用します。
CheckBoxとDatePickerは、それぞれ同じTagの同じ種類が必要です。

生成型の `Read()`・`Replace(data)` は、通常のオブジェクト対応付けを使います。
生成プロパティの「元の種類を優先する」挙動とは異なり、同名のMERGEFIELDと文字列Content Controlが両方あると対応先を決められず例外になります。
テンプレートでは、生成する項目の名前を一意にしておくとこの違いを意識せずに使えます。

必須項目検査は、[OOXMLの検証](#ooxmlの検証)とは別です。生成型のOpenにOOXML検証オプションは生成しません。

### MSBuildで自動生成する

SDK形式の利用側プロジェクトで `Marimo.DocumentAsData.Build` を参照し、
生成対象のDOCXを `DocumentAsData` 項目として指定します。
ローカルパッケージの復元元を設定したうえで、例えば次のように記述します。

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <RootNamespace>MyDocuments</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Marimo.DocumentAsData.Build" Version="0.1.0" />
    <DocumentAsData Include="template.docx" />
  </ItemGroup>
</Project>
```

通常ビルドすると、文書の隣に `template.DocumentAsData.g.cs` が生成され、コンパイル対象へ追加されます。
生成コードの名前空間にはプロジェクトの `RootNamespace` を使用します。
パッケージ内のprops/targetsは自動で読み込まれ、DLLパスの指定や手動importは不要です。
`DocumentAsData` をビルドアクションの候補へ登録する設定も含みます。
コード生成には.NET 10 SDKを使用します。Visual Studio版MSBuildの確認範囲は[手引き](build-and-release.md)に記載しています。

#### 名前変換辞書

例えば `template.docx` に対して `template.documentasdata.json` を置きます。

```json
{
  "CustomerName": "ClientName"
}
```

文書と同じディレクトリを優先し、見つからなければプロジェクト直下を探します。
両方の辞書をマージする動作ではありません。
辞書の変更・削除を反映するには通常ビルドを実行してください。
生成名の衝突や不正なJSONは、対象ファイルを示すビルドエラーとして報告します。

#### ビルド・デザイン時・Cleanの動作

| 場面 | 動作 |
| --- | --- |
| 通常ビルド | コンパイル前に生成する。事前に生成ファイルを作る必要はない |
| 同じ生成内容でのビルド | ファイルを書き直さず、更新日時を維持する |
| プロジェクトの項目評価 | 生成済みソースを `Compile` へ登録する。評価だけでは生成しない |
| `DesignTimeBuild=true` | コンパイル経路でも再生成せず、生成済みソースを使う |
| Clean | 現在の `DocumentAsData` 項目に対応する生成ソースだけを削除する。元DOCXと手書きソースは残す |
| 生成対象からDOCXを外した後 | 古い生成ソースをコンパイル対象から除外するが、自動削除はしない |

通常評価とデザイン時の両方で、生成ソースを `DependentUpon` によって元DOCXへ紐付けます。
元DOCXには `LastGenOutput` として生成ファイル名を設定します。
SDK形式プロジェクトの既定の `Compile` 項目との二重登録を避け、生成ソースが `None`・`Content` 項目にも重複しないようにします。
これらのMSBuild連携テストは、Visual Studio全体の画面操作を保証するものではありません。

#### MSBuildタスクを直接利用する場合

通常のパッケージ利用では、この設定は不要です。

`Marimo.DocumentAsData.Build.GenerateDocumentAsData` タスクは、次の入出力を持ちます。

| 名前 | 内容 |
| --- | --- |
| `DocumentFiles` | 生成対象の `ITaskItem[]` |
| `ProjectDirectory` | 名前変換辞書を探すプロジェクトのディレクトリ |
| `RootNamespace` | 生成コードの名前空間。タスク単体の既定値は `Generated` |
| `GeneratedFiles` | 生成したファイルを示す出力項目 |

生成先は文書の隣の `<文書ファイル名の拡張子を除いた部分>.DocumentAsData.g.cs` です。
出力項目には `DependentUpon` と `DesignTimeSharedInput` メタデータを設定します。

ソースから利用する場合は、
`src/DocumentAsData.Build/buildTransitive/Marimo.DocumentAsData.Build.targets` を明示的にimportし、
`DocumentAsDataTaskAssembly` にビルド済みタスクDLLのパスを指定できます。
`DocumentAsData` 項目から `GenerateDocumentAsDataSources` ターゲットで生成します。
PowerShellから `dotnet msbuild` でこのターゲットを呼び出す動作もテストしています。

### プログラムから生成する

`Marimo.DocumentAsData.CodeGeneration` パッケージ、または
`src/DocumentAsData.CodeGeneration/DocumentAsData.CodeGeneration.csproj` を参照します。

```csharp
using System.IO;
using Marimo.DocumentAsData.CodeGeneration;

var sources = DocumentWrapperGenerator.GenerateSources(
    "template.docx",
    options => options.Namespace = "MyDocuments");

// 現在は一つの文書から一つのソースが返ります。
File.WriteAllText("Template.DocumentAsData.g.cs", sources[0]);
```

生成ソースをCoreを参照する利用側プロジェクトへ追加し、コンパイルして使います。
この方法では、出力ファイルの保存やビルドへの組み込みは呼び出し側で行います。

#### 生成名を指定する

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
`NameMappings` は自動変換より優先します。キーには生成後の名前ではなく、文書内の元名を指定します。
同じ元名への設定は、項目の種類を区別せず適用されます。辞書そのものを代入することもできます。

生成Dataプロパティ名が元名と異なる場合は `DocumentItemName` 属性が生成され、読み書き先を保持します。
書式文字を含む名前も、コンパイル後の識別子から書式文字が除かれるため、属性で元名を保持します。
文書側の名前やTagを書き換える機能ではありません。

#### 名前の衝突を診断する

`DocumentWrapperGenerator.GenerateDiagnostics(filePath, configure)` は、生成名の重複や
空の生成文書名を診断します。結果の `CodeGenerationDiagnostic` には次の情報があります。

| プロパティ | 内容 |
| --- | --- |
| `IsError` | エラーであるか |
| `GeneratedName` | 診断対象の生成後のC#名 |
| `SourceNames` | その生成名に対応する文書上の元名 |
| `InvalidSourceName` | 有効な文書名へ変換できなかった元名。該当しない場合は `null` |

異なる項目が同じC#名になる場合や、生成型の予約されたメンバー名などと衝突する場合は、
テンプレートまたは `NameMappings` を調整してください。自動で連番を付けて衝突を解消することはしません。
診断と生成には同じ設定を渡します。

`GenerateSources` は診断を自動実行しません。MSBuildタスクは生成前に診断します。
診断はすべてのC#コンパイルエラーを検出する機能ではないため、生成後は利用側プロジェクトでビルドしてください。

## オブジェクトとの対応付け

`Document.Read<T>()` は文書から新しいオブジェクトへ読み込み、
`Document.Replace(data)` はオブジェクトのプロパティを文書へ反映します。
生成型の `Read()`・`Replace(data)` もこの対応付けを利用します。

| プロパティの型 | 対応する文書項目 |
| --- | --- |
| `string` | MERGEFIELDまたは文字列Content Control |
| `bool` | チェックボックスContent Control |
| `DateTimeOffset` | 日付選択Content Control |

自分で用意する読み込み先のクラスは、引数なしで生成でき、対象プロパティへ値を設定できる形にしてください。
例えば、引数なしのコンストラクターと `public` な `get; set;` を持つクラスを使用します。
読み取ったオブジェクトは文書と連動しません。変更後は `Replace(data)`、その後 `Save()` または `SaveAs(path)` が必要です。

### 名前を解決する順序

1. プロパティに `DocumentItemName` 属性があれば、その名前を使います。
2. 属性がなければ、文書内の名前・Tagをコード生成と同じ規則でC#識別子へ変換し、プロパティ名と照合します。
3. 変換後の名前に該当がなければ、プロパティ名をそのまま文書内の名前・Tagとして使います。

例えば文書の `customerName` は、自動変換でDTOの `CustomerName` に対応します。
名前の候補には4種類のコレクションにある項目を使います。
同じプロパティ名に複数の候補が一致する場合は例外になります。
また、文字列の対応先として同名のMERGEFIELDと文字列Content Controlが両方あれば例外になります。

```csharp
using Marimo.DocumentAsData;

public sealed class CustomerData
{
    [DocumentItemName("CustomerName")]
    public string Name { get; set; } = "";

    public string Address { get; set; } = "";
}
```

この例では `Read<CustomerData>()` は文書の `CustomerName` を `Name` へ読み込み、
`Replace(data)` は `data.Name` を同じ文書項目へ書き込みます。
属性のない `Address` は自動名前対応を使います。

### Replace時の注意

* オブジェクトに対応プロパティのない文書項目は変更しません。
* 指定した文書項目が存在しない場合、対応していない型の場合、複数のプロパティが同じ項目へ対応する場合は `DocumentMappingException` になります。
* `DocumentItemName` 属性のない、publicなgetterを持たないプロパティは書き込み元にしません。
* `DocumentItemName` 属性で指定したプロパティにpublicなgetterがない場合は `DocumentMappingException` になります。

## コレクションと項目

対象の探索範囲は本文です。各コレクションは列挙と、名前を指定するインデクサーを提供します。
Content ControlはTagで識別します。Word画面のタイトルを検索するAPIではありません。

| 項目 | コレクション | 識別用プロパティ | 値のプロパティと型 |
| --- | --- | --- | --- |
| MERGEFIELD | `MergeFields` | `Name` | `Text`：`string` |
| 文字列Content Control | `ContentControls` | `Tag` | `Text`：`string` |
| チェックボックス | `CheckBoxes` | `Tag` | `IsChecked`：`bool` |
| 日付選択Content Control | `DatePickers` | `Tag` | `SelectedDateTime`：`DateTimeOffset` |

```csharp
foreach (var field in document.MergeFields)
{
    Console.WriteLine($"{field.Name}: {field.Text}");
}

document.MergeFields["CustomerName"].Text = "株式会社○○";
document.ContentControls["Address"].Text = "東京都…";
document.CheckBoxes["Agreement"].IsChecked = true;
```

インデクサーは各コレクション内で一致する項目が一つであることを要求します。
該当なしは `KeyNotFoundException`、複数ある場合は `InvalidOperationException` です。
同名の全項目へ一括設定する動作ではありません。
TagのないContent Controlはコレクションへ含めません。ネストした対象Content Controlはそれぞれ列挙します。
各項目の `Document` プロパティで所属文書を参照できます。

### MERGEFIELDと文字列

* MERGEFIELDは単純フィールドと、Wordが通常作る複合フィールドに対応します。
* `Text` は保存された表示結果を読み書きします。Word標準の差し込み印刷、CSV連携、フィールド計算は実行しません。
* `ContentControls` は文字列用だけのコレクションです。プレーンテキストとリッチテキストを扱います。
* プレースホルダー表示中の文字列Content Controlは、案内文ではなく空文字列として読み取ります。値を書き込むとプレースホルダー表示を解除します。
* 一つの項目内で文字列が複数のrunに分かれていても値をまとめて扱います。runはWordが書式などの単位で文字を分割して保持する内部要素です。
* 部分ごとの書式を指定するAPIはありません。リッチテキストを、書式付きデータとして読み書きする機能ではありません。

### チェックボックスと日時

チェックボックスはWordのContent Controlが対象です。表示文字の解析ではなく、チェック状態を `bool` として扱います。
自由入力されたチェック記号を判定する機能ではありません。

日付選択Content Controlでは、OOXMLに記録された日時をその時差も含めて `DateTimeOffset` として扱います。

```csharp
document.DatePickers["DeliveryDate"].SelectedDateTime = new DateTimeOffset(
    2027, 1, 2, 0, 0, 0, TimeSpan.FromHours(9));
```

表示文字列はテンプレートの表示形式と表示言語に従って更新します。
日時が未入力でも、表示形式・表示言語・表示用の文字列要素があれば日時を設定できます。
表示形式は、.NETと共通の `yyyy/MM/dd` と、OOXMLの仕様例の `MM-YYYY`、引用文字列を含む `'YYYY' YYYY/MM/dd` をテストしています。
大文字の年指定を.NET用の小文字へ変換し、単一引用符内の文字列は保持します。
Word独自の表示形式すべてとの互換性はありません。Word実機での表示確認ではなく、保存したOOXMLの表示文字列をテストしています。
`DateOnly` へ自動変換したり、日本標準時へ一律に変換したりはしません。
誕生日・期日などの業務上の意味付けは利用側で行ってください。

## 開く・保存する・閉じる

変更は作業領域に保持されます。`Close()`・`Dispose()`・`using` の終了だけでは元データへ保存しません。

| 開き方 | `Save()` | `SaveAs(path)` | Close・Dispose時 |
| --- | --- | --- | --- |
| `Open(string)` | 元ファイルへ保存 | 別ファイルへ保存。保存元は変えない | 未保存の変更を反映せず、元ファイルの束縛を解除 |
| `Open(Stream)` | `NotSupportedException` | 別ファイルへ保存。入力内容は変えない | 元Streamの内容を変更せず、Stream自体も閉じない |

パス版は元ファイルを開いている間束縛します。`Save()` は元ファイルへ書き戻し、`SaveAs(path)` は保存先を開いたまま保持しません。
どちらも呼び出しが戻った時点で保存されているため、DocumentのDisposeを待たずに保存した内容を読めます。
`SaveAs` 後の編集を別のファイルへ再び `SaveAs` することもできます。

### 入力Streamの契約

* 入力には同期読み取り（`Read`）の能力が必要です。固定容量の `MemoryStream`、読み取り専用Stream、シークできないStreamを使用できます。
* 呼び出した時点の現在位置をDOCXの先頭として扱い、そこから末尾までを読み込み対象とします。先頭に別のデータがある場合は、DOCXの開始位置まで進めてから渡します。
* 入力位置は読み取りに伴って進みます。再利用する際は呼び出し側で位置を戻してください。
* 入力Streamは呼び出し側が所有します。Documentを使用している間は閉じず、使用後に呼び出し側でDisposeしてください。
* 元Streamへ上書きしません。固定容量・拡張可能かによらず、`Save()` は書き込み前に `NotSupportedException` を投げます。
* 現在位置以降にデータがない入力は `InvalidDataException` で拒否します。Openが失敗しても元Streamを閉じたり、内容を書き換えたりしません。
* 読み取り専用・非シークの入力でも文書を編集して `SaveAs(path)` で保存できます。非シーク入力はライブラリ内でメモリへコピーするため、呼び出し側で編集用Streamを作る必要はありません。
* `SaveAs(path)` が保存先を開けずに失敗した場合も、元Streamと編集内容は保持され、別の保存先へ再試行できます。

### ブラウザでの利用範囲

本体はBlazorの `IBrowserFile` には依存せず、通常の `Stream` を受け取ります。
ブラウザでのダウンロードやHTTPレスポンスの生成は利用側の責務です。
現在はStreamへの出力APIがないため、Streamだけで編集結果の出力まで完結するブラウザ用途には未対応です。
Blazor WebAssembly上での実動作は、このリポジトリのテストでは確認していません。
非同期の `OpenAsync` APIも現在提供していません。

## OOXMLの検証

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open("template.docx", validate: true);
```

Stream版にも `validate` 引数があります。省略時は検証を実行しません。
`true` の場合はOpen XML SDKのValidatorで検証し、検証エラーがあれば `InvalidDataException` を投げます。
検証エラーのない文書であることは、表示・レイアウトの正しさや、業務上の値の妥当性を保証しません。
生成型の必須項目検査とも別の機能です。

## 非対応と注意点

### テンプレートで避けるもの

`w:dataBinding` によりCustom XMLへデータバインドされたContent Controlは未対応です。
表示内容とは別にCustom XML側にも値が保持されますが、DocumentAsDataは現在その同期を行いません。
読み書き対象として使用しないでください。

コンボボックス、ドロップダウン、画像、繰り返しセクションなどを文字列Content Controlとしては扱いません。
繰り返し行の追加・削除APIもありません。
文字列項目に別の種類の機能があることを期待せず、対応する項目をテンプレートで明示してください。

### ライブラリの対象外

* ヘッダー・フッター内の項目の読み書き
* ODTの読み書き
* PDFへのレンダリング、ページ・行・レイアウトの計算や編集
* Content Controlの除去・アンラップ、MERGEFIELDの通常テキスト化
* OCR、自然言語解析、自由入力文書から氏名や住所などを推測して抽出する処理

保存時はMERGEFIELDやContent Controlの構造を残します。
PDFが必要な場合は、生成したDOCXをWordやLibreOfficeなどで変換してください。変換は利用側の責務です。
PDF化のためだけにContent Controlを除去する機能は提供しません。

自動テストでは固定DOCX、保存後の再読込、OOXML構造、生成C#のコンパイルなどを確認していますが、
すべてのWord文書や表示・印刷結果を保証するものではありません。
このページのコード例は、各例で指定した名前付き項目がある文書を前提にしています。
