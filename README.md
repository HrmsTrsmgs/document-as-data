# DocumentAsData

Word文書（DOCX）の名前付き項目を、C#のプロパティとして読み書きする.NETライブラリです。
Open XML SDKの型やXML構造を扱わずに、文書とプログラムの間でデータを交換できます。

たとえば、テンプレートから生成した型を使うと、次のように書けます。

```csharp
using var document = TemplateDocument.Open("template.docx");

Console.WriteLine(document.CustomerName);
document.CustomerName = "株式会社サンプル";
document.Address = "大阪府";

document.SaveAs("output.docx");
```

`TemplateDocument` とそのプロパティは、DOCX内のフィールド名やTagから生成されます。
手書きのクラスや、名前を文字列で指定するAPIも使えます。

## どんな文書に使えるか

* **帳票を作る**：Wordの差し込みフィールド（MERGEFIELD）へ、プログラムから値を書き込む。
* **入力済みの文書を読む**：Content Control（コンテンツコントロール）に人が入力した内容を、プログラムへ取り込む。
* **フォームの値を文書へ戻す**：文字列、チェック状態、日付選択の値を読み書きする。

基本契約は「データとして扱いたい場所に、Word側で明示的な構造と名前を付ける」ことです。
自由な文章から氏名や住所を推測して抽出するライブラリではありません。
ライブラリの実行にWordのインストールは不要です。PDF化やレイアウト計算は行いません。

## はじめて使う：コード生成でDOCXを読み書きする

現在の対象は **.NET 10** です。初版候補の `0.1.0` はまだ公開していません。
以下は、ソースからローカルパッケージを作って試す手順です。
NuGet.orgから導入できることを前提にはしていません。

### 1. パッケージとサンプルを用意する

.NET SDKはリポジトリの `global.json` に合わせて用意してください（現在は `10.0.302`）。
以下はWindowsのPowerShellで、取得したリポジトリのルートから実行する例です。
最初の復元では依存パッケージを取得するためにネットワークを使用します。

```powershell
dotnet restore DocumentAsData.slnx

dotnet pack src/DocumentAsData/DocumentAsData.csproj -c Release --no-restore -o artifacts/local-packages
dotnet pack src/DocumentAsData.CodeGeneration/DocumentAsData.CodeGeneration.csproj -c Release --no-restore -o artifacts/local-packages
dotnet pack src/DocumentAsData.Build/DocumentAsData.Build.csproj -c Release --no-restore -o artifacts/local-packages

dotnet new console --framework net10.0 --name DocumentAsDataDemo --output artifacts/DocumentAsDataDemo
Copy-Item "tests/DocumentAsData.Tests/TestData/単純形式のMERGEFIELD.docx" "artifacts/DocumentAsDataDemo/template.docx"
```

このDOCXには、`CustomerName`（初期値：株式会社○○）と
`Address`（初期値：東京都）の2つのMERGEFIELDがあります。
最初はWordでテンプレートを作らずに試せます。

### 2. コード生成を設定する

`artifacts/DocumentAsDataDemo/DocumentAsDataDemo.csproj` を次の内容にします。

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net10.0</TargetFramework>
    <RootNamespace>DocumentAsDataDemo</RootNamespace>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Marimo.DocumentAsData.Build" Version="0.1.0" />
    <DocumentAsData Include="template.docx" />
    <None Update="template.docx" CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>
</Project>
```

`DocumentAsData` はコード生成の対象、`None` の設定は実行時に使うDOCXのコピーを指定します。
`Marimo.DocumentAsData.Build` には、読み書き用とコード生成用のパッケージも依存関係として含まれます。

同じフォルダーに `NuGet.Config` を作ります。
このサンプルだけで使うパッケージ取得先の設定です。

```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <clear />
    <add key="DocumentAsDataLocal" value="../local-packages" />
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
  </packageSources>
  <config>
    <add key="globalPackagesFolder" value="../demo-package-cache" />
  </config>
</configuration>
```

公開前の同じバージョンを繰り返し試す際、以前のパッケージと混ざらないように、サンプル用のキャッシュを分けています。
後日同じ `0.1.0` を作り直す場合は、`globalPackagesFolder` を新しいフォルダー名に変えて復元してください。

### 3. 値を読み書きする

`Program.cs` を次の内容にします。

```csharp
using DocumentAsDataDemo;

using var document = TemplateDocument.Open(
    Path.Combine(AppContext.BaseDirectory, "template.docx"));

Console.WriteLine(document.CustomerName);

document.CustomerName = "株式会社サンプル";
document.Address = "大阪府";

var outputPath = Path.Combine(AppContext.BaseDirectory, "output.docx");
document.SaveAs(outputPath);

Console.WriteLine(outputPath);
```

リポジトリのルートから実行します。

```powershell
Set-Location artifacts/DocumentAsDataDemo
dotnet run
```

通常のビルド時に `template.DocumentAsData.g.cs` がDOCXと同じフォルダーへ生成され、
`TemplateDocument` と、データ受け渡し用の `TemplateData` がコンパイル対象になります。
生成されたファイルを手で編集する必要はありません。
Visual Studioで生成型がまだ表示されない場合も、まず通常のビルドを行ってください。

### 4. 結果を確認する

コンソールに元の値 `株式会社○○` と、保存した `output.docx` のフルパスが表示されます。
出力先はこの例では `bin/Debug/net10.0/output.docx` です。
Wordなどで開くと、名前が `株式会社サンプル`、住所が `大阪府` に変わっています。

元の `template.docx` は変更しません。
プロパティへの代入は開いている文書に反映され、`SaveAs` でファイルとして保存します。
`Dispose` や `using` の終了では暗黙に保存しません。

## 次に読む

1. [チュートリアル](https://github.com/HrmsTrsmgs/document-as-data/blob/main/docs/tutorial.md)：自分のテンプレート、生成したデータ型、手書きの型、名前指定のAPIの順に進みます。
2. [機能リファレンス](https://github.com/HrmsTrsmgs/document-as-data/blob/main/docs/reference.md)：項目の種類、保存とStream、検証、コード生成、MSBuild設定、制限事項を調べます。
3. [ビルドとリリースの手引き](https://github.com/HrmsTrsmgs/document-as-data/blob/main/docs/build-and-release.md)：ライブラリ開発や公開前検証の手順です。

ソースを取得している場合は、リポジトリ内の `docs/tutorial.md`、
`docs/reference.md`、`docs/build-and-release.md` でも読めます。

## 利用方法を選ぶ

| やりたいこと | 利用するAPI・パッケージ |
| --- | --- |
| DOCXから型を生成し、プロパティで読み書きしたい | `Marimo.DocumentAsData.Build` |
| 自分の処理からC#コードを生成したい | `DocumentWrapperGenerator`：`Marimo.DocumentAsData.CodeGeneration` |
| 手書きの型や項目名で読み書きしたい | `Document`：`Marimo.DocumentAsData.Core` |

コード生成を使わない場合は、たとえば次のように同じ文書を扱えます。
名前空間は `Marimo.DocumentAsData` です。

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open("template.docx");
document.MergeFields["CustomerName"].Text = "株式会社サンプル";
document.MergeFields["Address"].Text = "大阪府";
document.SaveAs("output.docx");
```

## 利用前に知っておくこと

* 対象はDOCX本文です。ヘッダー・フッター内の項目は対象外です。
* 同じコレクション内で名前やTagが重複する項目を、まとめて更新する仕様ではありません。名前で一意に取得できない場合は例外になります。
* 生成型の `Open` は、その型が必要とする項目を検査します。単にDOCXなら何でも同じ型で開けるわけではありません。
* `ContentControls` は文字列用です。チェックボックスは `CheckBoxes`、日付選択は `DatePickers` で扱います。
* `Open(Stream)` は入力用Streamを借用し、内容を変更したり閉じたりしません。Streamで開いた文書の `Save()` は使用できません。現時点の保存先はファイルパスであり、`SaveAs(Stream)` はありません。
* MERGEFIELDやContent Controlを普通のテキストに変換・除去する処理、PDF化、レイアウト編集、OCR、自由文章からの推測抽出は行いません。
* カスタムXMLとデータ連結されたContent Control、選択リスト、画像コントロール、繰り返し項目などには制約があります。詳細は機能リファレンスを参照してください。

## 開発する

```powershell
dotnet restore DocumentAsData.slnx
dotnet build DocumentAsData.slnx --no-restore
dotnet test DocumentAsData.slnx --no-restore
```

テスト補助スクリプトにはPowerShell 7（`pwsh`）を使用します。
アプリからライブラリを使うためにPowerShell 7が必要なわけではありません。

## ライセンス

MIT。詳細は [LICENSE](https://github.com/HrmsTrsmgs/document-as-data/blob/main/LICENSE) を参照してください。
