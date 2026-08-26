# DocumentAsData v0.1 設計案

## 目的と範囲

v0.1 は DOCX 本文にある MERGEFIELD と、Tag を持つテキスト系 Content Control の
読み書きに限定する。Open XML SDK の型は公開APIへ露出させない。

次の項目は後続バージョンで必要性を確認してから扱う。

- ヘッダー、フッター、脚注、文末脚注、コメントなど本文以外の story
- 画像、チェックボックス、日付、リストなどテキスト以外の Content Control
- Content Control の除去、アンラップ
- MERGEFIELD の通常テキスト化
- ODTなどDOCX以外の形式

本文中の表は本文ツリーに含まれるため、特別扱いせず検索対象に含める。ネストされた
Content Control は各 Tag を独立した項目として列挙する。ただし、親の値を書き換えると
子の表示内容にも影響し得るため、同じ操作で親子の両方を書き換える利用方法は v0.1 の
保証対象外とする。

## 公開API案

```csharp
using Marimo.DocumentAsData;

using var document = Document.Open("template.docx");

document.MergeFields["CustomerName"].Value = "株式会社○○";
document.ContentControls["CustomerName"].Value = "山田太郎";

document.SaveAs("output.docx");
```

```csharp
namespace Marimo.DocumentAsData;

public sealed class Document : IDisposable
{
    public static Document Open(string path);
    public MergeFieldCollection MergeFields { get; }
    public ContentControlCollection ContentControls { get; }
    public void SaveAs(string path);
    public void Dispose();
}

public sealed class MergeFieldCollection :
    IReadOnlyList<MergeField>,
    IReadOnlyDictionary<string, MergeField>
{
    public MergeField this[int index] { get; }
    public MergeField this[string name] { get; }
    public bool ContainsKey(string name);
    public bool TryGetValue(string name, out MergeField? field);
}

public sealed class MergeField
{
    public string Name { get; }
    public int Count { get; }
    public string Value { get; set; }
    public IReadOnlyList<string> Values { get; }
}

public sealed class ContentControlCollection :
    IReadOnlyList<ContentControl>,
    IReadOnlyDictionary<string, ContentControl>
{
    public ContentControl this[int index] { get; }
    public ContentControl this[string tag] { get; }
    public bool ContainsKey(string tag);
    public bool TryGetValue(string tag, out ContentControl? control);
}

public sealed class ContentControl
{
    public string Tag { get; }
    public int Count { get; }
    public string Value { get; set; }
    public IReadOnlyList<string> Values { get; }
}
```

コレクションは名前またはTagごとの論理項目を列挙する。比較は大文字・小文字を区別する。
存在しないキーをインデクサーへ渡した場合は `KeyNotFoundException` とし、分岐が必要な
場合は `TryGetValue` を使う。

同じ名前またはTagが複数ある場合、`Count` は出現数、`Values` は文書順の全現在値を返す。
`Value` の取得は文書内の最初の値を返し、設定は全出現箇所へ同じ値を書き込む。このため、
帳票生成の主要用途を簡潔に保ちつつ、入力値が食い違う文書も `Values` で検出できる。

`Open` は元ファイルをメモリへ読み込んで編集し、元ファイル自体を変更しない。v0.1では
意図しない上書きを避けるため `Save` は設けず、`SaveAs` のみを提供する。同じインスタンス
から複数回 `SaveAs` できる設計とする。

## MERGEFIELD の扱い

- `w:fldSimple` の `w:instr` と、`begin / instrText / separate / result / end` 形式の
  複合フィールドを扱う。
- 命令名 `MERGEFIELD` は大文字・小文字を区別せず認識する。
- フィールド名は引用符付き、引用符なしの双方を扱い、表示書式など後続スイッチは名前に
  含めない。
- 読み取り値は現在の field result のテキストとする。
- 書き込み時もフィールド構造と命令を残し、field result の表示テキストだけを更新する。
- 不完全な複合フィールドはデータ項目として列挙せず、文書内容も変更しない。

## Content Control の扱い

- `w:sdtPr/w:tag` を識別子とし、Title/Alias は識別に使わない。
- 値は `w:sdtContent` 内の表示テキストとして読み書きする。
- v0.1 の書き込み対象はテキストとして表現できる Content Control に限定する。
- 書き込み後も Content Control、Tag、その他のプロパティは残す。
- Tag がない Content Control は列挙しない。

## 内部構成案

SpreadsheetAsData の「名前付き要素をコレクションとインデクサで扱う」「Open XML SDKを
internalへ閉じる」「日本語のXMLドキュメントとテスト名を使う」という方針を引き継ぐ。
名前空間も同系列の `Marimo.DocumentAsData` とする。一方、コード生成、型付きDTO、複合
NuGetパッケージは今回の目的に不要なため移植しない。

```text
src/DocumentAsData/
  Document.cs
  MergeFields/
    MergeField.cs
    MergeFieldCollection.cs
    MergeFieldScanner.cs
  ContentControls/
    ContentControl.cs
    ContentControlCollection.cs
    ContentControlScanner.cs
  Internal/
    WordprocessingDocumentStore.cs
tests/DocumentAsData.Tests/
  DocumentTests.cs
  MergeFields/
    MergeFieldCollectionTests.cs
    MergeFieldWriteTests.cs
  ContentControls/
    ContentControlCollectionTests.cs
    ContentControlReadWriteTests.cs
  TestDocuments/
    DocumentBuilder.cs
```

Scanner と Store は内部実装とし、利用側には Open XML SDK の要素、Part、Package を返さない。

## TDDの進め方

テストは移植元と同じく xUnit と FluentAssertions を使用し、日本語のテスト名で仕様を表す。
テスト用DOCXはテスト内のビルダーで最小の Open XML を生成する。固定バイナリだけに依存
せず、何を検証する文書かをテストコードから読めるようにする。保存後は再度DOCXとして
開き直し、公開APIから見える値に加えて Open XML Validator でも妥当性を確認する。

実装順は次の Red / Green 単位とする。

1. DOCXを開き、別パスへ保存でき、元ファイルが変化しない。
2. simple MERGEFIELDを列挙し、名前検索と現在値の読み取りができる。
3. complex MERGEFIELDを列挙できる。命令が複数の `instrText` に分割された場合も扱う。
4. MERGEFIELDへの書き込みが field result のみを変更し、保存後もフィールドを維持する。
5. 同名MERGEFIELDの `Value` 設定が全出現箇所を更新し、`Values` が文書順になる。
6. Tag付きContent Controlを列挙し、値を読み取れる。Tagなしは除外する。
7. Content Controlへ書き込み、保存後もTagとコントロールを維持する。
8. 同一Tag、表内、ネストの基本ケースを追加する。

各段階で、対象外のフィールド、存在しないキー、空文字、Unicode、空の結果、壊れた複合
フィールドを必要な範囲で境界値として追加する。まず失敗理由を確認してから最小実装を行い、
Green後にだけ重複を整理する。

## 実装開始前に確定する事項

- v0.1 のターゲットは SpreadsheetAsData と揃えて `net10.0` とする。
- Open XML SDK は直接依存として追加するが、公開シグネチャには含めない。
- v0.1 の文書範囲は main document part の本文とし、表内は含める。
- 重複キーの取得・設定規則は上記の `Value` / `Values` とする。
