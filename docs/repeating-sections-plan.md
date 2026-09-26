# 繰り返しセクションの対応計画

この文書は未実装機能のレビュー用です。以下のAPIは案であり、現時点では利用できません。
実装予定はSkip付きテストで管理し、一件ずつSkipを外してRedを確認してからGreenへ進めます。

## 今回の範囲

Word側で繰り返しセクションとして設定された領域を、明細のコレクションとして扱います。
明細が1件でも、複数件用のAPIであることは変わりません。
単に同じTagが複数ある文書から、繰り返しを推測する機能ではありません。

```csharp
var section = document.RepeatingSections["Items"];
var name = section.Items[0].ContentControls["ItemName"].Text;

var data = section.Read<ItemData>().ToArray();
data[0].ItemName = "変更後";
section.Replace(data);
```

* `RepeatingSections` は繰り返し全体のTagで検索します。
* `Items` は文書順の明細です。明細自身にTagがあることは要求しません。
* 明細内には既存の `ContentControls`、`MergeFields`、`CheckBoxes`、`DatePickers` に対応する検索範囲を持たせます。
* 異なる明細の同じTagは区別します。同じ明細内のTag重複はエラーにします。
* `Read<T>()` は各明細を自作クラスへ対応付け、`Replace<T>(IEnumerable<T>)` は内容と件数を置き換えます。
* 型付きコード生成では明細データを列挙するプロパティと `Replace` を提供する案です。明細データの変更は `Replace` するまで文書へ反映しません。

## 一件用APIとの比較

既存の `document.ContentControls["ItemName"]` は文書全体から検索します。
繰り返しの内側でも、そのTagが文書全体で1件なら取得でき、複数なら例外になります。
新APIは明細ごとに検索します。既存APIの検索範囲を暗黙に変える計画にはしていません。

生成されたテンプレート専用プロパティでは、文書直下の項目と明細内の項目を別々に扱う案をSkipテストで提示しています。
このため、生成コード側にも項目の所属範囲を区別する対応が必要です。

## Skipテストの内訳

合計50件です。49件は期待結果を記述した未レビューの仕様案、1件は契約が未決定の検討メモです。
最初の列挙テスト1件はすでにRedを確認してからSkipへ戻しており、今回追加した分は49件です。
Skipを外す際には、そのテストが要求するAPIと振る舞いを改めてレビューします。

| 対象 | 件数 | 主な観点 |
| --- | ---: | --- |
| [RepeatingSectionCollection](../tests/DocumentAsData.Tests/RepeatingSectionCollectionのテスト.cs) | 7 | 1件でも列挙、複数セクションの順序、通常項目との区別、Tag検索・不足・重複・Tagなし |
| [RepeatingSectionItem](../tests/DocumentAsData.Tests/RepeatingSectionItemのテスト.cs) | 9 | 明細内の検索範囲、同名の読み書き、既存4種類への接続、親がXMLマッピング済みの場合の拒否 |
| [RepeatingSection](../tests/DocumentAsData.Tests/RepeatingSectionのテスト.cs) | 18 | 明細列挙、Read/Replace、名前属性、項目不足、増減、変更後の再取得、保存、表行構造・書式・ID、XMLマッピング拒否、0件の検討メモ |
| [既存ContentControlCollectionとの比較](../tests/DocumentAsData.Tests/ContentControlCollectionのテスト.cs) | 2 | 繰り返し内でも文書全体で一意なら取得でき、複数なら例外になる |
| [コード生成・名前付き項目](../tests/DocumentAsData.CodeGeneration.Tests/コード生成名前付き項目のテスト.cs) | 9 | 1件と複数件、生成元と実行時の件数差、Replace、通常項目との名前の区別、値のスナップショット、既存の値型 |
| [コード生成・型構造](../tests/DocumentAsData.CodeGeneration.Tests/コード生成型構造のテスト.cs) | 2 | Open時のセクション不足、明細内の必須項目不足 |
| [コード生成・名前設定](../tests/DocumentAsData.CodeGeneration.Tests/コード生成名前設定のテスト.cs) | 3 | セクションと明細プロパティの名前変更、元名への書き戻し、明細内の名前衝突 |

値変換の全組み合わせ、既存のファイル・Stream保存契約、MSBuildの動作を、この機能のために繰り返しテストする計画にはしていません。
今回のテストは既存の読み書き機能と明細の範囲が正しく接続されることを中心にしています。

## 未決定・今回の対象にしていないこと

* `Replace` に0件を渡した場合の文書表現、明細の再追加に使う複製元、保存して開き直した後の再利用方法は未決定です。対応するSkipメソッドは期待結果を持たない検討メモであり、実装量を見積もれる仕様テストとは分けて扱います。
* 文書全体の `Read<T>()` と `Replace(data)` にリストプロパティを対応付ける機能は、明細APIが成立した後に検討します。生成Documentの `Read()` と `Replace(data)` に明細リストを組み込むことも、今回の計画には含めていません。
* 繰り返しセクションの入れ子、明細ごとに異なる項目構成、画像・ブックマーク等を含む任意の文書構造の複製は、今回の基本対応には含めていません。将来対応時には、対応する範囲と拒否の契約も明示的にレビューします。
* 個別の `Add`・`Remove` APIはまだ計画していません。まず件数を指定した `Replace` を対象にします。
* XMLマッピングされたセクションは、連結先を同期する機能を追加せず、読み書きを拒否する案です。

## 固定テストデータ

テスト実行時にDOCXを生成しません。以下の10ファイルを追加し、既存の「繰り返しセクションのContent Control.docx」（明細1件）と併用します。
いずれもOOXMLの構造を確認するための最小データです。Open XML SDKのOffice 2013向け検証でエラーがないことを確認しました。Wordでの表示・編集確認は未実施です。

| ファイル名 | 内容 |
| --- | --- |
| 繰り返しセクションに2件の明細.docx | Itemsに2件。各明細のItemNameは商品A・商品B |
| 異なるTagの繰り返しセクション.docx | ItemsとOptionsがこの順で存在 |
| 同じTagの繰り返しセクション.docx | ItemsというTagのセクションが2つ存在 |
| Tagなしを含む繰り返しセクション.docx | Itemsと、Tagのないセクション |
| 明細の内外に同じTag.docx | 外側のItemNameと、Items内の2件のItemName |
| 1件の明細内に同じTag.docx | 明細1件の中にItemNameが2つ存在 |
| 明細に4種類の項目.docx | 各明細にItemName、MERGEFIELDのCode、Agreement、DeliveryDate。チェックは両方true |
| 繰り返しセクションの明細に項目が不足.docx | 1件目にItemNameがあり、2件目にはない |
| 表行を繰り返す2件の明細.docx | 見出し行、太字の明細2行、表外の後書き。Content Controlの識別ID付き |
| XMLマッピングされた繰り返しセクション.docx | 外側のItemsにdataBindingがあり、内側のItemName自体にはない。Custom XML部品を含む |

コード生成テストからはプロジェクトのリンクで同じ固定ファイルを参照し、データを二重管理しません。
そこで使用する `repeatingTemplate.docx` 等の名前は、生成型名を固定するためのリンク名です。

## 変更対象のファイル

* 上表のテスト7ファイル。
* [RepeatingSection.cs](../src/DocumentAsData/RepeatingSection.cs)：Items・Read・Replaceの公開APIスタブ。
* [RepeatingSectionItem.cs](../src/DocumentAsData/RepeatingSectionItem.cs)：明細内の各種コレクションの公開APIスタブ。
* [コード生成テストのプロジェクト](../tests/DocumentAsData.CodeGeneration.Tests/DocumentAsData.CodeGeneration.Tests.csproj)：固定DOCXへのリンク。
* [README](../README.md)とこの計画書：未実装であること、レビュー対象、未決定事項を説明。
* `tests/DocumentAsData.Tests/TestData/` 配下の上記DOCX10ファイル。

本体に追加したAPIの処理は `NotImplementedException` のままです。
コード生成本体への先行実装は行っていません。
