# 繰り返しセクションの対応計画

この文書は繰り返しセクション対応のレビュー用です。現在はセクションの列挙・Tag検索、明細の取得、明細ごとの文字列Content Control・MERGEFIELD・チェックボックス・日付選択の読み書きまで確認しています。
`Read<T>()` は明細ごとの読み取り・名前属性・項目不足を、`Replace<T>()` は同件数の書き込み・名前属性・4種類への接続と、1件以上の明細の増減の基本ケースまで確認しています。変更後の再取得・保存、表行の構造・書式の保持、複製時のContent Control識別ID調整も確認しています。コード生成は文字列Content Controlを持つ単数・複数の明細を列挙する基本ケースまで対応しています。0件の契約は未決定です。
実装予定はSkip付きテストで管理し、一件ずつSkipを外してRedを確認してからGreenへ進めます。
精査で見つかった問題点と未確認事項は、末尾の「精査で確認した残課題」で管理します。これらは、上記の予定テストを消化したことだけでは解決済みになりません。

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
* `Read<T>()` は各明細を自作クラスへ対応付けます。`Replace<T>(IEnumerable<T>)` は内容と件数を置き換えます。現在は同件数の内容変更と、1件以上の明細の増減の基本ケースを確認しています。
* 型付きコード生成では `RepeatingSection<T>` を返すプロパティから、明細データの列挙と `Replace` ができます。文字列Content Controlの読み取り・書き込みと、MERGEFIELD・チェックボックス・日付選択の読み取りを確認しています。取得した明細データの変更は、`Replace` するまで文書へ反映されません。

## 一件用APIとの比較

既存の `document.ContentControls["ItemName"]` は文書全体から検索します。
繰り返しの内側でも、そのTagが文書全体で1件なら取得でき、複数なら例外になります。
新APIは明細ごとに検索します。既存APIの検索範囲を暗黙に変える計画にはしていません。

生成されたテンプレート専用プロパティでは、文書直下と明細内の同名の文字列Content Controlを別々に読み取れることを確認しています。
同名が内外にある場合の、文書直下の生成プロパティへの書き込みは未対応です。

## 予定テストの内訳

合計50件のうち、セクションの列挙・Tag検索7件、明細取得1件、明細内の文字列Content Control4件、MERGEFIELD1件、チェックボックス1件、日付選択1件、親セクションのXMLマッピングによる文字列読み書き拒否2件、自作クラスへの読み取り3件、同件数の書き込み5件、明細追加1件、明細削減1件、変更後の再取得1件、保存1件、表行構造・書式1件、識別ID1件が成功しています。
さらにXMLマッピング済みセクションのRead・Replace拒否2件、文書全体検索との比較2件、生成された明細の列挙3件、生成データによるReplace1件、通常項目との同名併存の読み取り1件、明細データが文書と直接連動しないこと1件、明細内MERGEFIELD・チェックボックス・日付選択の生成各1件、名前変換の読み書き2件、明細内の重複名診断1件、Open時の不足検査2件が成功し、成功は合計49件です。
このうち20件はRedからGreenにし、残り29件は既存実装で成功しました。
残り1件はSkip中で、0件の契約が未決定の検討メモです。予定テストの消化は、未検証・未対応部分の解消を意味しません。
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
* XMLマッピングされたセクションは、連結先を同期せず、文字列項目を読み書きする際の拒否をRead・Replace経由でも確認しています。

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

`Replace` の増減は1件以上の基本ケースと、再取得・保存・表行構造・書式・識別IDを確認しています。0件の契約は引き続き検討メモとして保留しています。
コード生成は文字列明細の列挙を、生成元・実行時とも単数・複数で確認し、生成データによるReplaceも確認しています。明細内の4種類の項目は文書直下のプロパティ生成から除き、文字列Content Controlは通常項目と同名でも別々に読み取れます。明細データ型には文字列Content Control・MERGEFIELD・チェックボックス・日付選択を生成します。名前変換の読み書きと、Open時の各明細の文字列Content Control不足検査を確認しています。

## 精査で確認した残課題

以下は未解決事項の記録であり、新しい仕様や実装予定の承認ではありません。
コードから確認できる事実と、テストで未確認の挙動、ユーザーとの判断が必要な契約を分けて記載します。
対応時は既存テストとの重複を確認し、期待する外部挙動をレビューしてから必要なSkipテストまたはRedへ移します。
RS-02には別途テストを8件追加しました。うち5件をGreenにし、上記の予定50件には含めず、全体のSkipは既存の1件と合わせて4件です。

### RS-01：Replace失敗時に残る変更

* 確認済み：[RepeatingSection.Replace](../src/DocumentAsData/RepeatingSection.cs)は、入力を配列化した後、明細の件数を変更してから各明細へ順番に値を書き込みます。途中で失敗した際に、既に行った件数変更や値の変更を戻す処理はありません。今回のリファクタリング前からの処理順序です。
* 未確認：[既存テスト](../tests/DocumentAsData.Tests/RepeatingSectionのテスト.cs)にはXMLマッピングされたセクションのReplace拒否がありますが、例外後の件数・値・構造がどう残るかは検証していません。件数変更を伴う失敗や、途中の明細での失敗も対象外です。
* 要判断：失敗時に変更を残さない契約にするか、部分変更を許容して呼び出し側へ明示するか。ロールバックを既定の要件と決めつけず、期待する失敗後の状態を決めてからテストにします。

### RS-02：生成診断の検索範囲と対象名

* 変更前に確認した問題：[GenerateDiagnostics](../src/DocumentAsData.CodeGeneration/DocumentWrapperGenerator.cs)は4種類の項目名を文書全体でまとめて診断していたため、異なる明細の同名も衝突と判定していました。現在は文字列Content Controlだけ各セクションの最初の明細を診断し、MERGEFIELD・CheckBox・DatePickerは明細内を診断対象から外して、それぞれ別明細間の誤診断を解消しています。明細内の後者3種類の重複を引き続き検出できるかは未確認です。
* 確認済み：セクションから生成するプロパティ名と明細データ型名は診断対象に含まれず、予約メンバー一覧にも`RepeatingSections`は含まれていません。
* 確認済み：通常項目と明細の文字列Content Controlが同名でも、別の生成データ型に属するため衝突と診断しません。
* 未確認：新しく生成する名前の衝突を検出することは、まだ保証できません。
* テスト：[コード生成診断のテスト](../tests/DocumentAsData.CodeGeneration.Tests/コード生成診断のテスト.cs)に8件を追加しました。別明細の同名4種類と通常項目・明細の同名はGreen、セクション名と予約メンバー・通常項目の衝突、明細データ型名の衝突はSkipです。
* 次の確認：実際に生成される型・メンバーの範囲に対応した診断をレビューします。明細内Content Controlの重複を拒否する既存テストは維持し、明細内MERGEFIELD・CheckBox・DatePickerの重複診断は別途確認します。

### RS-03：文書直下の生成プロパティの検索範囲

* 確認済み：[生成プロパティ](../src/DocumentAsData.CodeGeneration/DocumentWrapperComponents.cs)で明細内を除外しているのは、文字列Content Controlのgetterの優先検索です。そのsetter、MERGEFIELD・CheckBox・DatePickerのgetter/setter、文字列項目の代替検索先は文書全体を対象にしています。
* 未確認：通常項目と明細内項目が同名の場合、既に確認した文字列Content Controlの優先読み取り以外で、意図した通常項目だけを読み書きできるか。単一項目だけの既存テストからは判断できません。
* 次の確認：生成プロパティの各検索経路を整理して必要なテストを選びます。低レイヤーのコレクションが文書全体を検索する既存契約は、これに合わせて変更しません。

### RS-04：生成Data経由の通常項目と明細内の同名

* 確認済み：生成Data型から明細内項目は除外していますが、生成Documentの`Read()`／`Replace(data)`は、文書全体を検索する[既存のオブジェクト対応付け](../src/DocumentAsData/DocumentObjectMapper.cs)へ委譲します。
* 未確認：通常項目と明細内項目が同名の場合の、生成Data経由の読み書きです。RS-03の個別プロパティとは別の経路です。
* 扱い：生成プロパティと汎用Read/Replaceの用途を区別する既存方針を、ここで変更しません。まず現行の制約と利用例を確認します。生成Dataの検索範囲を通常項目に限定する変更は未合意です。明細リストをDataへ追加する話とも区別します。

### RS-05：型付きOpenでセクション・通常項目の不足を独立に確認できていない

* 確認済み：[セクション不足の既存テスト](../tests/DocumentAsData.CodeGeneration.Tests/コード生成型構造のテスト.cs)では、セクションだけでなく内側の`ItemName`も不足します。期待する例外が出ても、セクション自体の不足を独立に確認したことにはなりません。
* 確認済み：[文書全体の必須項目検査](../src/DocumentAsData.CodeGeneration/DocumentWrapperComponents.Validation.cs)は、明細内を含む名前で照合し、通常項目の所属位置までは照合していません。
* 未確認：必要な項目名が別の場所に残っていてセクションだけがない場合と、通常項目だけがなくなり同名が明細内に残っている場合です。
* 次の確認：他の不足によって先に失敗しないデータで、それぞれOpen時に何を保証するかをレビューします。明細への検索切り替えなど、新たな互換性を暗黙には追加しません。

### RS-06：明細ごとの必須項目検査の範囲

* 確認済み：[明細別の検査生成](../src/DocumentAsData.CodeGeneration/DocumentWrapperComponents.Validation.cs)は、生成元の最初の明細にある文字列Content Controlを、各明細の文字列Content Controlと照合します。
* 未対応：各明細のMERGEFIELD・CheckBox・DatePickerの不足検査と、文字列Content ControlとMERGEFIELDを置き換えた場合の互換性です。文書全体の名前検査や、生成した値を読み取れるテストでは代替できません。
* 次の確認：種類ごとの不足と、文字列項目の互換性を別の論点としてレビューし、必要なテストを選びます。

### RS-07：OpenのOOXML検証バージョンとの整合

* 確認済み：[Document.Openの任意検証](../src/DocumentAsData/Document.cs)はOffice 2010を対象とします。一方、繰り返しセクションではOffice 2013の要素を扱い、固定テストデータもOffice 2013向けに検証しています。
* 未確認：繰り返しセクションを含む文書を`Open(..., validate: true)`で開く契約と結果です。通常のOpenや、SDKを使ったOffice 2013向け検証とは分けて確認する必要があります。
* 要判断：対応対象の検証バージョンをどう扱うか。検証対象を変えると他の文書にも影響するため、単なるリファクタリングとして変更しません。

### その他の未確認事項

* 複製時のContent Control識別IDは、[文書内の最大値に加算](../src/DocumentAsData/RepeatingSectionItem.cs)して割り当てています。既存テストは通常範囲のIDで1件追加した場合を確認しています。上限付近のIDや複数件を一度に追加する場合は、このテストだけでは保証できません。不具合を再現したという記録ではなく、追加確認候補です。
* 固定データと更新後の文書のWord上の表示・編集は未確認です。SDKの構造検証や再読み込みの成功と、Wordでの確認を区別します。
* 0件の契約、入れ子、任意構造の複製などは、既存の「未決定・今回の対象にしていないこと」を引き継ぎます。この精査で対応範囲を広げたものではありません。
