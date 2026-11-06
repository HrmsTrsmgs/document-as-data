# レビュー指摘と対応状況

## 現在の対応状況

2026-11-05確認。現在状態はこの表で更新します。再開位置と最新の検証結果は[current.md](current.md)、仕様と設計判断は[docs](../docs/design-decisions.md)を参照してください。
以下の観測・検証は各レビュー時点の記録であり、修正後の現状を表すものではありません。行番号も調査時点のものです。

| 指摘 | 状態・解決コミット |
| --- | --- |
| B01 | 破損XMLのOpen失敗後の解放を修正：`4323936`。Validator自体の例外経路は未確認 |
| B02 | Read：`52c0e8a`、Replace：`6187f8d` |
| B03 | placeholder=falseの読み取りを修正：`3840cdf` |
| B04 | 元Tagの対応付けを修正：`a6abc44`。コメント整理：`91a89df` |
| B05 | checked要素省略：`c336842`、val省略：`84632e3`。テスト表記整理：`bb17762`。setterの要素省略は未確認 |
| B06 | checkedStateのval省略を修正：`0055208`。unchecked側・要素全体省略は別の未確認事項 |
| B07 | 保留。ユーザーが2026-11-05にWord操作復旧まで修正しないと決定。実機確認を省いて先行する案は不採用。復旧後にWordで表示・保存を観測してからRedへ |
| B08 | 日付設定時のplaceholder解除：`f8dd2a0`。解除処理の共通化：`9a601d5` |
| B09 | `29a77e2` で修正・テスト整理をコミット済み。実装Refactor検討は追加変更なし |
| B10 | 明細内の生成名衝突は `d179403`、Refactorは `da59e99`。明細型名との衝突は `4c76f7d`、Refactorは `14a086b`。外側の明細型名とセクション名の衝突は `f68548e` で修正済み（コミットメッセージは誤り）。他ケースは未着手 |
| B11 | CR入りCheckBox Tagは `4195d95` でGreenをコミット済み。Refactor検討は追加変更なしで完了。他ケースは未着手 |
| B12 | 未着手。下記の再現記録を参照し、一件ずつ契約を確認 |
| V01〜V05、R01〜R03、P01〜P02 | 調査・改善候補。記載を実装許可とみなさない |
| D01 | 記録の配置・現在状態と履歴の整理を実施中。公開文書・既存コメントとの整合確認は未完了 |
| D02 | 記録先と主担当／工程担当の入口はAGENTSで指定。下記のその他の指示差異は今回再検証していない |

## 過去の調査と判断の根拠

完了した整理の理由・不採用案と、未着手の再現条件を保持しています。過去の「次のRed」「未対応」は当時の提案です。作業順序や完了状態には上表を使用してください。

# 次版準備のリファクタリング精査

最新の全体レビューは、末尾の[2026-10-23の精査結果](#全体レビュー2026-10-23)を参照してください。以下の2026-09-27の記述は当時の作業記録です。

2026-09-27、RS-03完了（`b6e43cd`）を基準に実施しました。
機能・公開API・生成コードの契約は変更せず、内部の責務と重複を整理する作業です。
バージョン変更や、未合意の残課題の実装は含みません。

## 確認した範囲

* Core：文書とStreamの生存期間、保存、項目・コレクション、OOXMLのフィールド解析と値更新、オブジェクト対応付け、繰り返しセクション。
* CodeGeneration：生成の入口、型・プロパティのテンプレート、必須項目検査、名前変換と衝突診断。
* Build：生成タスク、辞書の読み込み、props/targets、パッケージ構成。
* テスト：変更対象の既存仕様テスト、生成コードのコンパイル・観測補助、MSBuild連携の補助。テストの共通化によって仕様の記述を隠さないかも確認しました。

## 実施した整理

| 対象 | 整理した責務・重複 | 維持する振る舞い |
| --- | --- | --- |
| [DatePicker.cs](../src/DocumentAsData/DatePicker.cs)、[DatePicker.DisplayFormat.cs](../src/DocumentAsData/DatePicker.DisplayFormat.cs) | Word書式の字句解析と.NET書式への変換を、専用の内部クラス`DisplayFormat`へ移動 | 正規表現、変換表、表示設定の解決順序、日時を書き換えるタイミング |
| [TextContentControl.cs](../src/DocumentAsData/TextContentControl.cs) | 後続段落を削除できる条件を`ContainsOnlyTextAndFormatting`として抽出 | 自身の直下だけを対象とし、画像や入れ子を残す条件、削除前の配列化 |
| [DocumentWrapperGenerator.cs](../src/DocumentAsData.CodeGeneration/DocumentWrapperGenerator.cs)、[GeneratedNameDiagnostics.cs](../src/DocumentAsData.CodeGeneration/GeneratedNameDiagnostics.cs) | 文書のOpenと設定適用を入口に残し、診断を`GeneratedNameDiagnostics`へ分離 | 診断の対象種類・範囲、元名の順序、予約名、結果の順序 |
| [DocumentItemScopes.cs](../src/DocumentAsData.CodeGeneration/DocumentItemScopes.cs) | 通常項目の選別を、プロパティ生成と名前診断から共有 | 明細内の項目を除く既存条件。文書全体の必須項目検査には適用しない |
| [DocumentWrapperComponents.cs](../src/DocumentAsData.CodeGeneration/DocumentWrapperComponents.cs) | 一つのアクセス式からgetterとsetterを生成 | 生成メンバーは増やさず、検索・代替先・例外・毎回の評価を維持 |

日付書式はDatePickerだけに属するため、汎用の公開書式変換APIにはしません。
名前診断の分離はSpreadsheetAsDataの`GeneratedNameDiagnostics`も参照しましたが、Excel固有の診断や未対応のWord項目を取り込む変更はしていません。
テストの削除・統合・Skipの変更はありません。[コード生成名前付き項目のテスト.cs](../tests/DocumentAsData.CodeGeneration.Tests/コード生成名前付き項目のテスト.cs)では、RS-03で有効化済みのgetterを「未対応」としていたコメントだけを訂正しました。

## 統合・削除しなかったもの

* `DocumentSession`と`CopyOnWriteStream`：ファイルの束縛・保存と、Streamのコピー切り替えは異なる責務です。既存の所有権を維持しました。
* MERGEFIELDの単純形式・複合形式：値の公開APIは同じでも、境界や削除範囲が異なります。既存の内部クラスによる分離を維持しました。
* 汎用Read/Replaceと生成プロパティ：前者の一意な対応付けと、後者の元種類優先・通常項目の検索を統合しません。
* オブジェクト対応付けの読み取りと書き込み：対象プロパティの選別、事前検査、例外の契約が異なります。共通の分岐表へ押し込む変更はしません。
* コレクション：Content Controlの共通部分は既存のジェネリック基底クラスで共有済みです。MERGEFIELDの状態を持つ解析まで同じ基底クラスへ寄せません。
* 一時変数・配列：複数回使う値、削除前のOOXML要素列、明細の件数変更前の一覧、設定検査の評価順序を固定するものは残しました。短縮のために再評価や副作用の順序を変えません。
* MSBuildのテンプレートとテスト補助：実行方法や生成対象の違いを読み取れる既存の区分を維持しました。大量の文字列を一つの汎用ビルダーへ隠す変更はしません。

## 振る舞いの判断が必要な残課題

[繰り返しセクションの残課題](repeating-sections-plan.md#精査で確認した残課題)を引き続き管理元とします。
RS-03は完了済みで、RS-01の失敗時の部分変更、RS-02の残る診断範囲、RS-04〜07、0件の契約は今回変更していません。

今回のコード確認で気付いた次の点も、リファクタリングへ修正を混ぜていません。

* **生成XMLコメントのエスケープに差がある**：文字列Content Controlとセクション名にはアンパサンド・CR・LFの処理がありますが、MERGEFIELD・CheckBox・DatePickerのコメントはLFだけを処理しています。文字列リテラルの処理とは別です。コード上の差は確認済みですが、各種類の入力で問題となる診断や表示は未検証です。期待するコメントと固定データを確認し、必要なら別のRedへ移します。
* **XMLマッピング拒否の経路に差がある**：Content Controlは自身と親の繰り返しセクションを確認しますが、MERGEFIELDの読み書きには同じ検査がありません。既存のセクション拒否テストは文字列Content Controlを通るため、MERGEFIELDだけの明細は保証していません。拒否対象とするかも含め、実装ではなく次の仕様レビュー事項です。

## 再精査（同日）

Core、CodeGeneration、Build・テスト補助を分担して読み直し、親レビューで差分と指摘を再確認しました。
今回のリファクタリングによる挙動差は見つかりませんでしたが、次の既存不具合を再現しました。

### 複合MERGEFIELDを別の検索範囲から取得した場合の不整合

0.4.0の公開準備で対応しました。固定DOCX「明細内に複合MERGEFIELDを持つ繰り返しセクション.docx」と、
[MergeFieldの回帰テスト](../tests/DocumentAsData.Tests/MergeFieldのテスト.cs)を追加してRedを確認後、同じOOXML要素のMERGEFIELDを文書単位で共有してGreenにしました。
Release全体テストは595件成功、Skip 1件です。以下は修正前の精査記録です。

* `MergeFieldCollection.MergeFieldReader`のキャッシュはコレクション単位です。文書全体と明細内のコレクションは、同じOOXML要素に対して別の`MergeField`を作ります。
* 複合形式の`ComplexFieldContent`は表示要素の一覧を保持します。一方のオブジェクトから書き換えると、他方の一覧には削除済みの要素が残ります。
* 公開APIで再現確認済み：文書全体と明細内から同じ複合フィールドを先に取得し、明細側から`Text`を更新すると、文書全体側は古い文字列を返しました。続けて文書全体側から書き込むと、`InvalidOperationException`（`The parent of this element is null.`）になりました。
* 再現入力は、固定の「明細に4種類の項目.docx」の最初の単純フィールドを、固定の「複合形式のMERGEFIELD.docx」の複合フィールドへメモリ上だけで置き換えました。元ファイル・保存済みテストデータは変更していません。
* 現行の固定DOCXには繰り返しセクションと複合フィールドを併せ持つものがなく、明細のMERGEFIELDテストは単純形式を使用しています。この再現確認は既存テストの成功とは別で、まだ回帰テストにはしていません。
* 今回の差分に起因する問題ではなく、RS-03の通常項目の検索範囲とも別です。修正前に固定データと公開APIのRedを追加し、別経路でも同じ項目の現在値を扱う契約を確認します。状態共有の範囲や、表示要素を再取得する設計はその後に検討します。

### 追加の小さな整理候補

* `MSBuild連携テストプロジェクト.AddPowerShellPackageReferenceSample`は、パッケージ情報の取得、復元設定、利用プロジェクト、共用ライブラリ利用アプリ、実行用コード、PowerShellの準備を一つのメソッドに含みます。汎用ビルダーにはせず、復元設定や利用アプリの作成を個別のメソッドへ分けると、呼び出し元から準備の流れを追いやすくなります。テスト本文の仕様やサンプル種別は隠しません。
* `コード生成名前付き項目のテスト.cs`のCRを含むTagのテストには、Red時点の「現状は…生の改行が入ります」というコメントが残っています。現在はエスケープ済みなので、確認している仕様の説明へ訂正する候補です。

再精査ではコード・テストを変更せず、この記録のみ更新しました。新たな不具合の修正をリファクタリングへ混ぜていません。ビルド・全体テストは以下の前回結果を参照し、同じコードに対しては再実行していません。

## 検証

* 変更前の全体テスト：594件成功、失敗0件、Skip 1件。制限付き実行ではNuGet設定へのアクセス拒否でMSBuild連携16件が失敗したため、必要なアクセス権で再実行して確認しました。
* 段階別：日付選択関連68件、文字列Content Control30件、名前診断42件、MSBuild連携以外のコード生成244件が成功しました。
* アクセサー部品の共通化前後で、固定DOCX 146件の生成結果を比較しました。生成できた144件のC#ソースはSHA-256が一致し、不正入力2件の例外型も一致しました。これは既存入力での同等性確認であり、未検証入力の正しさの証明ではありません。
* 最終の全体テスト：`dotnet test DocumentAsData.slnx --no-build --no-restore`で594件成功、失敗0件、Skip 1件（Core 309件、コード生成・MSBuild連携285件）。残るSkipは明細0件の契約メモです。
* 最終の診断ビルド：`--no-incremental -p:UseSharedCompilation=false -p:EnforceCodeStyleInBuild=true -p:RunAnalyzersDuringBuild=true -p:ErrorLog=obj/ide-diagnostics.sarif`を付けて実行。約14秒、警告・エラー0件でした。全プロジェクトのSARIFを確認し、変更箇所への新しい指摘はありません。
* `dotnet format DocumentAsData.slnx --no-restore --verify-no-changes --severity info`は既存のCA1816（Document.Dispose）とxUnit1004（仕様未決定のSkip）により終了コード1です。成功扱いにはしていません。SARIFには、これらに加えて正規表現の生成ソースと意図的なテスト用インスタンスプロパティに対する抑制済み診断が残ります。
* 改行・BOMの正規化と`git diff --check`も実施しました。バージョン変更・push・公開は行っていません。
* Wordでの画面・編集確認は実施していません。今回はDOCXの処理規則を変えていません。

## 全体レビュー（2026-10-23）

### 対象と結論

基準は `0d56672` と、開始時から存在するコード生成型構造テストのSkip解除1件です。
今回は改善点の一覧化だけを行い、製品コード・既存テスト・仕様・開発規則は変更していません。
下記の候補をすべて実装する決定や、次版の必須要件を追加する決定ではありません。

* 製品C#全36ファイル・4,085行、テスト／テスト補助C#全28ファイル・12,055行を分担して全行確認しました。関連するプロジェクト、props/targets、スクリプト、CI、README、チュートリアル、仕様・計画・AI指示、直近のGit履歴も照合しました。
* 観点は、公開契約、境界値・例外、副作用・所有権、OOXMLの意味、コード生成と診断、配布・ビルド、テストの観測範囲、責務・重複・読みやすさ、性能、文書と開発再開の整合です。
* 全テストは613件成功、失敗0件、Skip 5件でした。それとは別に、既存テストにない入力で読み取り値・生成型・リソース解放などの問題を再現しました。
* 大規模な継承再編や公開抽象化を増やす根拠は見つかりませんでした。まず振る舞いの不足をRedで扱い、その後に小さな共通化を行う順序が適切です。

優先度は、P1＝次の修正で優先したい誤った結果・解放漏れ、P2＝対応範囲内の境界や品質保証、P3＝整理・計測候補です。
「再現済み」は記載した入力・経路で確認したという意味で、同系統の全パターンやWord上の描画まで保証するものではありません。

### 振る舞いの改善候補

#### B01 / P1：Open失敗直後に元ファイルの束縛が残る — 新規・再現済み

* 対象：[Document.cs](../src/DocumentAsData/Document.cs) の57–72行、312行付近。
* ZIPとして成立するDOCXの `word/document.xml` だけを途中で切断して開くと、`XmlException` の直後に排他Openが `IOException` になります。Documentが返らないため、呼び出し側からDisposeできません。
* `DocumentSession.Open` の完了後、Documentの初期化でXMLを遅延読み込みする際の例外に解放経路がありません。プロセス終了後には排他Openできました。永久リークやGC後も残るとは断定しません。
* 次のRed：破損XMLによるOpen失敗直後にも排他Openできること。既存の空ファイル・検証エラー時の解放テストを弱めず、例外の発生段階を追加します。
* 関連する未確認箇所：131行付近のValidator自体が例外を投げる経路。同じ解放不足がコード上にありますが、こちらの実行再現は完了していません。

#### B02 / P1：明細型が文書全体のData型を隠す — 既知RS-02の実害を新たに再現

* 対象：[DocumentWrapperComponents.cs](../src/DocumentAsData.CodeGeneration/DocumentWrapperComponents.cs) の103、139、153行付近。
* `repeatingTemplate.docx` で `NameMappings["Items"] = "RepeatingTemplate"` とすると、診断なしでコンパイルできます。しかし `Read()` の戻り型と `Replace()` の引数型が、外側の `Generated.RepeatingTemplateData` ではなく、文書型内の同名明細型になります。実際のRead結果も明細の `ItemName` を持ちます。
* 次のRed：この名前配置でも、生成文書のReadが文書全体のData型を返すこと。名前衝突を拒否する契約にする場合は、その判断を先に確認します。コンパイル成功だけでは検証できない項目です。

#### B03 / P1：placeholder=falseの内容を空として読む — 新規・再現済み

* 対象：[TextContentControl.cs](../src/DocumentAsData/TextContentControl.cs) の169–170行。
* 値のある文字列Content Controlへ `<w:showingPlcHdr w:val="false"/>` を付けると、Validatorエラー0件でも `Text` は空文字です。現在は要素の存在だけを見ています。
* [ShowingPlaceholderの規則](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.showingplaceholder?view=openxml-3.0.1)では、trueの場合だけplaceholderとして扱います。
* 次のRed：false指定では元の文字列を取得すること。要素省略・trueの既存契約は維持します。

#### B04 / P2：`@`付きの元Tagと生成名が同じ場合に対応付け属性が欠ける — 新規・再現済み

* 対象：[DocumentWrapperComponents.cs](../src/DocumentAsData.CodeGeneration/DocumentWrapperComponents.cs) の316–320行。
* Tagを `@CustomerName` とし、辞書でも同じ生成名を指定すると、直接プロパティは値を読めますが、生成Data経由のReadは `InvalidOperationException` になります。診断・コンパイルは成功します。
* ソース上の名前が等しいため属性を省略しますが、C#のメタデータ上では `@` が除かれます。既存のUnicode書式文字への対処では捕捉できません。
* 次のRed：この名前設定で生成Dataへ値を読み込めること。Replaceも同じ対応付けを使うことはコード確認済みですが、今回の実行はReadです。

#### B05 / P2：CheckBoxの合法な状態省略を読めない — 新規・再現済み

* 対象：[CheckBox.cs](../src/DocumentAsData/CheckBox.cs) の44、65–66、101–103行。
* `w14:checked` 要素なしではgetterが例外、`<w14:checked/>` のval省略ではfalseになります。どちらの入力もOffice 2010のValidatorエラーは0件です。
* [CT_SdtCheckbox](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/d4682d85-ae4d-48a7-8376-ec2448ae7e35)では要素省略がfalse、[CT_OnOff](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/173a98b1-92a5-4dd8-a0be-0614b1cb6606)では要素あり・val省略がtrueです。
* 次のRed：この二つの入力を別々に確認します。要素省略時のsetterは追加の確認対象であり、今回実行済みではありません。

#### B06 / P2：CheckBoxの表示記号省略で状態だけ変更して失敗する — 新規・再現済み


* 対象：[CheckBox.cs](../src/DocumentAsData/CheckBox.cs) の49–53、79–94行。
* 未チェックの入力で `checkedState` のvalを省略しtrue設定すると、例外を投げますが、チェック状態はtrue、保存表示は `☐` のままです。入力のValidatorエラーは0件でした。
* [CheckedStateの規則](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.word.checkedstate?view=openxml-3.0.1)には省略時の記号U+2612が定義されています。
* 次のRed：この合法な省略入力への設定が成功し、状態と表示がtrue／`☒`になること。unchecked側と要素全体の省略は別の未確認観点です。任意の例外に対する完全なロールバックをここから要求するものではありません。

#### B07 / P2：CheckBoxの表示記号用フォントを反映しない — XMLで再現、Word表示は未確認


* 対象：[CheckBox.cs](../src/DocumentAsData/CheckBox.cs) の52、79–83行。
* checkedStateを `font="Wingdings", val="00FC"`、元の表示runをMS Gothicとしてtrue設定すると、保存文字はU+00FCの `ü`、runのフォントはMS Gothicのままです。入力のValidatorエラーは0件です。
* 表示記号はコードポイントとフォントの組で指定されます。[CheckedStateのfont属性](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.word.checkedstate?view=openxml-3.0.1)
* 次の確認：Wordが保存値をどう補正・表示するかを確認し、異なる記号フォントを持つ入力の切替結果をRedへします。文字だけの更新と、保存表示の整合を区別します。

#### B08 / P2：日付設定後にもplaceholder指定が残る — 未保証範囲を具体化


* 対象：[DatePicker.cs](../src/DocumentAsData/DatePicker.cs) の62–65行。
* 未入力の日付選択へ `w:showingPlcHdr` を付けて日時設定すると、fullDateと表示文字は更新されますが、placeholder指定はtrueのままです。入力のValidatorエラーは0件でした。
* 既存の未入力テストはコメントでplaceholderを対象外としています。既存テストの期待値違反とせず、公開説明の「未入力でも設定できる」に関連する保証の不足として扱います。
* 次の確認：対応範囲を確認後、日時・表示・placeholder解除を保存XMLで観測するRed。Word上の表示・編集結果はまだ確認していません。[placeholder指定の意味](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.showingplaceholder?view=openxml-3.0.1)

#### B09 / P2：明細複製時にIDがオーバーフローし重複する — 既知の上限候補を再現済みに更新


* 対象：[RepeatingSectionItem.cs](../src/DocumentAsData/RepeatingSectionItem.cs) の74–78行。
* 既存IDに `2147483647` と `-2147483648` がある2明細を3明細へ増やすと、加算で負数へ回り、既存の `-2147483648` と重複しました。入力・出力ともOffice 2013のValidatorエラーは0件です。
* 次のRed：この入力でも保存後のIDが一意であること。既存のID一意性契約の境界値であり、任意構造の複製対応を広げる話ではありません。

#### B10 / P2：生成先スコープと名前診断の範囲が一致しない — 一部既知RS-02・再現済み

対象：[GeneratedNameDiagnostics.cs](../src/DocumentAsData.CodeGeneration/GeneratedNameDiagnostics.cs) の45–63、87–89、103–109行。

| 入力 | 診断API | 生成コードの実結果 |
| --- | --- | --- |
| `repeatingKinds.docx`、`Code → ItemName` | エラーなし | 明細のMERGEFIELDと文字列項目が重複しCS0102 |
| `repeatingTemplate.docx`、`ItemName → ItemsData` | エラーなし | 明細型自身とプロパティが同名になりCS0542 |
| `repeatingSections.docx`、`Options → ItemsData` | エラーなし | 外側Documentで明細型とセクションプロパティが衝突しCS0102 |
| `repeatingTemplate.docx`、`ItemName → Read` | 予約名エラー | 明細DataのReadプロパティは警告なしでコンパイル可能 |

明細の診断は文字列Content Controlだけを対象としますが、生成は4種類を扱います。また、文書型の予約名を明細型にも一律に適用しています。
次のRedは、生成先・衝突の種類ごとに分けます。明細のReadもポリシーとして一律予約したいのかは確認し、C#上の衝突があることと混同しません。

#### B11 / P2：項目名の改行で生成C#が壊れる — 一部既知・再現済み

* 対象：[DocumentWrapperComponents.cs](../src/DocumentAsData.CodeGeneration/DocumentWrapperComponents.cs) の282、303、334、362、384、447–463行付近。
* CheckBoxのTag `Agree\rment` はコメントの行を分断し、CS1519等になります。文字列Content ControlのTagにU+0085・U+2028・U+2029を含めた3ケースも、文字列リテラル等が分断されCS1010等になります。診断APIはいずれもエラーなしでした。
* 次のRed：CRのあるCheckBox名と、Unicode改行のある文字列項目名を区別して生成・読み取りを確認します。CR差は以前の未検証指摘、Unicode改行は今回の追加確認です。

#### B12 / P3：生成XMLコメントのエスケープと検証が不十分 — 一部既知・再現済み

* 対象：[DocumentWrapperComponents.cs](../src/DocumentAsData.CodeGeneration/DocumentWrapperComponents.cs) の282、406、468行付近、[GeneratedSourceCompiler.cs](../tests/DocumentAsData.CodeGeneration.Tests/テスト補助/GeneratedSourceCompiler.cs) の20–40行。
* CheckBoxの `Agree&ment`、文字列Content Controlの `Customer<Name` は、XMLコメント診断を有効にするとCS1570になります。通常のコンパイルは成功するため、警告を見ない生成テストでは見逃します。
* 次のRed：既存のコメントテストの方法で診断を確認します。C#リテラルとXMLコメントは異なる出力形式です。必要なケースを固定した後、それぞれの用途内で処理を共通化します。

### ビルド・検証の改善候補

| ID・優先度 | 確認した内容・根拠 | 次の作業 |
| --- | --- | --- |
| V01 / P2 / 再現済み | [Normalize-ChangedTextFiles.ps1](../scripts/Normalize-ChangedTextFiles.ps1) の14行付近。存在しないGIT_DIRを子プロセスへ指定すると、Gitは失敗しても `-Check -IncludeUntracked` が終了0になる | Git失敗時は正規化検査も非0になるRed。対象0件の正常状態と取得失敗を区別 |
| V02 / P2 / 静的確認 | [Prepare-Release.ps1](../scripts/Prepare-Release.ps1) の49行。VS追加検証の「パッケージ」というテスト名フィルターが、共用ライブラリ配布テスト2件を選ばない。通常の全テストでは実行される | 役割を表すTrait等で選択し、追加検証の対象に2件が含まれることを確認 |
| V03 / P2 / 構成確認・差異の影響は未検証 | [MSBuild連携テストプロジェクト.cs](../tests/DocumentAsData.CodeGeneration.Tests/テスト補助/MSBuild連携テストプロジェクト.cs) の30–33、634–636行。OS一時領域を子プロセスの作業場所にし、global.jsonを引き継がない | 子プロセスのSDKも記録し、どのSDKを試験対象にするか固定。今回の指定SDK不在と、テスト用子プロセスの選択問題は別件 |
| V04 / P3 / 構成確認 | Releaseスクリプト35–43行と[CI](../.github/workflows/verify.yml)は増分ビルド・formatのseverity warn。開発規則の非増分解析・SARIF確認・severity infoとは保証範囲が異なる | 既知の診断をどう受容するか明示して検査経路を整理。診断をゼロにするためだけの実装変更はしない |
| V05 / P3 / 静的確認 | MSBuild連携の補助648行付近は子プロセスを無期限に待つ。途中停止時のタイムアウト・途中出力回収がない | テスト実行環境の責務として期限を決め、停止する子プロセスでRed。今回の接続待ち失敗の原因と断定しない |

追加調査に留める点：別ディレクトリからのMSBuildで相対DOCXパスがどう解決されるか、および異なるディレクトリの同名DOCXを同じ名前空間で同時生成・コンパイルした場合です。
既存の同名ファイルテストは生成先パスの区別までを確認しています。今回、実MSBuildでの再現までは行っていないため、確定不具合には数えません。

### 文書・運用の改善候補

#### D01 / P2：現在状態と履歴を分けて同期する — 確認済み

* [繰り返し計画](repeating-sections-plan.md) の107–122、202–212行付近には、RS-05/06の実装済み項目を未対応／Skipとして扱う記述があります。現在の作業ツリーでは該当テストは有効で、今回成功しました。残るSkipはRS-07の4件と0件契約メモ1件です。
* [機能リファレンス](../docs/reference.md) の420行には0.4.0時点の制約が残ります。公開済み0.4.0の説明と開発中の対応状況を明示的に分け、過去のCHANGELOGを現状へ書き換えないようにします。
* `DocumentItemScopes` のクラス説明、Validationの76・111行付近、コード生成名前付き項目テストの653行付近にも、現行コードと異なる説明があります。
* 開発コンテキストは、コミット済みのAI-DLC作業を未コミットとしていました。今回のレビュー終了状態だけは本作業で更新し、それ以外の文書・コメント修正は候補として残します。

#### D02 / P2：AI指示の参照先・担当範囲を整理する — 確認済みの記述差

* [run.md](../docs/ai-instructions/tdd-context/run.md) は候補を `docs/development-candidates.md` に記録するよう指定しますが、そのファイルはありません。AI-DLCでは既存の開発コンテキストと計画文書を使う方針です。新しい記録先を増やす前に正本を統一します。
* AGENTSの再開用記録の読み込み指示と、Greenへ将来候補・却下案を渡さない規則の担当範囲を明確にします。進行役が全体記録を読み、担当には必要な確定契約だけ渡す整理が候補です。
* C#指示113行が全テスト省略条件をTDD文書へ委ねますが、`tdd.md` は索引であり、参照先には以前の「論理的に影響がない場合だけ省略し理由を報告」という具体条件が揃っていません。全テスト原則そのものは残っています。
* Skillの「各工程に新しい担当」とrun.mdの「Red/Refactorは同領域なら継続可」の表現、工程文書内の重複文も揃えられます。TDDのRed/Green/Refactorの役割を変更する提案ではありません。
* [テスト記述方針](../docs/testing-style.md)はWorkbook・Table等の移植元の例が中心です。共有ルールとして残す箇所と、DocumentAsDataの実在APIで示す代表例を区別すると、存在しないAPIを実装の前提にする誤りを減らせます。

### 振る舞いを変えない整理候補

| ID・優先度 | 対象 | 整理の方向・維持すること |
| --- | --- | --- |
| R01 / P3 | [GenerateDocumentAsData.cs](../src/DocumentAsData.Build/GenerateDocumentAsData.cs) の77–98行 | 診断と生成で同じDOCXを別々に開く。公開APIを増やさず、1回の解析を内部で共有できる。例外・診断順序・所有権を維持する。速度改善量は未計測 |
| R02 / P3 | [Validation](../src/DocumentAsData.CodeGeneration/DocumentWrapperComponents.Validation.cs) の130–144行 | `section.Items[0]` を4回取得し、その都度明細一覧を作る。生成宣言側と同様に最初の明細を一度受け取る。複数回使う意味のある一時変数であり、無条件な変数削除とは逆の改善 |
| R03 / P3 / 既知 | [MSBuild連携テスト補助](../tests/DocumentAsData.CodeGeneration.Tests/テスト補助/MSBuild連携テストプロジェクト.cs) の97行からのサンプル準備 | 復元設定・利用プロジェクト・アプリ・実行準備を意味のあるメソッドへ分ける。テスト本文の意図を隠す汎用ビルダーにはしない |

### 計測してから判断する候補

* **P01 / P3：削除された要素のキャッシュ寿命。** [OpenXmlElementCache.cs](../src/DocumentAsData/OpenXmlElementCache.cs) は強参照Dictionaryです。明細の増減と取得を長時間繰り返す場合、削除した要素とwrapperがDocumentの生存中に残る構造です。保持量は未計測。同じ要素のwrapper共有という修正済みの契約を崩さず検討します。
* **P02 / P3：一操作内の繰り返し走査。** [DocumentObjectMapper.cs](../src/DocumentAsData/DocumentObjectMapper.cs) はプロパティごとに全種類の名前列挙・識別子化を行い、明細追加は1件ごとに文書全体の最大IDを走査します。多数項目・多数明細で計測し、必要なら操作内だけ再利用します。可変文書へ長寿命キャッシュを追加することや、性能を理由にテストを減らすことは先に決めません。

### 新たな不具合・追加要件に数えないもの

* 汎用Read/Replaceと生成プロパティの検索・同名解決の違いは、利用場面の違いを理由に確認済みの契約です。
* RS-01の途中失敗時の状態、RS-07のOffice 2013検証、明細0件、入れ子・任意構造の複製は、既存の計画・Skipで管理します。今回の発見として重複計上しません。
* MERGEFIELDの単純／複合形式、DocumentSessionとCopyOnWriteStream、読み取り／書き込みの対応付けには異なる変更理由があります。見かけの重複だけで統合しません。
* DTOのindexer・非public setter、任意のReplace失敗の原子性は、期待する契約を今回確定していません。挙動を見ただけで新仕様にしません。
* 大規模な継承再編、ODT、レンダリング、全書式対応など、ロードマップを先へ広げる提案は含めません。

### 今回の検証結果と限界

* **成功：診断ビルド。** 指定SDK 10.0.302はこのPCになく、標準のSDK選択は失敗しました。global.jsonは変更せず、インストール済み10.0.401のMSBuild.dllを直接指定し、Rebuild・共有コンパイラー無効・コードスタイル解析・SARIF出力を有効にして全プロジェクトを検証しました。指定SDKでの再現確認ではありません。
* **成功：全体テスト。** 同じSDKでMSBuildのVSTestを直列実行。CodeGeneration／MSBuild 303件、Core 310件の計613件成功、失敗0件。CoreのSkipは5件です。制限環境内の初回はtesthost接続タイムアウトで中止し、必要な権限で再実行して成功しました。初回を成功扱いにはしていません。
* **確認済み：SARIF。** 未抑制の情報診断はCA1816（Document.Dispose）とxUnit1004（5件のSkip）。生成された正規表現のCS0162/CS0164、意図的なテスト用メンバーのCA1822にはソース内抑制があります。隠れた警告を無視して「診断が全てゼロ」とはしていません。
* **実行できず：format検査。** `--verify-no-changes --severity info` は、制限環境内ではパイプ接続拒否、権限を変えた再試行では指定SDK不足で解析開始前に失敗しました。書式検査成功とは報告しません。SDK導入・設定変更は本レビューでは行いません。
* **成功：独立した再現確認。** Coreの8ケースとコード生成の13ケースで上記結果を確認しました。Coreのstyles破損の補助ケース1件は入力準備に失敗したため、証拠に使っていません。入力は既存固定DOCXの一時コピーだけを変更しました。これらは調査用プログラムであり、正式な回帰テストを追加したことにはなりません。
* **成功：差分・文書参照。** `git diff --check` とMarkdownの相対リンク先ファイルの存在確認を実施しました。開始時の未コミットテスト変更は保全しています。
* **確認済み：サンプルの静的整合。** READMEとチュートリアルの全19コードブロックをAPIと照合し、参照する3固定DOCXの項目名・初期値も一致しました。サンプルそのものの実コンパイル・実行、公開NuGetとURLの現在状態は未確認です。
* **未実施：** 全158固定DOCXの全XML要素・Word表示を個別に確認する作業、Word／LibreOfficeの描画・編集確認、性能・メモリ測定、OS横断・Blazor実機検証、外部の脆弱性データベースによる依存監査、GitHub上のCI再実行。ソース全行確認は、全入力や全実行環境を網羅したという意味ではありません。

本レビューで追加した追跡対象の変更は、本記録と開発コンテキストだけです。開始時からのテスト差分は別に残っています。コミット・push・バージョン変更・公開は行っていません。
