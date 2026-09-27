# 次版準備のリファクタリング精査

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
