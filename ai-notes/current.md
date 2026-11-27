# 現在の作業状態

確認日：2026-11-28。

## 停止位置

* `codex/dry-refactoring`、基準 `468591e`。生成対象の列挙を共通化した最初の変更はコミット済み。DTOマッピングの型と項目の対応について、boolの代表例を確認し、続いてほかの型への展開を検討する。
* [GitHub Project](https://github.com/users/HrmsTrsmgs/projects/1)の「DocumentAsDataのコード生成で対象項目の列挙を共通化する」は進行中。コミット `468591e` はmainへの取り込み待ち。DTOマッピングの対応をProjectに登録し進行中。生成XMLコメントの警告と生成名衝突の扱いは別作業。リリースはSpreadsheetAsDataとの足並みを確認するまで保留。

## 検証

* .NET 11で全6プロジェクトの非増分ビルドに成功。警告・エラー0。全644テスト成功（本体323、コード生成・MSBuild連携321）、失敗・Skip 0。`format --verify-no-changes --severity info` に成功。変更プロジェクトのSARIF診断0件。
* 固定DOCX 172件について基準版と変更版の生成ソース、診断、受け付けなかった入力の例外型と文言が一致した。

## 別作業

* **保存時の空白保持**：単純MERGEFIELD・複合MERGEFIELD・文字列Content Controlで、前後に空白がある文字列を保存しても `xml:space="preserve"` が付かない。既存テストはnull条件演算子がアサーションまで短絡していた。ユーザー判断で切り分け、同じ基準の `codex/xml-space-preservation` にアサーション3か所の未コミット差分を残した。3件とも属性がnullのため失敗することを再現済み。再開時はこのRedから。
* **改行正規化スクリプト**：サンプル入力 `abc\n \n` に対し1回目は `abc\r\n\r\n`、2回目は `abc\r\n`。末尾の空白行の扱いは別作業で検討する。
* B07のフォント修正はWord実機確認まで保留。これまでの指摘と性能候補は[レビュー記録](refactoring-review.md)、完了した繰り返し対応は[計画](repeating-sections-plan.md)を参照。
