# 現在の作業状態

確認日：2026-11-28。

## 停止位置

* `codex/dry-refactoring`、基準 `b959839`。コード生成テストの同じ文書準備27か所への展開はコミット済み。固定DOCXの同内容4組のうち、文字列Content Controlの代表1組の共有は確認済み。残る3組への展開は未了承。空文書4件は用途と名前が異なるため独立を維持する候補。
* [GitHub Project](https://github.com/users/HrmsTrsmgs/projects/1)の「DocumentAsDataのコード生成で対象項目の列挙を共通化する」は進行中。コミット `468591e` はmainへの取り込み待ち。DTOマッピング、生成メソッド名、既定名前空間の対応はProjectで進行中。名前解決の共通化は複雑さに見合わず保留。MSBuild接尾辞とテスト一時ファイルの共有はAPI・構成の変更が必要なため未着手。生成XMLコメントの警告と生成名衝突の扱いは別作業。リリースはSpreadsheetAsDataとの足並みを確認するまで保留。

## 検証

* .NET 11で全6プロジェクトの非増分ビルドに成功。警告・エラー0。全644テスト成功（本体323、コード生成・MSBuild連携321）、失敗・Skip 0。`format --verify-no-changes --severity info` に成功。変更プロジェクトのSARIF診断0件。
* 固定DOCX 172件について基準版と変更版の生成ソース、診断、受け付けなかった入力の例外型と文言が一致した。

## 別作業

* **保存時の空白保持**：単純MERGEFIELD・複合MERGEFIELD・文字列Content Controlで、前後に空白がある文字列を保存しても `xml:space="preserve"` が付かない。既存テストはnull条件演算子がアサーションまで短絡していた。ユーザー判断で切り分け、同じ基準の `codex/xml-space-preservation` にアサーション3か所の未コミット差分を残した。3件とも属性がnullのため失敗することを再現済み。再開時はこのRedから。
* **改行正規化スクリプト**：サンプル入力 `abc\n \n` に対し1回目は `abc\r\n\r\n`、2回目は `abc\r\n`。末尾の空白行の扱いは別作業で検討する。
* B07のフォント修正はWord実機確認まで保留。これまでの指摘と性能候補は[レビュー記録](refactoring-review.md)、完了した繰り返し対応は[計画](repeating-sections-plan.md)を参照。
