# 現在の作業状態

確認日：2026-11-25。

## Microsoft.Testing.Platform移行（レビュー待ち）

* `codex/microsoft-testing-platform`、基準 `25ef063`。テスト基盤・リリース検証の5ファイルに未コミット差分がある。依存変更はPRまで進める指示で了承済み。コミット前の差分レビューと了承待ち。
* Release検証で全644件、Visual Studio MSBuild追加13件、TRX、format、4パッケージの梱包が成功。成果物はローカル `artifacts/release/f7c8c111962945a79bb0b6865592fd0e/`。
* Visual Studioでも全644件成功。移行直後に選択範囲の一部を検出できない警告が1件出たが、その後の本体323件の選択実行では警告0件。原因は未確定。
* 次は未コミット差分の人間レビュー後にコミットし、Refactorを検討してPRを作成する。
## 停止位置

* 全体リファクタリングはレビュー・コミット了承済み。`codex/review-code-structure`（基準 `a204626`）で、Core・コード生成は `11f2d3c`、日付表示46件を含むテストの整理はこの記録と同じコミットにまとめた。
* テストでは日付の保存結果の検査を共有し、本文を読む前提は `RequiredDocument` で実際に検査する。読み取り専用のケースは固定文書を直接開く。
* 静的診断の既存PR #3とは作業ツリーを分けた。リポジトリに紐付いたGitHub Projectは確認できず、作業位置はここで管理する。リリースはSpreadsheetAsDataとの足並みを確認するまで保留。

## 検証

* .NET 11で全6プロジェクトを非増分ビルドし、警告・エラー0。SARIFに未抑制の診断なし。ソリューション全体のformat（severity info）成功。全644テスト成功（本体323、コード生成・MSBuild連携321）、失敗・Skipともに0。
* 基準版と変更版のコード生成を追跡対象の固定DOCX全172件で比較。生成C#（XMLコメントを含む）、診断、入力を受け付けなかった場合の例外の型・メッセージが一致。比較用プログラムと結果はメイン作業ツリーのローカル `artifacts/refactor-probe-e5ab3cac4f0f4f42b7806ee2cf662c99/` にある。

## 別作業

* **保存時の空白保持**：単純MERGEFIELD・複合MERGEFIELD・文字列Content Controlで、前後に空白がある文字列を保存しても `xml:space="preserve"` が付かない。既存テストはnull条件演算子がアサーションまで短絡していた。ユーザー判断で切り分け、同じ基準の `codex/xml-space-preservation` にアサーション3か所の未コミット差分を残した。3件とも属性がnullのため失敗することを再現済み。再開時はこのRedから。
* **改行正規化スクリプト**：サンプル入力 `abc\n \n` に対し1回目は `abc\r\n\r\n`、2回目は `abc\r\n`。末尾の空白行の扱いは別作業で検討する。
* B07のフォント修正はWord実機確認まで保留。これまでの指摘と性能候補は[レビュー記録](refactoring-review.md)、完了した繰り返し対応は[計画](repeating-sections-plan.md)を参照。
