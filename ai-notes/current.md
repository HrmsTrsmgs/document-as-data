# 現在の作業状態

確認日：2026-11-28。

## 停止位置

* `codex/dry-refactoring`、基準 `a8d5af6`。保留していたMSBuild生成ファイル名の接尾辞を、タスクと利用側のMSBuild設定で共有した。実装と検証は確認済み。DRY変更はPRでmainへの取り込み待ち。
* [GitHub Project](https://github.com/users/HrmsTrsmgs/projects/1)の接尾辞共通化は進行中。このブランチのDRY変更はローカルにあり、PRで取り込む。名前解決の共通化は複雑さに見合わず保留。生成XMLコメントの警告と生成名衝突の扱いは別作業。リリースはSpreadsheetAsDataとの足並みを確認するまで保留。

## 検証

* 接尾辞共通化の未コミット差分でRelease非増分ビルド（警告・エラー0）、全644テスト（失敗・Skip 0）、書式・静的診断・差分検査、警告をエラーにしたBuildパッケージの梱包に成功。配布物の命名用`.props`はソースと一致した。既定の出力先はDLLロックでビルドに失敗したため、作業ツリー内のプロジェクトだけ別の出力先で検証した。
* `.targets`単独と`.props`を含む読み込みで生成ファイル名・IDE向けの項目登録を確認した。評価後に追加した文書のCleanは、対象の生成ソースだけを削除することをdotnet版とVisual Studio版MSBuildで確認した。Visual Studio版MSBuildのパッケージ・共用ライブラリ連携13テストも成功した。
* 生成対象の列挙の共通化では、固定DOCX 172件の生成ソース、診断、例外の型と文言が基準版と一致した。

## 別作業

* **保存時の空白保持**：単純MERGEFIELD・複合MERGEFIELD・文字列Content Controlで、前後に空白がある文字列を保存しても `xml:space="preserve"` が付かない。既存テストはnull条件演算子がアサーションまで短絡していた。ユーザー判断で切り分け、同じ基準の `codex/xml-space-preservation` にアサーション3か所の未コミット差分を残した。3件とも属性がnullのため失敗することを再現済み。再開時はこのRedから。
* **改行正規化スクリプト**：サンプル入力 `abc\n \n` に対し1回目は `abc\r\n\r\n`、2回目は `abc\r\n`。末尾の空白行の扱いは別作業で検討する。
* B07のフォント修正はWord実機確認まで保留。これまでの指摘と性能候補は[レビュー記録](refactoring-review.md)、完了した繰り返し対応は[計画](repeating-sections-plan.md)を参照。
