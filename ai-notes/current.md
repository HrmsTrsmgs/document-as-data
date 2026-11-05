# 現在の作業状態

確認日：2026-11-05。主担当はこの記録とGit状態を照合して再開します。記録方法は[README](README.md)を参照してください。

## 停止位置と次の作業

* 明細型名との衝突診断は `4c76f7d` でGreenをコミット済み。
* Refactorで明細の文字列Content ControlとMERGEFIELDの元名取得を共通化し、診断の元名順序を確認するテスト1件を追加。実装・テストはユーザー了承済みで、この記録とともにコミットします。
* 次は別の衝突ケースへ進みます。[指摘と対応状況](refactoring-review.md#現在の対応状況)を参照。

## 最新の検証

* Refactor検証版：`4c76f7d`＋元名取得の共通化とテスト1件。全体626成功・失敗0・Skip 5（本体317、コード生成309）、コード生成3分49秒。解析Rebuild成功、変更ファイルのSARIF診断なし。差分検査成功。通常権限のテストホスト接続が失敗し、昇格した再実行で成功しました。
* SDK 10.0.401のMSBuild.dllを `dotnet exec` で直接実行し、ソリューションを `-t:Rebuild -restore:false -p:UseSharedCompilation=false -p:EnforceCodeStyleInBuild=true -p:RunAnalyzersDuringBuild=true -p:ErrorLog=obj/ide-diagnostics.sarif` で検証。その終了後、同SDKのvstest.console.dllでDebug/net10.0の本体・コード生成テストDLLを通常権限で直接実行しました。ビルドとテストは同時実行しません。
* `global.json`指定のSDK 10.0.302は未導入で、format検査は未完了。設定は変更していません。制限環境でのテスト応答停止は既知の未特定事項です。
* Skip 5件はRS-07の4件と明細0件の契約メモ1件。今回は新規テスト追加でSkip数に変化なし。未決定事項は[繰り返し計画](repeating-sections-plan.md)を参照します。

## 保留・環境制約

* B07のフォント修正は、2026-11-05のユーザー判断でWord操作復旧まで保留。仕様だけで先行修正する案は採用していません。[根拠と再開条件](refactoring-review.md#現在の対応状況)を参照。
* Wordの通常起動とウィンドウ一覧取得は成功。画面操作ツールの起動・状態取得は、画像を無効にしても応答しません。確認用Wordは終了済み。実機の表示・編集結果は未確認です。
* 調査用DOCXはローカルの `tests/DocumentAsData.Tests/obj/` 配下のみで、再現用の正式テストデータではありません。
