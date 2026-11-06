# 現在の作業状態

確認日：2026-11-06。主担当はこの記録とGit状態を照合して再開します。記録方法は[README](README.md)を参照してください。

## 停止位置と次の作業

* CR入りCheckBox Tagのコード生成修正は `4195d95` でコミット済み。親履歴なしのRefactor担当による検討も完了し、追加変更・追加Skipはありません。
* 工程担当は直接呼び出す `collaboration` ツールで起動・完了を確認済み。前回の「起動不可」は関数内ツール一覧だけを確認した誤判断で、環境設定の変更は不要でした。
* 次は他の改行を含む項目名のケースを一件ずつRedから進めます。[対応状況](refactoring-review.md#現在の対応状況)を参照。

## 最新の検証

* CR入りCheckBox Tag：`f68548e`＋Green・テスト・固定DOCX。全体628成功・失敗0・Skip 5（本体317、コード生成311）、コード生成3分56.8秒。解析Rebuild成功（18.3秒）、変更箇所のSARIF診断なし、既存診断13件。差分検査成功。Refactorは同一版を無変更と判断したため、ビルド・テストは再実行していません。
* SDK 10.0.401のMSBuild.dllを `dotnet exec` で直接実行し、ソリューションを `-t:Rebuild -restore:false -p:UseSharedCompilation=false -p:EnforceCodeStyleInBuild=true -p:RunAnalyzersDuringBuild=true -p:ErrorLog=obj/ide-diagnostics.sarif` で検証。その終了後、同SDKのvstest.console.dllでDebug/net10.0の本体・コード生成テストDLLを通常権限で直接実行しました。ビルドとテストは同時実行しません。
* `global.json`指定のSDK 10.0.302は未導入で、format検査は未完了。設定は変更していません。制限環境でのテスト応答停止は既知の未特定事項です。
* Skip 5件はRS-07の4件と明細0件の契約メモ1件。今回は新規テスト追加でSkip数に変化なし。未決定事項は[繰り返し計画](repeating-sections-plan.md)を参照します。

## 保留・環境制約

* B07のフォント修正は、2026-11-05のユーザー判断でWord操作復旧まで保留。仕様だけで先行修正する案は採用していません。[根拠と再開条件](refactoring-review.md#現在の対応状況)を参照。
* Wordの通常起動とウィンドウ一覧取得は成功。画面操作ツールの起動・状態取得は、画像を無効にしても応答しません。確認用Wordは終了済み。実機の表示・編集結果は未確認です。
* 調査用DOCXはローカルの `tests/DocumentAsData.Tests/obj/` 配下のみで、再現用の正式テストデータではありません。
