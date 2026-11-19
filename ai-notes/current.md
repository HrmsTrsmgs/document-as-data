# 現在の作業状態

確認日：2026-11-20。主担当はこの記録とGit状態を照合して再開します。記録方法は[README](README.md)を参照してください。

## 停止位置と次の作業

* V05は `66e091b` でコミット済み。Refactorは追加変更なし。R01はMSBuildの生成名診断とソース生成で1回開いたDOCXを共有する変更まで完了し、レビュー待ち。対象はBuildのタスクとCodeGenerationのAssemblyInfo。作業開始前からあるGeneratedNameDiagnostics.csの差分は別件。

## 最新の検証

* SDK 10.0.302でR01を検証。全体638成功・失敗0・既存Skip 5。非増分解析ビルドは警告・エラー0。formatは既存のxUnit1004の5件とGeneratedNameDiagnostics.csの空白11件で停止。R01対象への指摘なし。

## 保留・環境制約

* B07のフォント修正は、2026-11-05のユーザー判断でWord操作復旧まで保留。[根拠と再開条件](refactoring-review.md#現在の対応状況)を参照。
* Wordの通常起動とウィンドウ一覧取得は成功。画面操作ツールの起動・状態取得は応答せず、実機の表示・編集結果は未確認です。確認用Wordは終了済み。
* 調査用DOCXはローカルの `tests/DocumentAsData.Tests/obj/` 配下のみで、正式テストデータではありません。
