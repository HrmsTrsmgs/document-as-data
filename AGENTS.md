# DocumentAsData 作業指示

## 読むべき指示

* 進行役は、委任の利用条件とRGRのモデル・推論レベルを [モデル方針](docs/ai-instructions/tdd-context/models.md) に従う。導入済み共通Skillに旧割り当てが残る場合も、この方針を優先する。
* 共通方針は個人共通AGENTSを使う。未導入時は [参照版](docs/ai-instructions/common.md) を読む。
* 振る舞い追加・仕様変更・不具合修正は共通Skill `tdd-workflow`、挙動を維持するリファクタリングの進行は同Skillの `references/run.md` を入口とする。未導入時は [TDDの入口](docs/ai-instructions/tdd.md) を使う。運用文書の編集や性能調査・計測だけでは製品TDDを起動しない。
* C#/.NETは共通Skill `csharp-conventions` と [プロジェクトの補足](docs/ai-instructions/csharp.md#プロジェクトの補足) を使う。未導入時だけ同文書の参照版を読む。
* 固有の仕様・制約は [DocumentAsDataの指示](docs/ai-instructions/document-as-data.md) を確認する。工程担当は担当範囲の節だけを読む。

共通Skillと参照版は同じ用途に重ねて読まない。進行役は選んだ進行手順、起動済みの工程担当は同じ側の工程共通と担当工程だけを工程指示として読む。担当の再開時に入口を再実行せず、候補一覧や他工程の議論を渡さない。技術・固有規約は必要な範囲で補足する。私用の共通リポジトリへのアクセスや導入をOSSへの参加条件にしない。

## 仕様・判断と記録先

主担当は開始・再開時に [記録の索引](ai-notes/README.md)、[現在の作業](ai-notes/current.md)、必要な仕様・判断を読み、現在のGit状態と照合する。記録の更新・重複整理は共通AGENTSの「仕様・判断と再開記録」に従う。

| 内容 | 記録先 |
| --- | --- |
| 確定仕様・利用例 | 既存の機能文書・テストの該当箇所 |
| 判断理由・不採用案 | [docs/design-decisions.md](docs/design-decisions.md) |
| 未確定仕様・候補 | [ai-notes/current.md](ai-notes/current.md)、具体化済みの保留仕様はSkipテスト |
| 現在の工程・停止位置・検証結果・次の着手点 | [ai-notes/current.md](ai-notes/current.md) を置き換え更新 |

## 指示の更新

他の開発にも共有できる規則は `codex-workflows` の共通AGENTS・技術別Skill・工程別指示を正本とし、未導入環境向け参照版も同期する。製品の仕様、採用パッケージ、検証対象、記録先はこのプロジェクトに残す。一時的な操作ログや未合意の案を共通規則へ昇格させず、他プロジェクト固有のデータ・公開手順・実装を機械的に転用しない。
