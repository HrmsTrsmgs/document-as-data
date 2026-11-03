---
name: document-as-data-tdd
description: DocumentAsDataの既存TDD入口として指定されたときに、共通tdd-workflowへ進行を委ねる互換入口。共通Skill未導入時はリポジトリ内の工程別参照版を使う。運用文書の編集や性能計測だけの依頼では起動しない。
---

# DocumentAsDataのTDD入口

[AGENTS.md](../../../AGENTS.md) の固有制約と記録先を確認する。共通Skill `tdd-workflow` が利用可能なら、その `SKILL.md` を進行役の入口にする。

未導入時は [リポジトリ内の進行手順](../../../docs/ai-instructions/tdd-context/run.md) だけを読む。各工程は履歴を継承しない担当へ順番に渡し、担当へは同じ参照版の `common.md` と担当工程文書だけを指定する。同一チャットの役割変更を履歴分離と呼ばない。

共通Skillと参照版の規則を重ねて読み込まず、既存の人間レビュー・コミット了承を維持する。このSkillの編集だけでは製品TDDを開始しない。起動済みの工程担当はこの入口を再実行しない。
