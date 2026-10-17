---
name: document-as-data-tdd
description: DocumentAsDataの振る舞い追加、仕様変更、不具合修正を、Red・Green・Refactorの独立した担当コンテキストで順番に進める。
---

# 工程別TDDの入口

進行役として `docs/ai-instructions/tdd-context/run.md` だけを読む。工程担当の指示や思考履歴を進行役へ読み込まない。

各工程には履歴を継承しない新しい担当コンテキストを順番に割り当てる。担当へは `docs/ai-instructions/tdd-context/common.md` と担当工程の文書だけを指定する。別ファイルに分けただけの同一会話内の役割変更を、コンテキスト分離とは呼ばない。

担当同士を相談させず、工程間のAIレビュー・AI承認を増やさない。既存の人間レビュー、権限承認、コミット・push規則は維持する。worktreeはコンテキスト分割のために作り直さない。

このSkillは製品開発の開始要求があるときに使う。運用文書の導入・編集だけを依頼された場合は、Red・Green・Refactorを起動しない。
