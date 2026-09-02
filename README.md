# document-as-data
Word文書を構造化データとして扱うための.NETライブラリ

現在は v0.1 の設計・開発初期段階です。公開API案と対象範囲は
[v0.1 設計案](docs/design-v0.1.md)を参照してください。

## 現在対応していないもの

* `w:dataBinding` によりCustom XMLへデータバインドされたContent Control

データバインドされたContent Controlでは、表示内容とは別にCustom XML側にも値が保持されます。
DocumentAsDataは現在Custom XMLとの同期を行わないため、読み書きの対象には使用しないでください。
