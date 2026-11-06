namespace Marimo.DocumentAsData.Test.TestDocuments;

/// <summary>
/// 作業ディレクトリの変更が他のテストへ影響しないよう、同時実行を禁止します。
/// </summary>
[CollectionDefinition(DisableParallelization = true)]
public sealed class CurrentDirectoryCollection;
