namespace Marimo.DocumentAsData.CodeGeneration;

/// <summary>
/// Word文書からC#コードを生成する前に検出された問題を表します。
/// </summary>
/// <param name="IsError">コード生成を継続できないエラーである場合は <see langword="true"/>。</param>
/// <param name="GeneratedName">診断対象になった生成後のC#名。</param>
/// <param name="SourceNames">同じ生成名に対応したWord文書上の元名。</param>
/// <param name="InvalidSourceName">有効なC#名へ変換できなかったWord文書上の元名。</param>
public sealed record CodeGenerationDiagnostic(
    bool IsError,
    string GeneratedName,
    string[] SourceNames,
    string? InvalidSourceName = null);
