using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成コメントのテスト
{
    const string BasicStructureDocumentFilePath =
        @"TestData\コード生成\BasicStructure.docx";

    [Fact]
    public void Document型のコメントを生成します()
    {
        GeneratedCodeInspection
            .GenerateSources(BasicStructureDocumentFilePath)
            .TypeDeclaration("BasicStructureDocument")
            .SummaryText()
            .Should().Be("Word文書「BasicStructure」を型付きで表します。");
    }
}
