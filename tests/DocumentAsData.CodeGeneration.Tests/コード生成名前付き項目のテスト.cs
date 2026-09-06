using FluentAssertions;
using Marimo.DocumentAsData.CodeGeneration.Test.テスト補助;

namespace Marimo.DocumentAsData.CodeGeneration.Test;

public sealed class コード生成名前付き項目のテスト
{
    const string MergeFieldsDocumentFilePath =
        @"TestData\コード生成\MERGEFIELD.docx";

    [Fact]
    public void MERGEFIELDをDocumentのMergeFieldプロパティとして生成します()
    {
        var tested = GeneratedCodeInspection
            .AssemblyFrom(
                GeneratedCodeInspection.GenerateSources(
                    MergeFieldsDocumentFilePath))
            .GeneratedType("MERGEFIELDDocument")
            .GetProperty("CustomerName");

        tested.Should().NotBeNull();
        tested.PropertyType.Should().Be(typeof(MergeField));
    }
}
