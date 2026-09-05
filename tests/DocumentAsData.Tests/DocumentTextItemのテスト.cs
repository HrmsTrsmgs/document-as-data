using FluentAssertions;

namespace Marimo.DocumentAsData.Test;

public class DocumentTextItemのテスト
{
    [Fact]
    public void MergeFieldとContentControlはDocumentTextItemです()
    {
        Type[] itemTypes =
        [
            typeof(MergeField),
            typeof(ContentControl)
        ];

        itemTypes.Should().AllSatisfy(
            type => type.Should().BeDerivedFrom<DocumentTextItem>());
    }
}
