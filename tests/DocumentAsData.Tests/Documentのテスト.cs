using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class Documentのテスト
{
    [Fact]
    public void Openはファイルを束縛します()
    {
        var copyPath = TestDocument.CreateCopy();
        try
        {
            var tested = Document.Open(copyPath);
            try
            {
                FluentActions.Invoking(
                    () => File.Delete(copyPath)
                ).Should().Throw<IOException>();
            }
            finally
            {
                tested.Close();
            }
        }
        finally
        {
            File.Delete(copyPath);
        }
    }

    [Fact]
    public void Closeはファイルの束縛を解除します()
    {
        var copyPath = TestDocument.CreateCopy();
        var tested = Document.Open(copyPath);

        tested.Close();
        FluentActions.Invoking(
            () => File.Delete(copyPath)
        ).Should().NotThrow();
    }

    [Fact]
    public void Disposeはファイルの束縛を解除します()
    {
        var copyPath = TestDocument.CreateCopy();
        var tested = Document.Open(copyPath);
        var disposable = tested as IDisposable;

        disposable.Should().NotBeNull();
        disposable.Dispose();
        FluentActions.Invoking(
            () => File.Delete(copyPath)
        ).Should().NotThrow();
    }
}
