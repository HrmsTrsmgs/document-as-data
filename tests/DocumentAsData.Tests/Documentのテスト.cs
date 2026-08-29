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

    [Fact(Skip = "FileStreamからのOpenをGreen対象にするときに有効化します。")]
    public void OpenはFileStream上の文書を開きます()
    {
        var filePath = TestDocument.CreateCopy();
        try
        {
            using var stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite);

            var action = () =>
            {
                using var document = Document.Open(stream);
            };

            action.Should().NotThrow();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void OpenはMemoryStream上の文書を開きます()
    {
        using var stream = TestDocument.CreateMemoryStream();

        var action = () =>
        {
            using var document = Document.Open(stream);
        };

        action.Should().NotThrow();
    }

    [Fact(Skip = "Stream版での既存読み込み処理をGreen対象にするときに有効化します。")]
    public void Stream版でもMERGEFIELDを読み取れます()
    {
        using var stream = TestDocument.CreateMemoryStreamWithSimpleMergeField(
            "CustomerName",
            "株式会社○○");
        using var document = Document.Open(stream);

        document.MergeFields["CustomerName"].Value.Should().Be("株式会社○○");
    }

    [Fact(Skip = "Stream版での既存書き込み処理をGreen対象にするときに有効化します。")]
    public void Stream版でもMERGEFIELDの変更を文書へ書き込みます()
    {
        using var stream = TestDocument.CreateMemoryStreamWithSimpleMergeField(
            "CustomerName",
            "変更前");
        using (var document = Document.Open(stream))
        {
            document.MergeFields["CustomerName"].Value = "変更後";
        }

        stream.Position = 0;
        using var saved = Document.Open(stream);

        saved.MergeFields["CustomerName"].Value.Should().Be("変更後");
    }

    [Fact(Skip = "呼び出し側から渡されたStreamの所有権をGreen対象にするときに有効化します。")]
    public void Disposeは呼び出し側から渡されたStreamを閉じません()
    {
        using var stream = TestDocument.CreateMemoryStream();
        var document = Document.Open(stream);

        document.Dispose();

        stream.CanRead.Should().BeTrue();
        stream.CanWrite.Should().BeTrue();
    }

    [Fact(Skip = "SaveAsによるDOCX作成をGreen対象にするときに有効化します。")]
    public void SaveAsは指定したパスへDOCXを作成します()
    {
        var sourcePath = TestDocument.CreateCopy();
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using var document = Document.Open(sourcePath);

            document.SaveAs(outputPath);

            File.Exists(outputPath).Should().BeTrue();
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact(Skip = "SaveAsしたDOCXの再オープンをGreen対象にするときに有効化します。")]
    public void SaveAsしたDOCXをDocumentとして開けます()
    {
        var sourcePath = TestDocument.CreateCopy();
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.SaveAs(outputPath);
            }

            var action = () =>
            {
                using var saved = Document.Open(outputPath);
            };

            action.Should().NotThrow();
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact(Skip = "MERGEFIELD更新結果の保存をGreen対象にするときに有効化します。")]
    public void SaveAsはMERGEFIELDへ設定した値を保存します()
    {
        var sourcePath = TestDocument.CreateWithSimpleMergeFields(("CustomerName", "変更前"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Value = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.MergeFields["CustomerName"].Value.Should().Be("変更後");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact(Skip = "Content Control更新結果の保存をGreen対象にするときに有効化します。")]
    public void SaveAsはContentControlへ設定した値を保存します()
    {
        var sourcePath = TestDocument.CreateWithContentControls(("CustomerName", "変更前"));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Value = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Value.Should().Be("変更後");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }
}
