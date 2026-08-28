using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class Documentストリームのテスト
{
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

    [Fact(Skip = "MemoryStreamからのOpenをGreen対象にするときに有効化します。")]
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
}
