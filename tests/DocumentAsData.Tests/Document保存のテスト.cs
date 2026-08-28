using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class Document保存のテスト
{
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
