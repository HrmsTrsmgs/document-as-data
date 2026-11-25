using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using Wordprocessing = DocumentFormat.OpenXml.Wordprocessing;

namespace Marimo.DocumentAsData.Test.TestDocuments;

static class TestDocument
{
    extension(WordprocessingDocument self)
    {
        /// <summary>
        /// テストで本文の要素を調べる前に、メインパートの文書ルートが存在することを確認します。
        /// </summary>
        internal Wordprocessing.Document RequiredDocument
        {
            get
            {
                var document = self.MainDocumentPart?.Document;
                document.Should().NotBeNull();
                return document;
            }
        }
    }

    public static string CreateOutputPath() =>
        Path.Combine(
            Path.GetTempPath(),
            $"DocumentAsData-Output-{Guid.NewGuid():N}.docx");

    public static string CreateTemporaryCopy(string sourcePath)
    {
        var filePath = CreateOutputPath();
        File.Copy(sourcePath, filePath);
        return filePath;
    }

    public static MemoryStream CreateMemoryStream(string sourcePath)
    {
        var stream = new MemoryStream();
        stream.Write(File.ReadAllBytes(sourcePath));
        stream.Position = 0;
        return stream;
    }
}
