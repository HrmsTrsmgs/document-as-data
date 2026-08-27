using System.IO.Compression;
using System.Text;

namespace Marimo.DocumentAsData.Test.TestDocuments;

static class TestDocument
{
    public static string CreateCopy()
    {
        var filePath = Path.Combine(
            Path.GetTempPath(),
            $"DocumentAsData-{Guid.NewGuid():N}.docx");

        using var archive = ZipFile.Open(filePath, ZipArchiveMode.Create);

        WriteEntry(
            archive,
            "[Content_Types].xml",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
              <Default Extension="xml" ContentType="application/xml" />
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
            </Types>
            """);
        WriteEntry(
            archive,
            "_rels/.rels",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />
            </Relationships>
            """);
        WriteEntry(
            archive,
            "word/document.xml",
            """
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body><w:p /></w:body>
            </w:document>
            """);

        return filePath;
    }

    static void WriteEntry(ZipArchive archive, string entryName, string content)
    {
        using var writer = new StreamWriter(
            archive.CreateEntry(entryName).Open(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }
}
