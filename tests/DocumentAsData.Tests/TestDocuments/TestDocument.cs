using System.IO.Compression;
using System.Security;
using System.Text;

namespace Marimo.DocumentAsData.Test.TestDocuments;

static class TestDocument
{
    public static string CreateCopy() =>
        CreateWithBody("<w:p />");

    public static string CreateOutputPath() =>
        Path.Combine(
            Path.GetTempPath(),
            $"DocumentAsData-Output-{Guid.NewGuid():N}.docx");

    public static MemoryStream CreateMemoryStream()
    {
        var filePath = CreateCopy();
        try
        {
            return new MemoryStream(File.ReadAllBytes(filePath));
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    public static MemoryStream CreateMemoryStreamWithSimpleMergeField(
        string name,
        string value)
    {
        var filePath = CreateWithSimpleMergeFields((name, value));
        try
        {
            return new MemoryStream(File.ReadAllBytes(filePath));
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    public static string CreateWithSimpleMergeFields(params (string Name, string Value)[] fields) =>
        CreateWithBody(
            fields.Select(it =>
                $"""
                <w:p>
                  <w:fldSimple w:instr=" MERGEFIELD &quot;{Escape(it.Name)}&quot; ">
                    <w:r><w:t>{Escape(it.Value)}</w:t></w:r>
                  </w:fldSimple>
                </w:p>
                """).ToArray());

    public static string CreateWithComplexMergeField(string name, string value) =>
        CreateWithBody(
            $"""
            <w:p>
              <w:r><w:fldChar w:fldCharType="begin" /></w:r>
              <w:r><w:instrText xml:space="preserve"> MERGEFIELD &quot;{Escape(name)}&quot; </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate" /></w:r>
              <w:r><w:t>{Escape(value)}</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end" /></w:r>
            </w:p>
            """);

    public static string CreateWithSplitComplexMergeField(string name, string value) =>
        CreateWithBody(
            $"""
            <w:p>
              <w:r><w:fldChar w:fldCharType="begin" /></w:r>
              <w:r><w:instrText xml:space="preserve"> MERGE</w:instrText></w:r>
              <w:r><w:instrText xml:space="preserve">FIELD &quot;{Escape(name)}&quot; </w:instrText></w:r>
              <w:r><w:fldChar w:fldCharType="separate" /></w:r>
              <w:r><w:t>{Escape(value)}</w:t></w:r>
              <w:r><w:fldChar w:fldCharType="end" /></w:r>
            </w:p>
            """);

    public static string CreateWithContentControls(params (string? Tag, string Value)[] controls) =>
        CreateWithBody(
            controls.Select(it =>
                $"""
                <w:sdt>
                  <w:sdtPr>{Tag(it.Tag)}</w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>{Escape(it.Value)}</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>
                """).ToArray());

    public static string CreateWithTableMergeField(string name, string value) =>
        CreateWithBody(
            $"""
            <w:tbl>
              <w:tr>
                <w:tc>
                  <w:p>
                    <w:fldSimple w:instr=" MERGEFIELD &quot;{Escape(name)}&quot; ">
                      <w:r><w:t>{Escape(value)}</w:t></w:r>
                    </w:fldSimple>
                  </w:p>
                </w:tc>
              </w:tr>
            </w:tbl>
            """);

    public static string CreateWithTableContentControl(string tag, string value) =>
        CreateWithBody(
            $"""
            <w:tbl>
              <w:tr>
                <w:tc>
                  <w:sdt>
                    <w:sdtPr><w:tag w:val="{Escape(tag)}" /></w:sdtPr>
                    <w:sdtContent><w:p><w:r><w:t>{Escape(value)}</w:t></w:r></w:p></w:sdtContent>
                  </w:sdt>
                </w:tc>
              </w:tr>
            </w:tbl>
            """);

    public static string CreateWithNestedContentControls(
        string outerTag,
        string innerTag,
        string value) =>
        CreateWithBody(
            $"""
            <w:sdt>
              <w:sdtPr><w:tag w:val="{Escape(outerTag)}" /></w:sdtPr>
              <w:sdtContent>
                <w:sdt>
                  <w:sdtPr><w:tag w:val="{Escape(innerTag)}" /></w:sdtPr>
                  <w:sdtContent><w:p><w:r><w:t>{Escape(value)}</w:t></w:r></w:p></w:sdtContent>
                </w:sdt>
              </w:sdtContent>
            </w:sdt>
            """);

    static string CreateWithBody(params string[] bodyContents)
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
            $"""
            <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>{string.Join(Environment.NewLine, bodyContents)}<w:sectPr /></w:body>
            </w:document>
            """);

        return filePath;
    }

    static string Escape(string value) =>
        SecurityElement.Escape(value) ?? "";

    static string Tag(string? tag) =>
        tag == null
            ? ""
            : $"<w:tag w:val=\"{Escape(tag)}\" />";

    static void WriteEntry(ZipArchive archive, string entryName, string content)
    {
        using var writer = new StreamWriter(
            archive.CreateEntry(entryName).Open(),
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(content);
    }
}
