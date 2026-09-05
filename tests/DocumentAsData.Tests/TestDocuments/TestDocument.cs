namespace Marimo.DocumentAsData.Test.TestDocuments;

static class TestDocument
{
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
