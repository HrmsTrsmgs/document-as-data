namespace Marimo.DocumentAsData.Test.TestDocuments;

/// <summary>
/// ネットワーク入力などを模して、順方向の読み取りだけを公開します。
/// Length・Positionへの参照やシーク、書き込みは失敗させ、入力能力への依存を検出します。
/// </summary>
/// <param name="source">このラッパーと共に閉じる読み取り元。</param>
sealed class NonSeekableReadStream(Stream source) : Stream
{
    /// <inheritdoc />
    public override bool CanRead => source.CanRead;

    /// <inheritdoc />
    public override bool CanSeek => false;

    /// <inheritdoc />
    public override bool CanWrite => false;

    /// <inheritdoc />
    public override long Length => throw new NotSupportedException();

    /// <inheritdoc />
    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    /// <inheritdoc />
    public override int Read(byte[] buffer, int offset, int count) =>
        source.Read(buffer, offset, count);

    /// <inheritdoc />
    public override long Seek(long offset, SeekOrigin origin) =>
        throw new NotSupportedException();

    /// <inheritdoc />
    public override void Write(byte[] buffer, int offset, int count) =>
        throw new NotSupportedException();

    /// <inheritdoc />
    public override void SetLength(long value) =>
        throw new NotSupportedException();

    /// <inheritdoc />
    public override void Flush() => throw new NotSupportedException();

    /// <inheritdoc />
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            source.Dispose();
        }
        base.Dispose(disposing);
    }
}
