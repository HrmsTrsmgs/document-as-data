using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

public class Documentのテスト
{
    static readonly string EmptyDocumentPath =
        Path.Combine("TestData", "空の文書.docx");
    static readonly string SimpleMergeFieldsPath =
        Path.Combine("TestData", "単純形式のMERGEFIELD.docx");
    static readonly string ComplexMergeFieldPath =
        Path.Combine("TestData", "複合形式のMERGEFIELD.docx");
    static readonly string ContentControlPath =
        Path.Combine("TestData", "単一のContent Control.docx");
    static readonly string ContentControlsPath =
        Path.Combine("TestData", "複数のContent Control.docx");
    static readonly string AmbiguousCustomerNamePath =
        Path.Combine("TestData", "同名のMERGEFIELDとContent Control.docx");
    static readonly string InvalidContentControlWithoutTagValuePath =
        Path.Combine("TestData", "Tagの値が欠落した不正なContent Control.docx");
    static readonly string DatePickerContentControlPath =
        Path.Combine("TestData", "日付選択のContent Control.docx");

    [Fact]
    public void Openはファイルを束縛します()
    {
        var copyPath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
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
        var copyPath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        var tested = Document.Open(copyPath);

        tested.Close();
        FluentActions.Invoking(
            () => File.Delete(copyPath)
        ).Should().NotThrow();
    }

    [Fact]
    public void Disposeはファイルの束縛を解除します()
    {
        var copyPath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        var tested = Document.Open(copyPath);
        var disposable = tested as IDisposable;

        disposable.Should().NotBeNull();
        disposable.Dispose();
        FluentActions.Invoking(
            () => File.Delete(copyPath)
        ).Should().NotThrow();
    }

    [Fact]
    public void OpenはFileStream上の文書を開きます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
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
        using var stream = TestDocument.CreateMemoryStream(EmptyDocumentPath);

        var action = () =>
        {
            using var document = Document.Open(stream);
        };

        action.Should().NotThrow();
    }

    [Fact]
    public void Openは検証指定を省略した場合不正なOOXMLも開きます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(InvalidContentControlWithoutTagValuePath);
        try
        {
            var action = () =>
            {
                using var document = Document.Open(filePath);
            };

            action.Should().NotThrow();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Openは検証しない場合正常なOOXMLを開きます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        try
        {
            var action = () =>
            {
                using var document = Document.Open(filePath, false);
            };

            action.Should().NotThrow();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Openは検証しない場合不正なOOXMLも開きます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(InvalidContentControlWithoutTagValuePath);
        try
        {
            var action = () =>
            {
                using var document = Document.Open(filePath, false);
            };

            action.Should().NotThrow();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Openは検証する場合正常なOOXMLを開きます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        try
        {
            var action = () =>
            {
                using var document = Document.Open(filePath, true);
            };

            action.Should().NotThrow();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Openは検証する場合不正なOOXMLで失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(InvalidContentControlWithoutTagValuePath);
        try
        {
            var action = () =>
            {
                using var document = Document.Open(filePath, true);
            };

            action.Should().Throw<InvalidDataException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Openは検証に失敗してもファイルを束縛しません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(InvalidContentControlWithoutTagValuePath);
        try
        {
            FluentActions.Invoking(
                () => Document.Open(filePath, true)
            ).Should().Throw<InvalidDataException>();

            FluentActions.Invoking(
                () => File.Delete(filePath)
            ).Should().NotThrow();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Stream版Openは検証指定を省略した場合不正なOOXMLも開きます()
    {
        using var stream = TestDocument.CreateMemoryStream(InvalidContentControlWithoutTagValuePath);

        var action = () =>
        {
            using var document = Document.Open(stream);
        };

        action.Should().NotThrow();
    }

    [Fact]
    public void Stream版Openは検証しない場合正常なOOXMLを開きます()
    {
        using var stream = TestDocument.CreateMemoryStream(EmptyDocumentPath);

        var action = () =>
        {
            using var document = Document.Open(stream, false);
        };

        action.Should().NotThrow();
    }

    [Fact]
    public void Stream版Openは検証しない場合不正なOOXMLも開きます()
    {
        using var stream = TestDocument.CreateMemoryStream(InvalidContentControlWithoutTagValuePath);

        var action = () =>
        {
            using var document = Document.Open(stream, false);
        };

        action.Should().NotThrow();
    }

    [Fact]
    public void Stream版Openは検証する場合正常なOOXMLを開きます()
    {
        using var stream = TestDocument.CreateMemoryStream(EmptyDocumentPath);

        var action = () =>
        {
            using var document = Document.Open(stream, true);
        };

        action.Should().NotThrow();
    }

    [Fact]
    public void Stream版Openは検証する場合不正なOOXMLで失敗します()
    {
        using var stream = TestDocument.CreateMemoryStream(InvalidContentControlWithoutTagValuePath);

        var action = () =>
        {
            using var document = Document.Open(stream, true);
        };

        action.Should().Throw<InvalidDataException>();
    }

    [Fact]
    public void Stream版Openは検証に失敗しても呼び出し側のStreamを閉じません()
    {
        using var stream = TestDocument.CreateMemoryStream(InvalidContentControlWithoutTagValuePath);

        FluentActions.Invoking(
            () => Document.Open(stream, true)
        ).Should().Throw<InvalidDataException>();

        stream.CanRead.Should().BeTrue();
        stream.CanWrite.Should().BeTrue();
    }

    [Fact]
    public void Stream版でもMERGEFIELDを読み取れます()
    {
        using var stream = TestDocument.CreateMemoryStream(SimpleMergeFieldsPath);
        using var document = Document.Open(stream);

        document.MergeFields["CustomerName"].Text.Should().Be("株式会社○○");
    }

    [Fact]
    public void Stream版でもMERGEFIELDの変更を文書へ書き込みます()
    {
        using var stream = TestDocument.CreateMemoryStream(SimpleMergeFieldsPath);
        using (var document = Document.Open(stream))
        {
            document.MergeFields["CustomerName"].Text = "変更後";
        }

        stream.Position = 0;
        using var saved = Document.Open(stream);

        saved.MergeFields["CustomerName"].Text.Should().Be("変更後");
    }

    [Fact]
    public void Disposeは呼び出し側から渡されたStreamを閉じません()
    {
        using var stream = TestDocument.CreateMemoryStream(EmptyDocumentPath);
        var document = Document.Open(stream);

        document.Dispose();

        stream.CanRead.Should().BeTrue();
        stream.CanWrite.Should().BeTrue();
    }

    [Fact]
    public void Readはプロパティ名と同じTagのContentControlからオブジェクトを読み込みます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.Read<DocumentData>()
                .Should().BeEquivalentTo(
                    new DocumentData
                    {
                        CustomerName = "山田太郎",
                        Address = "東京都"
                    });
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Readはプロパティ名と同じMERGEFIELDからオブジェクトを読み込みます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.Read<CustomerNameOnlyDocumentData>()
                .Should().BeEquivalentTo(
                    new CustomerNameOnlyDocumentData
                    {
                        CustomerName = "株式会社○○"
                    });
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Readは自動変換したプロパティ名と一致するTagからオブジェクトを読み込みます()
    {
        using var document = Document.Open(@"TestData\customerForm.docx");

        // 文書のTagはcustomerNameです。対応属性を付けずにCustomerNameへ読み込みます。
        document.Read<CustomerNameOnlyDocumentData>()
            .CustomerName.Should().Be("山田太郎");
    }

    [Fact]
    public void ReadはDocumentItemName属性で指定した名前からオブジェクトを読み込みます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.Read<AttributedDocumentData>()
                .Should().BeEquivalentTo(
                    new AttributedDocumentData
                    {
                        Name = "山田太郎",
                        Location = "東京都"
                    });
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Readは日付選択ContentControlからDateTimeOffsetプロパティを読み込みます()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.Read<DateDocumentData>().DeliveryDate
                .Should().Be(
                    new DateTimeOffset(
                        2026,
                        9,
                        4,
                        0,
                        0,
                        0,
                        TimeSpan.Zero));
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Readは対応する日付選択ContentControlがない場合に失敗します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            var action = () => document.Read<MissingDateDocumentData>();

            action.Should().Throw<DocumentMappingException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Readは同名のMERGEFIELDとContentControlがある場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(AmbiguousCustomerNamePath);
        try
        {
            using var document = Document.Open(filePath, true);

            var action = () =>
                document.Read<CustomerNameOnlyDocumentData>();

            action.Should().Throw<InvalidOperationException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Replaceはプロパティ名と同じTagのContentControlへオブジェクトを書き込みます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.Replace(
                new DocumentData
                {
                    CustomerName = "変更後の氏名",
                    Address = "変更後の住所"
                });

            document.ContentControls["CustomerName"].Text
                .Should().Be("変更後の氏名");
            document.ContentControls["Address"].Text
                .Should().Be("変更後の住所");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ReplaceはDocumentItemName属性で指定した名前へオブジェクトを書き込みます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.Replace(
                new AttributedDocumentData
                {
                    Name = "変更後の氏名",
                    Location = "変更後の住所"
                });

            document.ContentControls["CustomerName"].Text
                .Should().Be("変更後の氏名");
            document.ContentControls["Address"].Text
                .Should().Be("変更後の住所");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ReplaceはDocumentItemName属性で指定した項目が存在しない場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () =>
                document.Replace(
                    new MissingDocumentItemData
                    {
                        Value = "変更後"
                    });

            action.Should().Throw<DocumentMappingException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Replaceは複数プロパティが同じ文書項目を指定した場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () =>
                document.Replace(
                    new DuplicateDocumentItemData
                    {
                        FirstValue = "1",
                        SecondValue = "2"
                    });

            action.Should().Throw<DocumentMappingException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ReplaceはDocumentItemName属性を付けたプロパティにpublicなgetterがない場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () =>
                document.Replace(
                    new AttributedPropertyWithoutPublicGetterData());

            action.Should().Throw<DocumentMappingException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ReplaceはDocumentItemName属性のない書き込み専用プロパティを無視します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.Replace(
                new DocumentDataWithWriteOnlyProperty
                {
                    CustomerName = "変更後の氏名"
                });

            document.ContentControls["CustomerName"].Text
                .Should().Be("変更後の氏名");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Replaceは対応していないプロパティ型の場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            var action = () =>
                document.Replace(
                    new UnsupportedPropertyTypeData
                    {
                        Value = DateTime.Today
                    });

            action.Should().Throw<DocumentMappingException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void ReplaceはDateTimeOffsetプロパティを日付選択ContentControlへ書き込みます()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            document.Replace(
                new DateDocumentData
                {
                    DeliveryDate = new DateTimeOffset(
                        2027,
                        1,
                        2,
                        0,
                        0,
                        0,
                        TimeSpan.FromHours(9))
                });

            document.DatePickers["DeliveryDate"].SelectedDateTime
                .Should().Be(
                    new DateTimeOffset(
                        2027,
                        1,
                        2,
                        0,
                        0,
                        0,
                        TimeSpan.FromHours(9)));
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Replaceは対応する日付選択ContentControlがない場合に失敗します()
    {
        var filePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        try
        {
            using var document = Document.Open(filePath, true);

            var action = () =>
                document.Replace(
                    new MissingDateDocumentData
                    {
                        DeliveryDate = new DateTimeOffset(
                            2027,
                            1,
                            2,
                            0,
                            0,
                            0,
                            TimeSpan.FromHours(9))
                    });

            action.Should().Throw<DocumentMappingException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Replaceはプロパティ名と同じMERGEFIELDへオブジェクトを書き込みます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.Replace(
                new CustomerNameOnlyDocumentData
                {
                    CustomerName = "変更後の氏名"
                });

            document.MergeFields["CustomerName"].Text
                .Should().Be("変更後の氏名");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Replaceは同名のMERGEFIELDとContentControlがある場合に失敗します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(AmbiguousCustomerNamePath);
        try
        {
            using var document = Document.Open(filePath, true);

            var action = () =>
                document.Replace(
                    new CustomerNameOnlyDocumentData
                    {
                        CustomerName = "変更後"
                    });

            action.Should().Throw<InvalidOperationException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Replaceは書き込み元に対応プロパティがないContentControlを変更しません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(ContentControlsPath);
        try
        {
            using var document = Document.Open(filePath);

            document.Replace(
                new CustomerNameOnlyDocumentData
                {
                    CustomerName = "変更後の氏名"
                });

            document.ContentControls["CustomerName"].Text
                .Should().Be("変更後の氏名");
            document.ContentControls["Address"].Text
                .Should().Be("東京都");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void SaveAsは指定したパスへDOCXを作成します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
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

    [Fact]
    public void SaveAsは保存先ファイルを束縛しません()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using var document = Document.Open(sourcePath);

            document.SaveAs(outputPath);

            FluentActions.Invoking(
                () => File.Delete(outputPath)
            ).Should().NotThrow();
        }
        finally
        {
            File.Delete(sourcePath);
        }
    }

    [Fact]
    public void SaveAsしたDOCXをDocumentとして開けます()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
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

    [Fact]
    public void SaveAsはMERGEFIELDへ設定した値を保存します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Text = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.MergeFields["CustomerName"].Text.Should().Be("変更後");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SaveAsは保存元ファイルを変更しません()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Text = "変更後";
                document.SaveAs(outputPath);
            }

            using var source = Document.Open(sourcePath);
            source.MergeFields["CustomerName"].Text.Should().Be("株式会社○○");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void Disposeは保存元ファイルへ変更を書き込みません()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Text = "変更後";
            }

            using var source = Document.Open(sourcePath);
            source.MergeFields["CustomerName"].Text.Should().Be("株式会社○○");
        }
        finally
        {
            File.Delete(sourcePath);
        }
    }

    [Fact]
    public void SaveAsはContentControlへ設定した値を保存します()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Text = "変更後";
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath);
            saved.ContentControls["CustomerName"].Text.Should().Be("変更後");
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SaveAsで保存した単純MERGEFIELD文書は検証して開けます()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Text = "変更後";
                document.SaveAs(outputPath);
            }

            var action = () =>
            {
                using var saved = Document.Open(outputPath, true);
            };

            action.Should().NotThrow();
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SaveAsで保存した複合MERGEFIELD文書は検証して開けます()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(ComplexMergeFieldPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.MergeFields["CustomerName"].Text = "変更後";
                document.SaveAs(outputPath);
            }

            var action = () =>
            {
                using var saved = Document.Open(outputPath, true);
            };

            action.Should().NotThrow();
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SaveAsで保存したContentControl文書は検証して開けます()
    {
        var sourcePath = TestDocument.CreateTemporaryCopy(ContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(sourcePath))
            {
                document.ContentControls["CustomerName"].Text = "変更後";
                document.SaveAs(outputPath);
            }

            var action = () =>
            {
                using var saved = Document.Open(outputPath, true);
            };

            action.Should().NotThrow();
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SaveAsで保存した日付選択ContentControl文書は検証して開けます()
    {
        var sourcePath =
            TestDocument.CreateTemporaryCopy(DatePickerContentControlPath);
        var outputPath = TestDocument.CreateOutputPath();
        var value = new DateTimeOffset(
            2027,
            1,
            2,
            0,
            0,
            0,
            TimeSpan.FromHours(9));
        try
        {
            using (var document = Document.Open(sourcePath, true))
            {
                document.DatePickers["DeliveryDate"].SelectedDateTime = value;
                document.SaveAs(outputPath);
            }

            using var saved = Document.Open(outputPath, true);
            var actual = saved.DatePickers["DeliveryDate"].SelectedDateTime;

            actual.Should().Be(value);
            actual.Offset.Should().Be(value.Offset);
        }
        finally
        {
            File.Delete(sourcePath);
            File.Delete(outputPath);
        }
    }

    public sealed class DocumentData
    {
        public string CustomerName { get; set; } = "";

        public string Address { get; set; } = "";
    }

    public sealed class CustomerNameOnlyDocumentData
    {
        public string CustomerName { get; set; } = "";
    }

    public sealed class AttributedDocumentData
    {
        [DocumentItemName("CustomerName")]
        public string Name { get; set; } = "";

        [DocumentItemName("Address")]
        public string Location { get; set; } = "";
    }

    public sealed class MissingDocumentItemData
    {
        [DocumentItemName("Missing")]
        public string Value { get; set; } = "";
    }

    public sealed class DuplicateDocumentItemData
    {
        [DocumentItemName("CustomerName")]
        public string FirstValue { get; set; } = "";

        [DocumentItemName("CustomerName")]
        public string SecondValue { get; set; } = "";
    }

    public sealed class AttributedPropertyWithoutPublicGetterData
    {
        [DocumentItemName("CustomerName")]
        public string Value
        {
            set { }
        }
    }

    public sealed class DocumentDataWithWriteOnlyProperty
    {
        public string CustomerName { get; set; } = "";

        public string Ignored
        {
            set { }
        }
    }

    public sealed class UnsupportedPropertyTypeData
    {
        [DocumentItemName("CustomerName")]
        public DateTime Value { get; set; }
    }

    public sealed class DateDocumentData
    {
        public DateTimeOffset DeliveryDate { get; set; }
    }

    public sealed class MissingDateDocumentData
    {
        [DocumentItemName("Missing")]
        public DateTimeOffset DeliveryDate { get; set; }
    }

}
