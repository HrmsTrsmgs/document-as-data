using System.Diagnostics.CodeAnalysis;
using FluentAssertions;
using Marimo.DocumentAsData.Test.TestDocuments;

namespace Marimo.DocumentAsData.Test;

[Collection(nameof(CurrentDirectoryCollection))]
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
    public void Open中は他のStreamから元ファイルへ書き込めません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        try
        {
            using var document = Document.Open(filePath);
            var action = () =>
            {
                // 相手側は読み書き共有を許可し、Document側が書き込みを拒むことを確認します。
                // 削除禁止だけでは、Open後の外部更新をSaveで上書きする事故を防げません。
                using var writer = File.Open(filePath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
            };

            action.Should().Throw<IOException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Disposeはファイルの束縛を解除します()
    {
        var copyPath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        var tested = Document.Open(copyPath);

        tested.Should().NotBeNull();
        tested.Dispose();
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
    public void Openは読み取り専用のStream上の文書を開きます()
    {
        using var stream = new MemoryStream(File.ReadAllBytes(SimpleMergeFieldsPath), writable: false);
        using var tested = Document.Open(stream);

        tested.MergeFields["CustomerName"].Text.Should().Be("株式会社○○");
    }

    [Fact]
    public void OpenはシークできないStream上の文書を開きます()
    {
        using var stream = new NonSeekableReadStream(File.OpenRead(SimpleMergeFieldsPath));
        using var tested = Document.Open(stream);

        tested.MergeFields["CustomerName"].Text.Should().Be("株式会社○○");
    }

    [Theory]
    [InlineData(1)]
    [InlineData(17)]
    public void OpenはStreamの現在位置を文書の先頭として読み込みます(int prefixLength)
    {
        using var stream = new MemoryStream(
            [.. new byte[prefixLength], .. File.ReadAllBytes(SimpleMergeFieldsPath)]);
        stream.Position = prefixLength;
        using var tested = Document.Open(stream);

        tested.MergeFields["CustomerName"].Text.Should().Be("株式会社○○");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Openは空のStreamで失敗しても呼び出し側のStreamを閉じません(bool seekable)
    {
        using var source = new MemoryStream();
        using Stream stream = seekable ? source : new NonSeekableReadStream(source);
        var tested = () =>
        {
            using var document = Document.Open(stream);
        };

        tested.Should().Throw<InvalidDataException>();
        stream.CanRead.Should().BeTrue();
        source.ToArray().Should().BeEmpty();
    }

    [Fact]
    public void Openは内容が0バイトのファイルを文書として開きません()
    {
        // 空白のページを持つ正常なDOCXではなく、ZIPパッケージ自体がない0バイトの固定データです。
        var action = () =>
        {
            using var document = Document.Open(@"TestData\内容が0バイトの不正文書.docx");
        };

        action.Should().Throw<InvalidDataException>();
    }

    [Fact]
    public void Openは0バイトのファイルで失敗してもファイルを束縛しません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(@"TestData\内容が0バイトの不正文書.docx");
        try
        {
            // Documentを取得できない失敗でも、利用側が元ファイルを削除できることを確認します。
            FluentActions.Invoking(
                () => Document.Open(filePath)
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
    public void Openは読み取り不可のStreamを内容を変更せず閉じずに拒否します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        var original = File.ReadAllBytes(filePath);
        try
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Write))
            {
                var tested = () =>
                {
                    using var document = Document.Open(stream);
                };

                tested.Should().Throw<Exception>();
                stream.CanWrite.Should().BeTrue();
            }

            File.ReadAllBytes(filePath).Should().Equal(original);
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Openは開閉でサイズが増える文書も固定容量のMemoryStreamで開いて閉じられます()
    {
        // byte[]を渡すMemoryStreamは拡張できません。共通補助の拡張可能Streamとは異なる条件です。
        // 空の文書.docxは正常なOOXMLですが、SDK 3.5.1で開閉するとZIPが918から924バイトへ増えます。
        // XMLの内容が同じでも再圧縮・再保存すると再現しなくなるため、原本のバイト列を保ちます。
        using var stream = new MemoryStream(File.ReadAllBytes(EmptyDocumentPath));

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

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Stream版のSaveは元のStreamを変更せずにNotSupportedExceptionを投げます(bool expandable)
    {
        var original = File.ReadAllBytes(SimpleMergeFieldsPath);
        using var stream = expandable
            ? TestDocument.CreateMemoryStream(SimpleMergeFieldsPath)
            : new MemoryStream([.. original]);

        using (var document = Document.Open(stream))
        {
            document.MergeFields["CustomerName"].Text = "変更後";
            var action = () => document.Save();

            // 容量不足ではなく、Stream版では元へのSaveを提供しないという制約です。
            // 拡張可能なStreamでも同じように拒否します。
            action.Should().Throw<NotSupportedException>();
            stream.ToArray().Should().Equal(original);
        }

        // 拒否した保存がDispose時に実行されることもありません。
        stream.ToArray().Should().Equal(original);
    }

    [Fact]
    public void DisposeはSaveしていない変更を元のStreamへ書き込みません()
    {
        var original = File.ReadAllBytes(SimpleMergeFieldsPath);
        using var stream = TestDocument.CreateMemoryStream(SimpleMergeFieldsPath);

        using (var document = Document.Open(stream))
        {
            document.MergeFields["CustomerName"].Text = "保存しない変更";
        }

        // 値だけでなく、SDKによるZIPの書き直しも元Streamへ漏らさないことを確認します。
        stream.ToArray().Should().Equal(original);
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
    public void ReadはContentControlから構造体のプロパティへ読み込みます()
    {
        using var document = Document.Open(ContentControlsPath);

        document.Read<StructDocumentData>()
            .Should().BeEquivalentTo(
                new StructDocumentData
                {
                    CustomerName = "山田太郎",
                    Address = "東京都"
                });
    }

    [Fact]
    public void Readはstaticプロパティを変更せずインスタンスのプロパティへ読み込みます()
    {
        using var document = Document.Open(ContentControlsPath);
        DocumentDataWithStaticProperty.CustomerName = "共有値";
        try
        {
            document.Read<DocumentDataWithStaticProperty>()
                .Address.Should().Be("東京都");
            DocumentDataWithStaticProperty.CustomerName.Should().Be("共有値");
        }
        finally
        {
            DocumentDataWithStaticProperty.CustomerName = "";
        }
    }

    [Fact]
    public void Readは自動変換したプロパティ名と一致するMERGEFIELDからオブジェクトを読み込みます()
    {
        using var document = Document.Open(@"TestData\customerData.docx");

        // MERGEFIELD名はcustomerNameです。対応属性なしのCustomerNameへ読み込みます。
        document.Read<CustomerNameOnlyDocumentData>()
            .CustomerName.Should().Be("株式会社○○");
    }

    [Fact]
    public void Readは自動変換したプロパティ名と一致するCheckBoxのTagからオブジェクトを読み込みます()
    {
        using var document = Document.Open(@"TestData\acceptanceForm.docx");

        // チェック済みのTagはtermsAcceptedです。対応属性なしのTermsAcceptedへ読み込みます。
        document.Read<AcceptanceDocumentData>()
            .TermsAccepted.Should().BeTrue();
    }

    [Fact]
    public void Readは自動変換したプロパティ名と一致する日付選択ContentControlのTagからオブジェクトを読み込みます()
    {
        using var document = Document.Open(@"TestData\deliveryForm.docx");

        // 文書のTagはdeliveryDateです。対応属性なしのDeliveryDateへ読み込みます。
        document.Read<DateDocumentData>().DeliveryDate
            .Should().Be(new DateTimeOffset(2026, 9, 4, 0, 0, 0, TimeSpan.Zero));
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
    public void Readはobject型のプロパティへの読み込みに失敗します()
    {
        using var document = Document.Open(ContentControlsPath);

        var action = () => document.Read<ObjectPropertyDocumentData>();

        action.Should().Throw<DocumentMappingException>();
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
    public void Replaceは自動変換したプロパティ名と一致するTagへオブジェクトを書き込みます()
    {
        using var document = Document.Open(@"TestData\customerForm.docx");

        // 対応属性なしのCustomerNameから、文書のTag customerNameへ書き込みます。
        document.Replace(
            new CustomerNameOnlyDocumentData
            {
                CustomerName = "変更後の氏名"
            });

        document.ContentControls["customerName"].Text.Should().Be("変更後の氏名");
    }

    [Fact]
    public void Replaceはstaticプロパティを使わずインスタンスのプロパティだけを文書へ書き込みます()
    {
        using var document = Document.Open(ContentControlsPath);
        DocumentDataWithStaticProperty.CustomerName = "共有値";
        try
        {
            document.Replace(new DocumentDataWithStaticProperty { Address = "大阪府" });

            document.ContentControls["Address"].Text.Should().Be("大阪府");
            document.ContentControls["CustomerName"].Text.Should().Be("山田太郎");
        }
        finally
        {
            DocumentDataWithStaticProperty.CustomerName = "";
        }
    }

    [Fact]
    public void Replaceは自動変換したプロパティ名と一致するCheckBoxのTagへオブジェクトを書き込みます()
    {
        using var document = Document.Open(@"TestData\acceptanceForm.docx");

        // 対応属性なしのTermsAcceptedから、Tag termsAcceptedのチェック済み項目を解除します。
        document.Replace(
            new AcceptanceDocumentData
            {
                TermsAccepted = false
            });

        document.CheckBoxes["termsAccepted"].IsChecked.Should().BeFalse();
    }

    [Fact]
    public void Replaceは自動変換したプロパティ名と一致する日付選択ContentControlのTagへオブジェクトを書き込みます()
    {
        using var document = Document.Open(@"TestData\deliveryForm.docx");

        // 対応属性なしのDeliveryDateから、Tag deliveryDateの日付を変更します。
        document.Replace(
            new DateDocumentData
            {
                DeliveryDate = new DateTimeOffset(2026, 10, 1, 0, 0, 0, TimeSpan.Zero)
            });

        document.DatePickers["deliveryDate"].SelectedDateTime
            .Should().Be(new DateTimeOffset(2026, 10, 1, 0, 0, 0, TimeSpan.Zero));
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
    public void Replaceは自動変換したプロパティ名と一致するMERGEFIELDへオブジェクトを書き込みます()
    {
        using var document = Document.Open(@"TestData\customerData.docx");

        // 対応属性なしのCustomerNameから、MERGEFIELD名customerNameへ書き込みます。
        document.Replace(
            new CustomerNameOnlyDocumentData
            {
                CustomerName = "変更後の氏名"
            });

        document.MergeFields["customerName"].Text.Should().Be("変更後の氏名");
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
    public void Saveは開いているファイルへ変更を保存します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        try
        {
            using (var document = Document.Open(filePath))
            {
                document.MergeFields["CustomerName"].Text = "変更後";
                document.Save();
            }

            using var saved = Document.Open(filePath);
            saved.MergeFields["CustomerName"].Text.Should().Be("変更後");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Saveは作業ディレクトリが変わっても相対パスで開いた元ファイルへ保存します()
    {
        var originalDirectory = Environment.CurrentDirectory;
        var sourceDirectory = Directory.CreateTempSubdirectory("DocumentAsData-Source-");
        var otherDirectory = Directory.CreateTempSubdirectory("DocumentAsData-Other-");
        var sourcePath = Path.Combine(sourceDirectory.FullName, "文書.docx");
        var otherPath = Path.Combine(otherDirectory.FullName, "文書.docx");

        try
        {
            File.Copy(SimpleMergeFieldsPath, sourcePath);
            File.Copy(EmptyDocumentPath, otherPath);
            Environment.CurrentDirectory = sourceDirectory.FullName;

            using (var document = Document.Open("文書.docx"))
            {
                document.MergeFields["CustomerName"].Text = "変更後";
                Environment.CurrentDirectory = otherDirectory.FullName;

                document.Save();
            }

            Environment.CurrentDirectory = originalDirectory;
            using var saved = Document.Open(sourcePath);
            saved.MergeFields["CustomerName"].Text.Should().Be("変更後");
            // 新しい作業ディレクトリにある同名ファイルへ上書きしてはいけません。
            File.ReadAllBytes(otherPath).Should().Equal(File.ReadAllBytes(EmptyDocumentPath));
        }
        finally
        {
            Environment.CurrentDirectory = originalDirectory;
            File.Delete(sourcePath);
            File.Delete(otherPath);
            sourceDirectory.Delete();
            otherDirectory.Delete();
        }
    }

    [Fact]
    public void Save後も他のStreamから元ファイルへ書き込めません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        try
        {
            using var document = Document.Open(filePath);
            document.Save();

            var action = () =>
            {
                // Saveは保存中だけ元ファイルの束縛を解除するため、戻った後に書き込み禁止が復元されることを確認します。
                using var writer = File.Open(filePath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
            };

            action.Should().Throw<IOException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Saveに失敗した後も他のStreamから元ファイルへ書き込めません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        try
        {
            using var document = Document.Open(filePath);
            // 別の読み取り用Streamが書き込みを拒否することで、Saveを失敗させます。
            using (var reader = File.OpenRead(filePath))
            {
                FluentActions.Invoking(document.Save).Should().Throw<IOException>();
            }

            var action = () =>
            {
                // readerを閉じた後なので、ここでの書き込み禁止はDocument自身によるものです。
                using var writer = File.Open(filePath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite);
            };

            action.Should().Throw<IOException>();
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Saveが元ファイルへの書き込みに失敗しても編集内容を保持して再保存できます()
    {
        var filePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        try
        {
            using (var document = Document.Open(filePath))
            {
                document.MergeFields["CustomerName"].Text = "変更後";
                // 別の読み取り用Streamで保存を妨げます。原因を取り除いた後、同じDocumentで再試行します。
                using (var reader = File.OpenRead(filePath))
                {
                    FluentActions.Invoking(document.Save).Should().Throw<IOException>();
                }

                document.Save();
            }

            using var saved = Document.Open(filePath);
            saved.MergeFields["CustomerName"].Text.Should().Be("変更後");
        }
        finally
        {
            File.Delete(filePath);
        }
    }

    [Fact]
    public void Dispose後のSaveは例外を投げて元ファイルを再び束縛しません()
    {
        var filePath = TestDocument.CreateTemporaryCopy(EmptyDocumentPath);
        using var tested = Document.Open(filePath);
        tested.Dispose();

        FluentActions.Invoking(tested.Save)
            .Should().Throw<ObjectDisposedException>();

        FluentActions.Invoking(() => File.Delete(filePath))
            .Should().NotThrow();
    }

    [Fact]
    public void Saveは文書を閉じる前に元ファイルへ変更を反映します()
    {
        var filePath = TestDocument.CreateTemporaryCopy(SimpleMergeFieldsPath);
        try
        {
            using var document = Document.Open(filePath);
            document.MergeFields["CustomerName"].Text = "変更後";

            document.Save();

            // documentをDisposeする前に別の文書として読み直し、Saveの完了時点を確認します。
            using var saved = Document.Open(filePath);
            saved.MergeFields["CustomerName"].Text.Should().Be("変更後");
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
    public void SaveAsは固定容量の元Streamを変更せずに編集結果を別ファイルへ保存します()
    {
        var original = File.ReadAllBytes(SimpleMergeFieldsPath);
        // byte[]を渡して固定容量にします。比較用のoriginalとは別の配列を使います。
        using var stream = new MemoryStream(File.ReadAllBytes(SimpleMergeFieldsPath));
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(stream))
            {
                document.MergeFields["CustomerName"].Text = "保存先にだけ書き込む変更後の値";

                document.SaveAs(outputPath);

                // 元文書をDisposeする前でも、保存先は完成したDOCXとして読み取れます。
                using var saved = Document.Open(outputPath, true);
                saved.MergeFields["CustomerName"].Text.Should().Be("保存先にだけ書き込む変更後の値");
                stream.ToArray().Should().Equal(original);
            }

            stream.ToArray().Should().Equal(original);
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SaveAsは書き込み不可のStreamからも元データを変更せず編集結果を保存します(bool seekable)
    {
        var original = File.ReadAllBytes(SimpleMergeFieldsPath);
        using var source = new MemoryStream([.. original], writable: false);
        using Stream stream = seekable ? source : new NonSeekableReadStream(source);
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(stream))
            {
                document.MergeFields["CustomerName"].Text = "保存した文字列";
                document.SaveAs(outputPath);

                using var saved = Document.Open(outputPath, true);
                saved.MergeFields["CustomerName"].Text.Should().Be("保存した文字列");
                source.ToArray().Should().Equal(original);
            }

            source.ToArray().Should().Equal(original);
            stream.CanRead.Should().BeTrue();
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SaveAsは途中位置から開いたStreamの前置データと元文書を変更せず編集結果を保存します()
    {
        byte[] prefix = [1, 2, 3, 4];
        byte[] original = [.. prefix, .. File.ReadAllBytes(SimpleMergeFieldsPath)];
        using var stream = new MemoryStream([.. original]);
        stream.Position = prefix.Length;
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(stream))
            {
                document.MergeFields["CustomerName"].Text = "保存した文字列";
                document.SaveAs(outputPath);

                using var saved = Document.Open(outputPath, true);
                saved.MergeFields["CustomerName"].Text.Should().Be("保存した文字列");
                stream.ToArray().Should().Equal(original);
            }

            stream.ToArray().Should().Equal(original);
            stream.CanRead.Should().BeTrue();
        }
        finally
        {
            File.Delete(outputPath);
        }
    }

    [Fact]
    public void SaveAs後に編集して再びSaveAsすると各保存時点の内容を別々に保持します()
    {
        var original = File.ReadAllBytes(SimpleMergeFieldsPath);
        using var stream = new MemoryStream([.. original]);
        var firstPath = TestDocument.CreateOutputPath();
        var secondPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(stream))
            {
                document.MergeFields["CustomerName"].Text = "一回目";
                document.SaveAs(firstPath);
                document.MergeFields["CustomerName"].Text = "二回目";
                document.SaveAs(secondPath);
                document.MergeFields["CustomerName"].Text = "保存しない変更";
            }

            using var first = Document.Open(firstPath, true);
            using var second = Document.Open(secondPath, true);
            first.MergeFields["CustomerName"].Text.Should().Be("一回目");
            second.MergeFields["CustomerName"].Text.Should().Be("二回目");
            stream.ToArray().Should().Equal(original);
        }
        finally
        {
            File.Delete(firstPath);
            File.Delete(secondPath);
        }
    }

    [Fact]
    public void SaveAsが保存先を開けず失敗しても元Streamと編集内容を保持して再保存できます()
    {
        var original = File.ReadAllBytes(SimpleMergeFieldsPath);
        using var stream = new MemoryStream([.. original]);
        var blockedPath = TestDocument.CreateOutputPath();
        var outputPath = TestDocument.CreateOutputPath();
        try
        {
            using (var document = Document.Open(stream))
            {
                document.MergeFields["CustomerName"].Text = "保存した文字列";

                using (var blockedFile = File.Create(blockedPath))
                {
                    var tested = () => document.SaveAs(blockedPath);

                    tested.Should().Throw<IOException>();
                }

                stream.CanRead.Should().BeTrue();
                stream.ToArray().Should().Equal(original);
                document.SaveAs(outputPath);

                using var saved = Document.Open(outputPath, true);
                saved.MergeFields["CustomerName"].Text.Should().Be("保存した文字列");
            }

            stream.ToArray().Should().Equal(original);
            stream.CanRead.Should().BeTrue();
        }
        finally
        {
            File.Delete(blockedPath);
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

    /// <summary>
    /// Readの戻り値が値型でも、複数のプロパティへ読み込んだ値が保持されることを確認します。
    /// </summary>
    public struct StructDocumentData
    {
        /// <summary>名前付き項目CustomerNameから読み込む文字列です。</summary>
        public string CustomerName { get; set; }

        /// <summary>名前付き項目Addressから読み込む文字列です。</summary>
        public string Address { get; set; }
    }

    public sealed class CustomerNameOnlyDocumentData
    {
        public string CustomerName { get; set; } = "";
    }

    /// <summary>
    /// ReadとReplaceがインスタンスの値だけを扱い、型全体の共有状態を対応付けないことを確認します。
    /// </summary>
    public sealed class DocumentDataWithStaticProperty
    {
        /// <summary>文書項目と同名でも読み込み対象にしない共有状態です。</summary>
        public static string CustomerName { get; set; } = "";

        /// <summary>通常の読み込み対象となるインスタンスの値です。</summary>
        public string Address { get; set; } = "";
    }

    public sealed class AcceptanceDocumentData
    {
        public bool TermsAccepted { get; set; }
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
        // staticの除外ではなく、属性で指定したインスタンスプロパティのgetter不足を検証します。
        [DocumentItemName("CustomerName")]
        [SuppressMessage("Performance", "CA1822", Justification = "getterのないインスタンスプロパティを検証するため、staticにはしません。")]
        public string Value
        {
            set { }
        }
    }

    public sealed class DocumentDataWithWriteOnlyProperty
    {
        public string CustomerName { get; set; } = "";

        // staticの除外ではなく、getterのないインスタンスプロパティを無視することを検証します。
        [SuppressMessage("Performance", "CA1822", Justification = "getterのないインスタンスプロパティを検証するため、staticにはしません。")]
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

    /// <summary>
    /// 文字列を代入できるobject型でも、対応型として暗黙に受け入れないことを確認します。
    /// </summary>
    public sealed class ObjectPropertyDocumentData
    {
        /// <summary>文書項目と同名ですが、型が読み込み対象外のプロパティです。</summary>
        public object CustomerName { get; set; } = "";
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
