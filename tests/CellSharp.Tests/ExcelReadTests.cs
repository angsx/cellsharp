using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelReadTests
{
    [Fact]
    public void ReadRoundTripsWrittenRows()
    {
        var path = TemporaryPath();
        var customer = new Customer
        {
            Id = 7,
            Name = "Ada",
            BirthDate = new DateTime(1990, 1, 2, 3, 4, 5),
            Active = true,
            ExternalId = Guid.Parse("8b7f9a19-6520-4f29-9bb8-50b763162e7d"),
            Balance = 10.25m,
        };

        try
        {
            Excel.Write(path, new[] { customer });

            var result = Excel.Read<Customer>(path);

            Assert.True(result.IsValid);
            Assert.Empty(result.Errors);
            var read = Assert.Single(result.Rows);
            Assert.Equal(customer.Id, read.Id);
            Assert.Equal(customer.Name, read.Name);
            Assert.Equal(customer.BirthDate, read.BirthDate);
            Assert.Equal(customer.Active, read.Active);
            Assert.Equal(customer.ExternalId, read.ExternalId);
            Assert.Equal(customer.Balance, read.Balance);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadReleasesTheSourceFile()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, [new SimpleCustomer { Id = 1, Name = "Ada", Active = true }]);

            var result = Excel.Read<SimpleCustomer>(path);

            Assert.True(result.IsValid);
            using var exclusiveAccess = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadMapsColumnsInAnyOrder()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Name", "Active", "Id"], ["Alice", "true", "1"]);

            var result = Excel.Read<SimpleCustomer>(path);

            var customer = Assert.Single(result.Rows);
            Assert.True(result.IsValid);
            Assert.Equal(1, customer.Id);
            Assert.Equal("Alice", customer.Name);
            Assert.True(customer.Active);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadIgnoresColumnsWithoutMatchingProperties()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", "Name", "Active", "Comment"], ["1", "Alice", "false", "ignored"]);

            var result = Excel.Read<SimpleCustomer>(path);

            var customer = Assert.Single(result.Rows);
            Assert.True(result.IsValid);
            Assert.Equal("Alice", customer.Name);
            Assert.False(customer.Active);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadIgnoresEmptyHeaderCells()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", null, "Name", "Active"], ["1", "unused", "Alice", "true"]);

            var result = Excel.Read<SimpleCustomer>(path);

            Assert.True(result.IsValid);
            Assert.Equal("Alice", Assert.Single(result.Rows).Name);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadMatchesHeadersWithoutCaseSensitivity()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["id", "NAME", "active"], ["1", "Alice", "TRUE"]);

            var result = Excel.Read<SimpleCustomer>(path);

            Assert.True(result.IsValid);
            Assert.Equal("Alice", Assert.Single(result.Rows).Name);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadAllowsNullStringCells()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", "Name", "Score"], ["1", null, "4.5"]);

            var result = Excel.Read<NullableValues>(path);

            var row = Assert.Single(result.Rows);
            Assert.True(result.IsValid);
            Assert.Null(row.Name);
            Assert.Equal(4.5m, row.Score);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadAllowsNullNullableNumbers()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", "Name", "Score"], ["1", "Alice", null]);

            var result = Excel.Read<NullableValues>(path);

            var row = Assert.Single(result.Rows);
            Assert.True(result.IsValid);
            Assert.Null(row.Score);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadUsesConfiguredCultureForTextNumbersAndCanTreatEmptyTextAsNull()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", "Name", "Score"], ["1", "Ada", "45,498933"], ["2", "Grace", string.Empty]);

            var defaultResult = Excel.Read<NullableValues>(path);
            var italianResult = Excel.Read<NullableValues>(path, new ExcelReadOptions(
                culture: CultureInfo.GetCultureInfo("it-IT"),
                emptyStringAsNull: true));

            Assert.Empty(defaultResult.Rows);
            Assert.Equal(2, defaultResult.Errors.Count);
            Assert.True(italianResult.IsValid);
            Assert.Equal([45.498933M, null], italianResult.Rows.Select(row => row.Score));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadConvertsCellSharpDateTimeCells()
    {
        var path = TemporaryPath();
        var birthDate = new DateTime(1990, 1, 2, 3, 4, 5);

        try
        {
            Excel.Write(path, new[] { new DateCustomer { Id = 1, Name = "Alice", BirthDate = birthDate, Active = true } });

            var result = Excel.Read<DateCustomer>(path);

            Assert.True(result.IsValid);
            Assert.Equal(birthDate, Assert.Single(result.Rows).BirthDate);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadConvertsGuidAndBooleanValues()
    {
        var path = TemporaryPath();
        var externalId = Guid.Parse("4bd2433f-9bcd-4b86-8aa3-6e31c1a45432");

        try
        {
            Excel.Write(path, new[] { new GuidCustomer { Id = 1, ExternalId = externalId, Active = false } });

            var result = Excel.Read<GuidCustomer>(path);

            var customer = Assert.Single(result.Rows);
            Assert.True(result.IsValid);
            Assert.Equal(externalId, customer.ExternalId);
            Assert.False(customer.Active);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadSkipsCompletelyEmptyRows()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", "Name", "Active"], [null, null, null], ["1", "Alice", "true"]);

            var result = Excel.Read<SimpleCustomer>(path);

            Assert.True(result.IsValid);
            Assert.Equal("Alice", Assert.Single(result.Rows).Name);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadReportsInvalidNumericValuesAndKeepsOtherRows()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", "Name", "Active"], ["1", "Alice", "true"], ["not-a-number", "Bob", "true"]);

            var result = Excel.Read<SimpleCustomer>(path);

            Assert.False(result.IsValid);
            Assert.Equal("Alice", Assert.Single(result.Rows).Name);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.InvalidValue, error.Code);
            Assert.Equal(3U, error.RowNumber);
            Assert.Equal(1, error.ColumnNumber);
            Assert.Equal("A3", error.CellReference);
            Assert.Equal("Id", error.Header);
            Assert.Equal("not-a-number", error.Value);
            Assert.Equal(typeof(int), error.ExpectedType);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadReportsInvalidDateTimeValues()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", "Name", "BirthDate", "Active"], ["1", "Alice", "1990-01-01", "true"], ["2", "Bob", "invalid", "true"]);

            var result = Excel.Read<DateCustomer>(path);

            Assert.Equal("Alice", Assert.Single(result.Rows).Name);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.InvalidValue, error.Code);
            Assert.Equal(3U, error.RowNumber);
            Assert.Equal(3, error.ColumnNumber);
            Assert.Equal("C3", error.CellReference);
            Assert.Equal("BirthDate", error.Header);
            Assert.Equal("invalid", error.Value);
            Assert.Equal(typeof(DateTime), error.ExpectedType);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadReportsInvalidGuidValues()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", "ExternalId", "Active"], ["1", "not-a-guid", "true"]);

            var result = Excel.Read<GuidCustomer>(path);

            var error = Assert.Single(result.Errors);
            Assert.Empty(result.Rows);
            Assert.Equal(ExcelReadErrorCode.InvalidValue, error.Code);
            Assert.Equal("ExternalId", error.Header);
            Assert.Equal(typeof(Guid), error.ExpectedType);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadReportsDuplicateHeadersWithoutProducingRows()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id", "ID"], ["1", "2"]);

            var result = Excel.Read<SingleId>(path);

            var error = Assert.Single(result.Errors);
            Assert.Empty(result.Rows);
            Assert.Equal(ExcelReadErrorCode.DuplicateHeader, error.Code);
            Assert.Equal(1U, error.RowNumber);
            Assert.Equal(2, error.ColumnNumber);
            Assert.Equal("B1", error.CellReference);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadReportsMissingHeadersWithoutProducingRows()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id"], ["1"]);

            var result = Excel.Read<SimpleCustomer>(path);

            Assert.Empty(result.Rows);
            Assert.Contains(result.Errors, error => error.Code == ExcelReadErrorCode.MissingHeader && error.Header == "Name");
            Assert.Contains(result.Errors, error => error.Code == ExcelReadErrorCode.MissingHeader && error.Header == "Active");
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadThrowsForMissingFiles()
    {
        var path = Path.Combine(Path.GetTempPath(), $"cellsharp-{Guid.NewGuid():N}.xlsx");

        Assert.Throws<FileNotFoundException>(() => Excel.Read<SimpleCustomer>(path));
    }

    [Fact]
    public void ReadThrowsForTypesWithoutPublicParameterlessConstructors()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Id"], ["1"]);

            var exception = Assert.Throws<InvalidOperationException>(() => Excel.Read<NoDefaultConstructor>(path));

            Assert.Contains("public parameterless constructor", exception.Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadKeepsFormulaLikeTextAsText()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Value"], ["=SUM(A1:A2)"]);

            var result = Excel.Read<FormulaCandidate>(path);

            Assert.True(result.IsValid);
            Assert.Equal("=SUM(A1:A2)", Assert.Single(result.Rows).Value);
        }
        finally
        {
            Delete(path);
        }
    }

    private static void CreateWorkbook(string path, params string?[][] values)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        for (var rowIndex = 0; rowIndex < values.Length; rowIndex++)
        {
            var rowNumber = (uint)(rowIndex + 1);
            var row = new Row { RowIndex = rowNumber };
            for (var columnIndex = 0; columnIndex < values[rowIndex].Length; columnIndex++)
            {
                var value = values[rowIndex][columnIndex];
                if (value is not null)
                {
                    row.AppendChild(new Cell
                    {
                        CellReference = ColumnName(columnIndex + 1) + rowNumber,
                        DataType = CellValues.InlineString,
                        InlineString = new InlineString(new Text(value) { Space = SpaceProcessingModeValues.Preserve }),
                    });
                }
            }

            sheetData.AppendChild(row);
        }

        var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1U,
            Name = "Sheet1",
        });
        worksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }

    private static string ColumnName(int columnNumber)
    {
        var column = string.Empty;
        var value = columnNumber;
        while (value > 0)
        {
            value--;
            column = (char)('A' + (value % 26)) + column;
            value /= 26;
        }

        return column;
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class Customer
    {
        public int Id { get; set; }

        public string? Name { get; set; }

        public DateTime BirthDate { get; set; }

        public bool Active { get; set; }

        public Guid ExternalId { get; set; }

        public decimal? Balance { get; set; }
    }

    private sealed class SimpleCustomer
    {
        public int Id { get; set; }

        public string? Name { get; set; }

        public bool Active { get; set; }
    }

    private sealed class NullableValues
    {
        public int Id { get; set; }

        public string? Name { get; set; }

        public decimal? Score { get; set; }
    }

    private sealed class DateCustomer
    {
        public int Id { get; set; }

        public string? Name { get; set; }

        public DateTime BirthDate { get; set; }

        public bool Active { get; set; }
    }

    private sealed class GuidCustomer
    {
        public int Id { get; set; }

        public Guid ExternalId { get; set; }

        public bool Active { get; set; }
    }

    private sealed class SingleId
    {
        public int Id { get; set; }
    }

    private sealed class NoDefaultConstructor
    {
        public NoDefaultConstructor(int id)
        {
            Id = id;
        }

        public int Id { get; set; }
    }

    private sealed class FormulaCandidate
    {
        public string? Value { get; set; }
    }
}
