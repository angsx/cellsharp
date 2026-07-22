using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelSchemaTests
{
    private static readonly string[] ConfiguredHeaders = ["Customer Name", "Customer ID", "Active"];

    [Fact]
    public void WriteUsesCustomHeadersAndSchemaColumnOrder()
    {
        var path = TemporaryPath();
        var schema = CustomerSchema();

        try
        {
            Excel.Write(path, new[] { new SchemaCustomer { Id = 7, Name = "Ada", Active = true } }, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var headers = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .GetFirstChild<SheetData>()!
                .Elements<Row>()
                .First()
                .Elements<Cell>()
                .Select(Text)
                .ToArray();

            Assert.Equal(ConfiguredHeaders, headers);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteExcludesIgnoredAndUndeclaredProperties()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, new[] { new SchemaCustomer { Id = 7, Name = "Ada", Active = true, InternalCode = "private" } }, CustomerSchema());

            using var document = SpreadsheetDocument.Open(path, false);
            var headers = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .GetFirstChild<SheetData>()!
                .Elements<Row>()
                .First()
                .Elements<Cell>()
                .Select(Text)
                .ToArray();

            Assert.DoesNotContain("InternalCode", headers);
            Assert.DoesNotContain("private", headers);
            Assert.DoesNotContain("Score", headers);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteKeepsFormulaLikeTextAsTextWithSchema()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, new[] { new SchemaCustomer { Id = 7, Name = "=SUM(A1:A2)", Active = true } }, CustomerSchema());

            using var document = SpreadsheetDocument.Open(path, false);
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .GetFirstChild<SheetData>()!
                .Elements<Row>()
                .Skip(1)
                .Single()
                .Elements<Cell>()
                .First();

            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            Assert.Equal("=SUM(A1:A2)", Text(cell));
            Assert.Null(cell.CellFormula);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void SchemaCanBeReusedForMultipleReadAndWriteOperations()
    {
        var firstPath = TemporaryPath();
        var secondPath = TemporaryPath();
        var schema = CustomerSchema();

        try
        {
            Excel.Write(firstPath, new[] { new SchemaCustomer { Id = 1, Name = "Ada", Active = true } }, schema);
            Excel.Write(secondPath, new[] { new SchemaCustomer { Id = 2, Name = "Grace", Active = false } }, schema);

            Assert.Equal("Ada", Assert.Single(Excel.Read<SchemaCustomer>(firstPath, schema).Rows).Name);
            Assert.Equal("Grace", Assert.Single(Excel.Read<SchemaCustomer>(secondPath, schema).Rows).Name);
        }
        finally
        {
            Delete(firstPath);
            Delete(secondPath);
        }
    }

    [Fact]
    public void ReadMapsCustomHeadersCaseInsensitivelyInAnyFileOrder()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["ACTIVE", "customer id", "CUSTOMER NAME"], ["true", "7", "Ada"]);

            var result = Excel.Read<SchemaCustomer>(path, CustomerSchema());

            var customer = Assert.Single(result.Rows);
            Assert.True(result.IsValid);
            Assert.Equal(7, customer.Id);
            Assert.Equal("Ada", customer.Name);
            Assert.True(customer.Active);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadDoesNotRequireIgnoredProperties()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Customer ID", "Customer Name", "Active"], ["7", "Ada", "true"]);

            var result = Excel.Read<SchemaCustomer>(path, CustomerSchema());

            Assert.True(result.IsValid);
            Assert.Null(Assert.Single(result.Rows).InternalCode);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadAllowsMissingOptionalHeaders()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Customer ID", "Active"], ["7", "true"]);

            var result = Excel.Read<SchemaCustomer>(path, CustomerSchema());

            var customer = Assert.Single(result.Rows);
            Assert.True(result.IsValid);
            Assert.Null(customer.Name);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadReportsMissingRequiredHeaders()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Customer Name", "Active"], ["Ada", "true"]);

            var result = Excel.Read<SchemaCustomer>(path, CustomerSchema());

            Assert.Empty(result.Rows);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.MissingHeader, error.Code);
            Assert.Equal("Customer ID", error.Header);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadIgnoresExtraColumnsWithSchema()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Note", "Active", "Customer ID"], ["ignored", "false", "7"]);

            var result = Excel.Read<SchemaCustomer>(path, CustomerSchema());

            var customer = Assert.Single(result.Rows);
            Assert.True(result.IsValid);
            Assert.Equal(7, customer.Id);
            Assert.False(customer.Active);
            Assert.Null(customer.Name);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadKeepsOtherRowsWhenASchemaMappedValueIsInvalid()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Customer ID", "Active"], ["7", "true"], ["invalid", "false"]);

            var result = Excel.Read<SchemaCustomer>(path, CustomerSchema());

            Assert.Equal(7, Assert.Single(result.Rows).Id);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.InvalidValue, error.Code);
            Assert.Equal("Customer ID", error.Header);
            Assert.Equal("A3", error.CellReference);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void SchemaRejectsDuplicateProperties()
    {
        var builder = Excel.Schema<SchemaCustomer>()
            .Column(customer => customer.Id);

        var exception = Assert.Throws<ArgumentException>(() => builder.Column(customer => customer.Id));

        Assert.Contains("already configured", exception.Message);
    }

    [Fact]
    public void SchemaRejectsDuplicateHeaders()
    {
        var builder = Excel.Schema<SchemaCustomer>()
            .Column(customer => customer.Id, column => column.Header("Value"));

        var exception = Assert.Throws<ArgumentException>(() =>
            builder.Column(customer => customer.Name, column => column.Header("value")));

        Assert.Contains("already configured", exception.Message);
    }

    [Fact]
    public void SchemaRejectsNonPropertyExpressions()
    {
        var exception = Assert.Throws<ArgumentException>(() => Excel.Schema<SchemaCustomer>()
            .Column(customer => customer.Name!.ToUpperInvariant()));

        Assert.Contains("direct property access", exception.Message);
    }

    [Fact]
    public void SchemaRejectsUnsupportedProperties()
    {
        var exception = Assert.Throws<NotSupportedException>(() => Excel.Schema<UnsupportedCustomer>()
            .Column(customer => customer.Tags));

        Assert.Contains("unsupported type", exception.Message);
    }

    [Fact]
    public void SchemaCannotBeEmpty()
    {
        Assert.Throws<InvalidOperationException>(() => Excel.Schema<SchemaCustomer>().Build());
    }

    [Fact]
    public void SchemaWithOnlyIgnoredColumnsCannotBeBuilt()
    {
        Assert.Throws<InvalidOperationException>(() => Excel.Schema<SchemaCustomer>()
            .Column(customer => customer.InternalCode, column => column.Ignore())
            .Build());
    }

    [Fact]
    public void ReadKeepsRowsThatPassValidation()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Customer Age", "Email"], ["18", "alice@example.test"]);

            var result = Excel.Read<ValidationCustomer>(path, ValidationSchema());

            Assert.True(result.IsValid);
            Assert.Equal(18, Assert.Single(result.Rows).Age);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadReportsValidationFailuresAndKeepsOtherRows()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Customer Age", "Email"], ["17", "alice@example.test"], ["18", "bob@example.test"]);

            var result = Excel.Read<ValidationCustomer>(path, ValidationSchema());

            Assert.Equal(18, Assert.Single(result.Rows).Age);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.ValidationFailed, error.Code);
            Assert.Equal("Customer Age", error.Header);
            Assert.Equal("17", error.Value);
            Assert.Equal("Customer must be at least 18", error.Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadKeepsConversionAndValidationErrorsDistinct()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path, ["Customer Age", "Email"], ["invalid", "alice@example.test"]);

            var result = Excel.Read<ValidationCustomer>(path, ValidationSchema());

            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.InvalidValue, error.Code);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadReportsEveryFailedValidationRuleForAConvertedValue()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ValidationCustomer>()
            .Column(customer => customer.Age, column => column
                .Header("Customer Age")
                .Validate(age => age >= 18, "Customer must be at least 18")
                .Validate(age => age % 2 == 0, "Customer age must be even"))
            .Column(customer => customer.Email)
            .Build();

        try
        {
            CreateWorkbook(path, ["Customer Age", "Email"], ["17", "alice@example.test"]);

            var result = Excel.Read<ValidationCustomer>(path, schema);

            Assert.Empty(result.Rows);
            Assert.Equal(2, result.Errors.Count);
            Assert.All(result.Errors, error => Assert.Equal(ExcelReadErrorCode.ValidationFailed, error.Code));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadPassesNullToNullableValidators()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ValidationCustomer>()
            .Column(customer => customer.Age)
            .Column(customer => customer.Email, column => column.Validate(email => email is null, "Email must be blank"))
            .Build();

        try
        {
            CreateWorkbook(path, ["Age", "Email"], ["18", null]);

            var result = Excel.Read<ValidationCustomer>(path, schema);

            Assert.True(result.IsValid);
            Assert.Null(Assert.Single(result.Rows).Email);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadPropagatesValidatorExceptions()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ValidationCustomer>()
            .Column(customer => customer.Age, column => column.Validate(_ => throw new InvalidOperationException("validator failure"), "Unused"))
            .Column(customer => customer.Email)
            .Build();

        try
        {
            CreateWorkbook(path, ["Age", "Email"], ["18", "alice@example.test"]);

            var exception = Assert.Throws<InvalidOperationException>(() => Excel.Read<ValidationCustomer>(path, schema));

            Assert.Equal("validator failure", exception.Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ValidationSchemasCanBeReused()
    {
        var firstPath = TemporaryPath();
        var secondPath = TemporaryPath();
        var schema = ValidationSchema();

        try
        {
            CreateWorkbook(firstPath, ["Customer Age", "Email"], ["17", "alice@example.test"]);
            CreateWorkbook(secondPath, ["Customer Age", "Email"], ["18", "bob@example.test"]);

            Assert.False(Excel.Read<ValidationCustomer>(firstPath, schema).IsValid);
            Assert.True(Excel.Read<ValidationCustomer>(secondPath, schema).IsValid);
        }
        finally
        {
            Delete(firstPath);
            Delete(secondPath);
        }
    }

    private static ExcelSchema<SchemaCustomer> CustomerSchema() => Excel.Schema<SchemaCustomer>()
        .Column(customer => customer.Name, column => column.Header("Customer Name").Optional())
        .Column(customer => customer.Id, column => column.Header("Customer ID"))
        .Column(customer => customer.Active)
        .Column(customer => customer.InternalCode, column => column.Ignore())
        .Build();

    private static ExcelSchema<ValidationCustomer> ValidationSchema() => Excel.Schema<ValidationCustomer>()
        .Column(customer => customer.Age, column => column
            .Header("Customer Age")
            .Validate(age => age >= 18, "Customer must be at least 18"))
        .Column(customer => customer.Email)
        .Build();

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

    private static string Text(Cell cell) => cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class SchemaCustomer
    {
        public int Id { get; set; }

        public string? Name { get; set; }

        public string? InternalCode { get; set; }

        public bool Active { get; set; }

        public decimal? Score { get; set; }
    }

    private sealed class UnsupportedCustomer
    {
        public List<string>? Tags { get; set; }
    }

    private sealed class ValidationCustomer
    {
        public int Age { get; set; }

        public string? Email { get; set; }
    }
}
