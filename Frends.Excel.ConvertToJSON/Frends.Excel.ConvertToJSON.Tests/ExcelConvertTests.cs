using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Text.RegularExpressions;
using Frends.Excel.ConvertToJSON.Definitions;
using NUnit.Framework;

namespace Frends.Excel.ConvertToJSON.Tests;

[TestFixture]
public class ExcelConvertTests
{
    private static Input _input = new();
    private static Options _options = new();
    private readonly string excelFilesDir;

    // Cat image in example files is from Pixbay.com. It is licenced in CC0 Public Domain (Free for commercial use, No attribution required).
    // It is uploaded by Ben_Kerckx https://pixabay.com/en/cat-animal-pet-cats-close-up-300572/

    public ExcelConvertTests()
    {
        excelFilesDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"../../../../../TestData/");
    }

    [SetUp]
    public void Setup()
    {
        _input = new();
        _input.Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"../../../../../TestData/");
        _options = new();
        _options.ReadOnlyWorkSheetWithName = "";
    }

    [Test]
    public void TestConvertXlsxToJSON()
    {
        // Test converting all worksheets of xlsx file to JSON.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
        var result = Excel.ConvertToJSON(_input, _options, default);
        var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput1.xlsx"",""worksheets"":[{""name"":""Sheet1"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""Foo""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""Bar""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""Kanji 働""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""Summa""}]},{""RowNumber"":2,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""1""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""2""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""3""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""6""}]}]},{""name"":""OmituinenNimi"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""Kissa kuva""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""1""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""2""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""3""}]},{""RowNumber"":15,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""Foo""}]},{""RowNumber"":16,""Cells"":[{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""Bar""}]}]}]}}";
        Assert.AreEqual(expectedResult, result.JSON);
    }

    [Test]
    public void TestConvertXlsToJSON()
    {
        // Test converting all worksheets of xls file to JSON.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
        var result = Excel.ConvertToJSON(_input, _options, default);
        var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput2.xls"",""worksheets"":[{""name"":""Sheet1"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""Foo""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""Bar""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""Kanji 働""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""Summa""}]},{""RowNumber"":2,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""1""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""2""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""3""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""6""}]}]},{""name"":""OmituinenNimi"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""Kissa kuva""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""1""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""2""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""3""}]},{""RowNumber"":15,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""Foo""}]},{""RowNumber"":16,""Cells"":[{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""Bar""}]}]}]}}";
        Assert.AreEqual(expectedResult, result.JSON);
    }

    [Test]
    public void TestConvertXlsxOneWorksheetToJSON()
    {
        // Test converting one worksheet of xlsx file to JSON.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
        _options.ReadOnlyWorkSheetWithName = "Sheet1";
        var result = Excel.ConvertToJSON(_input, _options, default);
        var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput1.xlsx"",""worksheet"":{""name"":""Sheet1"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""Foo""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""Bar""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""Kanji 働""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""Summa""}]},{""RowNumber"":2,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""1""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""2""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""3""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""6""}]}]}}}";
        Assert.AreEqual(expectedResult, result.JSON);
    }

    [Test]
    public void TestConvertXlsOneWorksheetToJSON()
    {
        // Test converting one worksheet of xls file to JSON.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
        _options.ReadOnlyWorkSheetWithName = "Sheet1";
        var result = Excel.ConvertToJSON(_input, _options, default);
        var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput2.xls"",""worksheet"":{""name"":""Sheet1"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""Foo""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""Bar""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""Kanji 働""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""Summa""}]},{""RowNumber"":2,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""1""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""2""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""3""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""6""}]}]}}}";
        Assert.AreEqual(expectedResult, result.JSON);
    }

    [Test]
    public void TestConvertXlsxOneWorkSheetWithDatesDDMMYYYYToJSON()
    {
        // Test converting worksheet with dates into dd/MM/yyyy format
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xlsx");
        _options.ReadOnlyWorkSheetWithName = "Sheet2";
        _options.DateFormat = DateFormats.DDMMYYYY;
        _options.ShortDatePattern = false;
        var result = Excel.ConvertToJSON(_input, _options, default);
        var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xlsx"",""worksheet"":{""name"":""Sheet2"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""25/12/2021 0:00:00""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""25/02/2021 12:45:41""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""12/05/2020 0:00:00""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""30/12/2021 0:00:00""}]}]}}}";
        Assert.AreEqual(expectedResult, result.JSON);
    }

    [Test]
    public void TestQuotesForJSON()
    {
        var input = new Input { Path = Path.Combine(excelFilesDir, "TestQuotes.xlsx") };
        var options = new Options { ThrowErrorOnFailure = true };
        var result = Excel.ConvertToJSON(input, options, default);
        var expectedResult = @"{""workbook"":{""workbook_name"":""TestQuotes.xlsx"",""worksheets"":[{""name"":""Quote \"" again"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""Hello \"" quote""}]}]}]}}";
        Assert.AreEqual(expectedResult, result.JSON);
    }

    [Test]
    public void TestConvertXlsxOneWorkSheetWithDatesMMDDYYYYToJSON()
    {
        // Test converting worksheet with dates into MM/dd/yyyy format
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xls");
        _options.ReadOnlyWorkSheetWithName = "Sheet1";
        _options.DateFormat = DateFormats.MMDDYYYY;
        _options.ShortDatePattern = false;
        var result = Excel.ConvertToJSON(_input, _options, default);
        var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xls"",""worksheet"":{""name"":""Sheet1"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""1""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""2""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""3""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""4""}]},{""RowNumber"":2,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""12/12/2021 12:00:00 AM""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""02/25/2021 12:45:41 PM""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""05/12/2020 12:00:00 AM""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""12/12/2021 12:00:00 AM""}]}]}}}";
        Assert.AreEqual(expectedResult, result.JSON);
    }

    [Test]
    public void TestConvertXlsxOneWorkSheetWithDatesYYYYMDDToJSON()
    {
        // Test converting worksheet with dates into MM/dd/yyyy format
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xlsx");
        _options.ReadOnlyWorkSheetWithName = "Sheet2";
        _options.DateFormat = DateFormats.YYYYMMDD;
        _options.ShortDatePattern = false;
        var result = Excel.ConvertToJSON(_input, _options, default);
        var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xlsx"",""worksheet"":{""name"":""Sheet2"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""2021/12/25 0:00:00""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""2021/02/25 12:45:41""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""2020/05/12 0:00:00""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""2021/12/30 0:00:00""}]}]}}}";
        Assert.AreEqual(expectedResult, result.JSON);
    }

    [Test]
    public void TestConvertXlsxOneWorkSheetWithDatesDDMMYYYYWithShortPatternToJSON()
    {
        // Test converting worksheet with dates into dd/MM/yyyy format with ShortTimePattern enabled
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xls");
        _options.ReadOnlyWorkSheetWithName = "Sheet2";
        _options.DateFormat = DateFormats.DDMMYYYY;
        _options.ShortDatePattern = true;
        var result = Excel.ConvertToJSON(_input, _options, default);
        var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xls"",""worksheet"":{""name"":""Sheet2"",""rows"":[{""RowNumber"":1,""Cells"":[{""ColumnName"":""A"",""ColumnIndex"":1,""ColumnValue"":""25/12/2021""},{""ColumnName"":""B"",""ColumnIndex"":2,""ColumnValue"":""25/02/2021""},{""ColumnName"":""C"",""ColumnIndex"":3,""ColumnValue"":""12/05/2020""},{""ColumnName"":""D"",""ColumnIndex"":4,""ColumnValue"":""30/12/2021""}]}]}}}";
        Assert.AreEqual(expectedResult, result.JSON);
    }
}
