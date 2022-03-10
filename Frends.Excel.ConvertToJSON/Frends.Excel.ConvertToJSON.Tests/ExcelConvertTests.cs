using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using Frends.Excel.ConvertToJSON.Definitions;
using NUnit.Framework;

namespace Frends.Excel.ConvertToJSON.Tests;

[TestFixture]
public class ExcelConvertTests
{
    private readonly Input _input = new();
    private readonly Options _options = new();
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
        _input.Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"../../../../../TestData/");
        _options.ReadOnlyWorkSheetWithName = "";
    }

    [Test]
    public void TestConvertXlsxToJSON()
    {
        // Test converting all worksheets of xlsx file to JSON.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
        var result = Excel.ConvertToJSON(_input, _options, new CancellationToken());
        var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput1.xlsx"",""worksheets"":[{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""Foo""},{""B"":""Bar""},{""C"":""Kanji 働""},{""D"":""Summa""}]},{""2"":[{""A"":""1""},{""B"":""2""},{""C"":""3""},{""D"":""6""}]}]},{""name"":""OmituinenNimi"",""rows"":[{""1"":[{""A"":""Kissa kuva""},{""B"":""1""},{""C"":""2""},{""D"":""3""}]},{""15"":[{""A"":""Foo""}]},{""16"":[{""B"":""Bar""}]}]}]}}";
        Assert.That(Regex.Replace(result.JSON, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsToJSON()
    {
        // Test converting all worksheets of xls file to JSON.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
        var result = Excel.ConvertToJSON(_input, _options, new CancellationToken());
        var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput2.xls"",""worksheets"":[{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""Foo""},{""B"":""Bar""},{""C"":""Kanji 働""},{""D"":""Summa""}]},{""2"":[{""A"":""1""},{""B"":""2""},{""C"":""3""},{""D"":""6""}]}]},{""name"":""OmituinenNimi"",""rows"":[{""1"":[{""A"":""Kissa kuva""},{""B"":""1""},{""C"":""2""},{""D"":""3""}]},{""15"":[{""A"":""Foo""}]},{""16"":[{""B"":""Bar""}]}]}]}}";
        Assert.That(Regex.Replace(result.JSON, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }
    [Test]
    public void TestConvertXlsxOneWorksheetToJSON()
    {
        // Test converting one worksheet of xlsx file to JSON.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
        _options.ReadOnlyWorkSheetWithName = "Sheet1";
        var result = Excel.ConvertToJSON(_input, _options, new CancellationToken());
        var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput1.xlsx"",""worksheet"":{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""Foo""},{""B"":""Bar""},{""C"":""Kanji 働""},{""D"":""Summa""}]},{""2"":[{""A"":""1""},{""B"":""2""},{""C"":""3""},{""D"":""6""}]}]}}}";
        Assert.That(Regex.Replace(result.JSON, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsOneWorksheetToJSON()
    {
        // Test converting one worksheet of xls file to JSON.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
        _options.ReadOnlyWorkSheetWithName = "Sheet1";
        var result = Excel.ConvertToJSON(_input, _options, new CancellationToken());
        var expectedResult = @"{""workbook"":{""workbook_name"":""ExcelTestInput2.xls"",""worksheet"":{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""Foo""},{""B"":""Bar""},{""C"":""Kanji 働""},{""D"":""Summa""}]},{""2"":[{""A"":""1""},{""B"":""2""},{""C"":""3""},{""D"":""6""}]}]}}}";
        Assert.That(Regex.Replace(result.JSON, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsxOneWorkSheetWithDatesDDMMYYYYToJSON()
    {
        // Test converting worksheet with dates into dd/MM/yyyy format
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xlsx");
        _options.ReadOnlyWorkSheetWithName = "Sheet2";
        _options.DateFormat = DateFormats.DDMMYYYY;
        _options.ShortDatePattern = false;
        var result = Excel.ConvertToJSON(_input, _options, new CancellationToken());
        var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xlsx"",""worksheet"":{""name"":""Sheet2"",""rows"":[{""1"":[{""A"":""25/12/2021 0:00:00""},{""B"":""25/02/2021 12:45:41""}, {""C"":""12/05/2020 0:00:00""},{""D"":""30/12/2021 0:00:00""}]}]}}}";
        Assert.That(Regex.Replace(result.JSON, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestQuotesForJSON()
    {
        var input = new Input { Path = Path.Combine(excelFilesDir, "TestQuotes.xlsx") };
        var options = new Options { ThrowErrorOnFailure = true };
        var result = Excel.ConvertToJSON(input, options, new CancellationToken());

        var expectedResult = @"{""workbook"":{""workbook_name"":""TestQuotes.xlsx"",""worksheets"":[{""name"":""Quote \"" again"",""rows"":[{ ""1"":[{ ""A"":""Hello \"" quote""}]}]}]}}";
        Assert.That(Regex.Replace(result.JSON, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsxOneWorkSheetWithDatesMMDDYYYYToJSON()
    {
        // Test converting worksheet with dates into MM/dd/yyyy format
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xls");
        _options.ReadOnlyWorkSheetWithName = "Sheet1";
        _options.DateFormat = DateFormats.MMDDYYYY;
        _options.ShortDatePattern = false;
        var result = Excel.ConvertToJSON(_input, _options, new CancellationToken());
        var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xls"",""worksheet"":{""name"":""Sheet1"",""rows"":[{""1"":[{""A"":""1""},{""B"":""2""}, {""C"":""3""},{""D"":""4""}]}, {""2"":[{""A"":""12/12/2021 12:00:00AM""},{""B"":""02/25/2021 12:45:41PM""}, {""C"":""05/12/2020 12:00:00AM""},{""D"":""12/12/2021 12:00:00AM""}]}]}}}";
        Assert.That(Regex.Replace(result.JSON, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsxOneWorkSheetWithDatesYYYYMDDToJSON()
    {
        // Test converting worksheet with dates into MM/dd/yyyy format
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xlsx");
        _options.ReadOnlyWorkSheetWithName = "Sheet2";
        _options.DateFormat = DateFormats.YYYYMMDD;
        _options.ShortDatePattern = false;
        var result = Excel.ConvertToJSON(_input, _options, new CancellationToken());
        var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xlsx"",""worksheet"":{""name"":""Sheet2"",""rows"":[{""1"":[{""A"":""2021/12/25 0:00:00""},{""B"":""2021/02/25 12:45:41""}, {""C"":""2020/05/12 0:00:00""},{""D"":""2021/12/30 0:00:00""}]}]}}}";
        Assert.That(Regex.Replace(result.JSON, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsxOneWorkSheetWithDatesDDMMYYYYWithShortPatternToJSON()
    {
        // Test converting worksheet with dates into dd/MM/yyyy format with ShortTimePattern enabled
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xls");
        _options.ReadOnlyWorkSheetWithName = "Sheet2";
        _options.DateFormat = DateFormats.DDMMYYYY;
        _options.ShortDatePattern = true;
        var result = Excel.ConvertToJSON(_input, _options, new CancellationToken());
        var expectedResult = @"{""workbook"":{""workbook_name"":""TestDateFormat.xls"",""worksheet"":{""name"":""Sheet2"",""rows"":[{""1"":[{""A"":""25/12/2021""},{""B"":""25/02/2021""}, {""C"":""12/05/2020""},{""D"":""30/12/2021""}]}]}}}";
        Assert.That(Regex.Replace(result.JSON, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }
}
