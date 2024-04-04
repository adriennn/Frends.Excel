using System;
using System.IO;
using Frends.Excel.ConvertFromJSON.Definitions;
using NUnit.Framework;

namespace Frends.Excel.ConvertFromJSON.Tests;

[TestFixture]
public class ExcelConvertTests
{
    private static Input _input = new();
    private static Options _options = new();
    private readonly string excelFilesDir;


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
    public void TestConvertJSONToXlsx()
    {
        // Test converting all worksheets of xlsx file to JSON.
        _input.Path = Path.Combine(_input.Path, "JsonTestInput.json");
        var temp = Excel.ConvertFromJSON(_input, _options, default);
        // Write to excel, then read it back
        _input.Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"../../../../../TestData/");
        _input.Path = Path.Combine(_input.Path, temp);
        var result = Excel.ConvertToJSON(_input, _options, default);
        var expectedResult = @"...";
        Assert.AreEqual(expectedResult, result.JSON);
    }
}
