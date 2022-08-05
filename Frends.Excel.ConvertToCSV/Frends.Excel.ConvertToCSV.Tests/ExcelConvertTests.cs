using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using Frends.Excel.ConvertToCSV.Definitions;
using NUnit.Framework;

namespace Frends.Excel.ConvertToCSV.Tests;

[TestFixture]
public class ExcelConvertTests
{
    private readonly Input _input = new();
    private readonly Options _options = new();

    // Cat image in example files is from Pixbay.com. It is licenced in CC0 Public Domain (Free for commercial use, No attribution required). It is uploaded by Ben_Kerckx https://pixabay.com/en/cat-animal-pet-cats-close-up-300572/

    [SetUp]
    public void Setup()
    {
        _input.Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"../../../../../TestData/");
        _options.CsvSeparator = ",";
        _options.ReadOnlyWorkSheetWithName = "";
    }

    /// <summary>
    /// Test converting all worksheets of xlsx file to csv.
    /// </summary>
    [Test]
    public void TestConvertXlsxToCSV()
    {
        if (_input.Path != null)
        {
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
            var result = Excel.ConvertToCSV(_input, _options, new CancellationToken());
            var expectedResult =
                "Foo,Bar,Kanji 働,Summa\n1,2,3,6\nKissa kuva,1,2,3\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\nFoo,,,\n,Bar,,\n";
            if (result.CSV != null)
                Assert.That(Regex.Replace(result.CSV, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
            else
                Assert.Fail();
        }
        else
            throw new Exception($"Path missing, {_input.Path}");
    }

    /// <summary>
    /// Test converting all worksheets of xls file to csv.
    /// </summary>
    [Test]
    public void TestConvertXlsToCSV()
    {
        if (_input.Path != null)
        {
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = Excel.ConvertToCSV(_input, _options, new CancellationToken());
            var expectedResult =
                "Foo,Bar,Kanji 働,Summa\n1,2,3,6\nKissa kuva,1,2,3\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\n,,,\nFoo,,,\n,Bar,,\n";
            if (result.CSV != null)
                Assert.That(Regex.Replace(result.CSV, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
            else
                Assert.Fail();
        }
        else
            throw new Exception($"Path missing, {_input.Path}");
    }

    /// <summary>
    /// Test converting all worksheets of xlsx file to csv.
    /// </summary>
    [Test]
    public void TestConvertXlsxWithDatesToCSV()
    {
        if (_input.Path != null)
        {
            _input.Path = Path.Combine(_input.Path, "TestDateFormat.xlsx");
            _options.ReadOnlyWorkSheetWithName = "Sheet2";
            _options.DateFormat = DateFormats.DDMMYYYY;
            _options.ShortDatePattern = true;
            var result = Excel.ConvertToCSV(_input, _options, new CancellationToken());
            var expectedResult = "25/12/2021,25/02/2021,12/05/2020,30/12/2021";

            if (result.CSV != null)
                Assert.That(Regex.Replace(result.CSV, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
            else
                Assert.Fail();
        }
        else
            throw new Exception($"Path missing, {_input.Path}");
    }

    /// <summary>
    /// Test converting one worksheet of xls file to csv.
    /// </summary>
    [Test]
    public void TestConvertXlsOneWorksheetToCSV()
    {
        if (_input.Path != null)
        {
            _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
            var result = Excel.ConvertToCSV(_input, _options, new CancellationToken());
            var expectedResult = "Foo,Bar,Kanji働,Summa1,2,3,6";

            if (result.CSV != null)
                Assert.That(Regex.Replace(result.CSV, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
            else
                Assert.Fail();
        }
        else
            throw new Exception($"Path missing, {_input.Path}");
    }

    /// <summary>
    /// Test converting one worksheet of xls file to csv.
    /// </summary>
    [Test]
    public void ShouldThrowUnknownFileFormatError()
    {
        if (_input.Path != null)
        {
            _input.Path = Path.Combine(_input.Path, "UnitTestErrorFile.txt");
            _options.ThrowErrorOnFailure = true;
            Assert.That(() => Excel.ConvertToCSV(_input, _options, new CancellationToken()), Throws.Exception);
        }
        else
            throw new Exception($"Path missing, {_input.Path}");
    }

    /// <summary>
    /// Try to convert a file that does not exist 
    /// </summary>
    [Test]
    public void DoNotThrowOnFailure()
    {
        if (_input.Path != null)
        {
            _input.Path = Path.Combine(_input.Path, "thisfiledoesnotexist.txt");
            _options.ThrowErrorOnFailure = false;
            try
            {
                var result = Excel.ConvertToCSV(_input, _options, new CancellationToken());
                Assert.AreEqual(result.Success, false);
                Assert.AreEqual(result.CSV, null);
            }
            catch (Exception ex)
            {
                Assert.Fail("This should not happen: " + ex.Message);
            }
        }
        else
            throw new Exception($"Path missing, {_input.Path}");
    }
}