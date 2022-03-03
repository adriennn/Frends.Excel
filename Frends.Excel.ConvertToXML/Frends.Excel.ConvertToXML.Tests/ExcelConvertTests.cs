using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using Frends.Excel.ConvertToXML.Definitions;
using NUnit.Framework;

namespace Frends.Excel.ConvertToXML.Tests;

[TestFixture]
public class ExcelConvertTests
{
    private readonly Input _input = new();
    private readonly Options _options = new();

    // Cat image in example files is from Pixbay.com. It is licenced in CC0 Public Domain (Free for commercial use, No attribution required).
    // It is uploaded by Ben_Kerckx https://pixabay.com/en/cat-animal-pet-cats-close-up-300572/


    [SetUp]
    public void Setup()
    {
        _input.Path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"../../../../../TestData/");
        _options.ReadOnlyWorkSheetWithName = "";

    }

    [Test]
    public void TestConvertXlsxToXML()
    {
        // Test converting all worksheets of xlsx file to xml.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
        var result = Excel.ConvertToXML(_input, _options, new CancellationToken());
        var expectedResult = @"<workbookworkbook_name=""ExcelTestInput1.xlsx""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet><worksheetworksheet_name=""OmituinenNimi""><rowrow_header=""1""><columncolumn_header=""A"">Kissakuva</column><columncolumn_header=""B"">1</column><columncolumn_header=""C"">2</column><columncolumn_header=""D"">3</column></row><rowrow_header=""15""><columncolumn_header=""A"">Foo</column></row><rowrow_header=""16""><columncolumn_header=""B"">Bar</column></row></worksheet></workbook>";
        Assert.That(Regex.Replace(result.XML, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsToXML()
    {
        // Test converting all worksheets of xls file to xml.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput2.xls");
        var result = Excel.ConvertToXML(_input, _options, new CancellationToken());
        var expectedResult = @"<workbookworkbook_name=""ExcelTestInput2.xls""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet><worksheetworksheet_name=""OmituinenNimi""><rowrow_header=""1""><columncolumn_header=""A"">Kissakuva</column><columncolumn_header=""B"">1</column><columncolumn_header=""C"">2</column><columncolumn_header=""D"">3</column></row><rowrow_header=""15""><columncolumn_header=""A"">Foo</column></row><rowrow_header=""16""><columncolumn_header=""B"">Bar</column></row></worksheet></workbook>";
        Assert.That(Regex.Replace(result.XML, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsxWithDatesToXML()
    {
        // Test converting all worksheets of xlsx file to xml.
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xlsx");
        _options.ReadOnlyWorkSheetWithName = "Sheet2";
        _options.DateFormat = DateFormats.YYYYMMDD;
        _options.ShortDatePattern = false;
        var result = Excel.ConvertToXML(_input, _options, new CancellationToken());
        var expectedResult = @"<workbookworkbook_name=""TestDateFormat.xlsx""><worksheetworksheet_name=""Sheet2""><rowrow_header=""1""><columncolumn_header=""A"">2021/12/25 0:00:00</column><columncolumn_header=""B"">2021/02/25 12:45:41</column><columncolumn_header=""C"">2020/05/12 0:00:00</column><columncolumn_header=""D"">2021/12/30 0:00:00</column></row></worksheet></workbook>";
        Assert.That(Regex.Replace(result.XML, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsxWithDatesAndShortDatePatternToXML()
    {
        // Test converting all worksheets of xlsx file to xml.
        _input.Path = Path.Combine(_input.Path, "TestDateFormat.xlsx");
        _options.ReadOnlyWorkSheetWithName = "Sheet1";
        _options.DateFormat = DateFormats.MMDDYYYY;
        _options.ShortDatePattern = true;
        var result = Excel.ConvertToXML(_input, _options, new CancellationToken());
        var expectedResult = @"<workbookworkbook_name=""TestDateFormat.xlsx""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">4</column></row><rowrow_header=""2""><columncolumn_header=""A"">12/12/2021</column><columncolumn_header=""B"">02/25/2021</column><columncolumn_header=""C"">05/12/2020</column><columncolumn_header=""D"">12/12/2021</column></row></worksheet></workbook>";
        Assert.That(Regex.Replace(result.XML, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void TestConvertXlsxOneWorksheetToXML()
    {
        // Test converting one worksheet of xlsx file to xml.
        _input.Path = Path.Combine(_input.Path, "ExcelTestInput1.xlsx");
        _options.ReadOnlyWorkSheetWithName = "Sheet1";
        var result = Excel.ConvertToXML(_input, _options, new CancellationToken());
        var expectedResult = @"<workbookworkbook_name=""ExcelTestInput1.xlsx""><worksheetworksheet_name=""Sheet1""><rowrow_header=""1""><columncolumn_header=""A"">Foo</column><columncolumn_header=""B"">Bar</column><columncolumn_header=""C"">Kanji働</column><columncolumn_header=""D"">Summa</column></row><rowrow_header=""2""><columncolumn_header=""A"">1</column><columncolumn_header=""B"">2</column><columncolumn_header=""C"">3</column><columncolumn_header=""D"">6</column></row></worksheet></workbook>";
        Assert.That(Regex.Replace(result.XML, @"[\s+]", ""), Does.StartWith(Regex.Replace(expectedResult.ToString(), @"[\s+]", "")));
    }

    [Test]
    public void ShouldThrowUnknownFileFormatError()
    {
        // Test converting one worksheet of xls file to csv.
        _input.Path = Path.Combine(_input.Path, "UnitTestErrorFile.txt");
        _options.ThrowErrorOnFailure = true;
        Assert.That(() => Excel.ConvertToXML(_input, _options, new CancellationToken()), Throws.Exception);
    }

    [Test]
    public void DoNotThrowOnFailure()
    {
        //try to convert a file that does not exist 
        _input.Path = Path.Combine(_input.Path, "thisfiledoesnotexist.txt");
        _options.ThrowErrorOnFailure = false;
        try
        {
            var result = Excel.ConvertToXML(_input, _options, new CancellationToken());
            Assert.AreEqual(result.Success, false);
            Assert.AreEqual(result.XML, null);
        }
        catch (Exception ex)
        {
            Assert.Fail("This should not happen: " + ex.Message);
        }
    }
}