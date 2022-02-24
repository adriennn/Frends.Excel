# Frends.Excel.ConvertToXML

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![Build](https://github.com/FrendsPlatform/Frends.Excel/actions/workflows/ConvertToXML_main.yml/badge.svg)](https://github.com/FrendsPlatform/Frends.Excel/actions)
![MyGet](https://img.shields.io/myget/frends-tasks/v/Frends.Excel.ConvertToXML)
![Coverage](https://app-github-custom-badges.azurewebsites.net/Badge?key=FrendsPlatform/Frends.Excel/Frends.Excel.ConvertToXML|main)

Task for converting Excel files to XML.

## Installing

You can install the Task via frends UI Task View or you can find the NuGet package from the following NuGet feed
https://www.myget.org/F/frends-tasks/api/v2.

### Properties

| Property | Type | Description      | Example         |
| -------- | -------- |------------------|-----------------|
| Path | `string` | Excel file path. | `/tmp/file.xls` |

### Options

| Property  | Type  | Description |Example|
|-----------|-------|-------------|-------|
| ReadOnlyWorkSheetWithName  | string | Excel work sheet name to be read. If empty, all work sheets are read. |Sheet1| 
| UseNumbersAsColumnHeaders| bool | If set to true, numbers will be used as column headers instead of letters (A = 1, B = 2...) | true |
| ThrowErrorOnfailure| bool | Throws an exception if conversion fails. |  true |
| DateFormat | DateFormat | Selection for date format | Possible values: DDMMYYYY, MMDDYYYY, YYYYMMDD |
| ShortDatePattern | bool | Excludes timestamps from dates | false |

### Returns

| Property | Type   | Description                 |Example|
|----------|--------|-----------------------------|-------|
| XML      | string | Conversion result as string | |
| Success  | bool   | Task execution result.      | true |
| Message  | string | Exception message           | "File not found"|


## Building

Rebuild the project

`dotnet build`

Run tests

`dotnet test`

Create a NuGet package

`dotnet pack --configuration Release`
