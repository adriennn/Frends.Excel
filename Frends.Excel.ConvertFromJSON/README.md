# Frends.Excel.ConvertFromJSON

[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)
[![Build](https://github.com/FrendsPlatform/Frends.Excel/actions/workflows/ConvertFromJSON_main.yml/badge.svg)](https://github.com/FrendsPlatform/Frends.Excel/actions)
![MyGet](https://img.shields.io/myget/frends-tasks/v/Frends.Excel.ConvertFromJSON)
![Coverage](https://app-github-custom-badges.azurewebsites.net/Badge?key=FrendsPlatform/Frends.Excel/Frends.Excel.ConvertFromJSON|main)

Task for converting JSON Objects into an Excel Sheet.

## Installing

You can install the Task via frends UI Task View or you can find the NuGet package from the following NuGet feed
https://www.myget.org/F/frends-tasks/api/v2.

### Properties

| Property | Type | Description      | Example         |
| -------- | -------- |------------------|-----------------|
| Input | `string` | JToken representation of a JSON object. |  |

### Options

| Property  | Type  | Description |Example|
|-----------|-------|-------------|-------|
| WriteToWorkSheetName  | string | Excel work sheet name to be written to. If empty, the data is written to Sheet1. |Sheet1|
| HAsHeaders| bool | If set to true, The first row of the worksheet will contain the column names, else only the datarows will be written | true |
| ThrowErrorOnfailure| bool | Throws an exception if conversion fails. |  true |

### Returns

| Property | Type   | Description                  |Example|
|----------|--------|------------------------------|-------|
| ExcelFilePath     | string | Local path to the Excel file |       |
| Success  | bool   | Task execution result.       | true  |
| Message  | string | Exception message            | "File not found"|


## Building

Rebuild the project

`dotnet build`

Run tests

`dotnet test`

Create a NuGet package

`dotnet pack --configuration Release`
