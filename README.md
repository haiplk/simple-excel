# Super Simple Excel

Super Simple Excel is a lightweight .NET library for generating and reading Excel files with minimal settings, designed to keep things simple and straightforward. The library is based on OpenXML and supports .NET 8.

[![Buy Me a Coffee](data:image/webp;base64,UklGRuoEAABXRUJQVlA4TN4EAAAvqQAJEAfjuLZtpbnnK04BdEVDlEYfjCQyd/ekCDe1bTvRJlXxSsECPZ6wGgWE6Rm3sW2ryt73u2U/JKNQSqMRl9Cd6D+DbCPV3+JIXuUQzuC3QSD865+4OkchCtk28U3BtonxlwZW7SqXJOfc1eBgk5GtMoDrlSSXJkchCoVeSXI/FIpjJkkOQhBBEIIRRJAIgiAJRrByiV6iF/USvZ+/YG1wULvKXeWucji7wQ1usIUb3KG4Z4bMK0SoMGH97C2vk2ZZyYLUqBpRTmetpe379Pwm47az+708v8m47XTrAoXg3db2tq1t25KennNuOb09d/QkAARBECB6zo495z7u/zYI0Zbchn5H9H8CVvPT40P3dNU+O9Lhe3Q2O9FhfLJanelQPlsdHUxHpzqcjw+/Vx5beKNh7aFy/7H/sfWPJx7Mcj5M7t15+ce/Fz52/c49VagHybX1I48ufmR9TUOt6TA5/+eC59ekOOjAePv9d1+Xbj3+zItLX3jq8VtaOBKHPLiHz4CTVKf9eXuz2Xwo6blfP1n65c+vaelE0z90Mp3kYH8+3Lz+wUbSa9cWff38tUWQQlfID50JL3nK/ry/ee+j2c3nPl3y7dN3LiDJU/fHeHdZ3i9jFsgN540k6+0u3vh4s3lH0p0nP1/y3freEkuR1DPK5SClyZhxmDnbsGPoJrIxmeoldRW6LUOd+mBapsBoJJsqJBmQFFsJJqMA1e1Aeut1SXfWF17iybZLEBQZZcEFJkmW2vBMFcgjUI1MBXIrMR9bkQperjIPfkuUFAC8BYg7aT5Y//n74r+eXNYOyvQKVI0MkgZKwwLFV6h+wisw+S0eYg99q9LNMqUwkDyTpNwoxAE/kLt90JO/s/iv55b0UHMu4CteiaxKJ5nK0BBglSEq45WIgaExkuTBNTxVtcqBzXTkQG70ksAUXKFL9Hvw3F/L/r2mN999sxUZJClSQJqIAif10G8bpQhW4JUZKt44SdApghTGWZa1iowayeTYKHjJw0gWJKoxdmfX/l12flsfbt7dFmYBmOQgGJBsBa8yNYKUKLJgVYBBEScLUiFLIy1JiagOKIVBcswGoDoBBCV2dvt82fqu3vnwrVbGSzITE1kJfARvC9AlJsmClTKDPMhD9DJ1khwoQJJlnBVJAZKUmEcbAeszlGjVU6OXI+/s7nrRf2stzQzeD5UaKB7INOtEBSd5kFToZqaAkylESVCAYgpeshC7RIXQFUqcqMwrVIrkKxjZgt+Z1n/+vvCvZy/QzM4CZCABOVSoQVIgSxqrlaU5DZUiSQMQAZIkDQDFJYDqlIGpA0gWxgRQhsqo3b36+DMvtl946vFri/rYxxidpASDutjJeuuB0UqSr1FbE2RfgWxm6qJVqAxq9nnoJPU590bKjBBgJCsCte8ABrMj1xvppV8+aX/10zXt3lbA6+LOSiYEp33M+MRINhXJ9Z2RbAhWyy8jE6Vrr2755vmbe5Ao5VL2OuM7MlkZr0u/jKF20q3nPmt9+9Sd3VnwE+6q2SvRvPP4F63v1w9211MsVVetG0lMFszeaf1D67fHtfvIODJesR7ATUDR/j253vrcHgQAf8VMhaQA+B2cXtp+J2rQVTfeS3Leaoero4fB/8mj1dnBdLZanRxIJ6vVanV2dAAdna1W)](https://www.buymeacoffee.com/lJ7PtoK)


## Features

- Generate Excel files easily.
- Read Excel files effortlessly.
- Minimal configuration needed.
- Built on OpenXML.

## Installation

You can install Super Simple Excel via NuGet Package Manager:

```bash
dotnet add package SuperSimpleExcel
```


## Usage

### Generating an Excel File
Here's a simple example of how to generate an Excel file:

```csharp
using (var stream = await SimpleExcelFactory.CreateInstance().ExportExcelAsync(new SimpleExcel.Models.ExportTemplateSetting<StudentModel>
{
    Data = students,
    Autofit = true,
    SheetName = "Hello",
    HeaderStyle = new SimpleExcel.Models.StyleSettings { FontSize = 12, TextBold = true },
    FormatCellQueries = new Dictionary<string, SimpleExcel.Models.CellTemplateQuery<StudentModel>>
    {
        {
            nameof(StudentModel.FirstName),
            new SimpleExcel.Models.CellTemplateQuery<StudentModel> {
                    Query = student => student.YoB.HasValue && student.YoB.Value < 2000,
                    Style = new SimpleExcel.Models.StyleSettings {
                        ForeColorRgbHex = "FF0000"
                    }
        }   }
    }
}))
{
    using (FileStream fileStream = new FileStream("output.xlsx", FileMode.Create, FileAccess.Write))
    {
        stream.Seek(0, SeekOrigin.Begin);
        await stream.CopyToAsync(fileStream);

    }
}
```


### License
This project is licensed under the MIT License - see the LICENSE file for details.

### Contributing
Contributions are welcome! Please feel free to submit a pull request or open an issue.

### Acknowledgments
This library is built on top of the OpenXML SDK.

