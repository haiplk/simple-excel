# Super Simple Excel

Super Simple Excel is a lightweight .NET library for generating and reading Excel files with minimal settings, designed to keep things simple and straightforward. The library is based on OpenXML and supports .NET 8.

[![Buy Me a Coffee](https://raw.githubusercontent.com/haiplk/simple-excel/main/buycoffee.png)](https://www.buymeacoffee.com/lJ7PtoK)


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
 public class StudentModel
 {
     [HeaderTemplate("Student Id")]
     public Guid Id { get; set; }

     [HeaderTemplate("First name")]
     public string FirstName { get; set; }

     [HeaderTemplate("Last name")]
     public string LastName { get; set; }

     [HeaderTemplate("Email")]
     public string Email { get; set; }

     [HeaderTemplate("Year of Birth")]
     public int? YoB { get; set; }
 }
```

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

