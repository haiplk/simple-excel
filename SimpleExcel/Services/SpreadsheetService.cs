using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using SimpleExcel.Models;
using SimpleExcel.Models.Attributes;
using System.Reflection;
namespace SimpleExcel.Services
{
    public class SpreadsheetService : ISpreadsheetService
    {
        const uint NormalStyleId = 1;
        const uint HeaderStyleId = 2;

        public Task<Stream> ExportExcelAsync<T>(ExportTemplateSetting<T> settings) where T : class
        {
            Stream memoryStream = new MemoryStream();
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                var allColumns = typeof(T).GetProperties()
                                          .Where(prop => Attribute.IsDefined(prop, typeof(HeaderTemplateAttribute)))
                                          .ToArray();

                var columnWidths = allColumns.ToDictionary(prop => prop, prop => 50);

                Stylesheet stylesheet = GenerateStylesheet(settings);
                Row header = GenerateHeaderRow(settings, allColumns, columnWidths);
                List<Row> rows = GenerateRowDataItems(settings, allColumns, columnWidths);

                if (settings.Autofit)
                {
                    AutofitColumns(settings, worksheetPart, allColumns, columnWidths);
                }

                // Adding style
                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = stylesheet;
                stylePart.Stylesheet.Save();

                // Add sheet
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = settings.SheetName };
                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                // Fill the data
                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                sheetData.AppendChild(header);
                sheetData.Append(rows);

                worksheetPart.Worksheet.Save();
            }

            return Task.FromResult(memoryStream);
        }

        private void AutofitColumns<T>(ExportTemplateSetting<T> settings, WorksheetPart worksheetPart, PropertyInfo[] allColumns, Dictionary<PropertyInfo, int> columnWidths) where T : class
        {
            Columns columns = new Columns();
            for (uint i = 0; i < allColumns.Length; i++)
            {
                var propInfo = allColumns[i];
                // Min = 1, Max = 1 ==> Apply this to column 1 (A)
                // Min = 2, Max = 2 ==> Apply this to column 2 (B)
                var columnIndex = i + 1 + settings.OffsetColumns;
                columns.AppendChild(new Column
                {
                    BestFit = false,
                    Min = columnIndex,
                    Max = columnIndex,
                    CustomWidth = true,
                    Width = columnWidths[propInfo] + settings.SpaceAfter
                });
            }
            worksheetPart.Worksheet.AppendChild(columns);
        }

        private List<Row> GenerateRowDataItems<T>(ExportTemplateSetting<T> settings, PropertyInfo[] allColumns, Dictionary<PropertyInfo, int> columnWidths) where T : class
        {
            List<Row> rows = new List<Row>();
            foreach (var data in settings.Data)
            {
                var row = new Row();

                foreach (var propInfo in allColumns)
                {
                    object value = propInfo.GetValue(data);
                    uint? styleId = null;

                    if (settings.FormatCellQueries != null && settings.FormatCellQueries.TryGetValue(propInfo.Name, out var queryTemplate))
                    {
                        if (queryTemplate.Query.Invoke(data))
                        {
                            styleId = queryTemplate.Style.Id;
                        }
                    }

                    row.AppendChild(ConstructCell(value, styleId));

                    // Autofit
                    var currentWidth = columnWidths[propInfo];
                    if (value != null && currentWidth < value.ToString().Length)
                    {
                        columnWidths[propInfo] = value.ToString().Length;
                    }
                }
                rows.Add(row);
            }

            return rows;
        }

        private Row GenerateHeaderRow<T>(ExportTemplateSetting<T> settings, PropertyInfo[] allColumns, Dictionary<PropertyInfo, int> columnWidths) where T : class
        {
            Row row = new Row();
            foreach (var prop in allColumns)
            {
                var attribute = prop.GetCustomAttribute<HeaderTemplateAttribute>();
                var cell = ConstructCell(attribute.DisplayName, settings.HeaderStyle.Id);
                row.AppendChild(cell);

                // Set custom width
                columnWidths[prop] = attribute.DisplayName.Length;
            }

            return row;
        }

        private Stylesheet GenerateStylesheet<T>(ExportTemplateSetting<T> settings) where T : class
        {
            Fonts fonts = new Fonts(ToFont(0, settings.DefaultStyle));

            fonts.Append(ToFont(NormalStyleId, settings.DefaultStyle));
            fonts.Append(ToFont(HeaderStyleId, settings.HeaderStyle));

            uint currentIndex = (uint)fonts.ToList().Count;
            if (settings.FormatCellQueries != null)
            {
                foreach (var item in settings.FormatCellQueries)
                {
                    fonts.Append(ToFont(currentIndex, item.Value.Style));
                    currentIndex++;
                }
            }

            var fills = new Fills(new Fill(new PatternFill() { PatternType = PatternValues.None }));
            var borders = new Borders(new Border());
            CellFormats cellFormats = new CellFormats(
                    new CellFormat(), // default
                    new CellFormat { FontId = settings.DefaultStyle.Id },
                    new CellFormat { FontId = settings.HeaderStyle.Id }
            );

            if (settings.FormatCellQueries != null)
            {
                foreach (var item in settings.FormatCellQueries)
                {
                    cellFormats.Append(new CellFormat { FontId = item.Value.Style.Id });
                }
            }

            Stylesheet styleSheet = new Stylesheet(fonts, fills, borders, cellFormats);

            return styleSheet;
        }

        private Font ToFont(uint index, StyleSettings styleSettings)
        {
            styleSettings.Id = index;
            var font = new Font();
            font.FontSize = new FontSize { Val = styleSettings.FontSize };
            font.FontName = new FontName { Val = styleSettings.FontName };
            if (styleSettings.ForeColorRgbHex != null)
            {
                font.Color = new Color() { Rgb = styleSettings.ForeColorRgbHex };
            }

            if (styleSettings.TextBold)
            {
                font.Bold = new Bold();
            }

            return font;
        }

        private Cell ConstructCell(object? value, uint? styleIndex = 0)
        {
            CellValue cellVal = new CellValue();
            var dataType = new EnumValue<CellValues>(CellValues.String);

            if (value != null)
            {
                Type valueType = value.GetType();
                if (valueType == typeof(DateTime))
                {
                    cellVal = new CellValue((DateTime)value);
                    dataType = new EnumValue<CellValues>(CellValues.Date);
                }
                else 
                if (valueType == typeof(int) ||
                 valueType == typeof(double) ||
                 valueType == typeof(float) ||
                 valueType == typeof(decimal) ||
                 valueType == typeof(long) ||
                 valueType == typeof(short) ||
                 valueType == typeof(uint) ||
                 valueType == typeof(ulong) ||
                 valueType == typeof(ushort) ||
                 valueType == typeof(byte) ||
                 valueType == typeof(sbyte))
                {
                    cellVal = new CellValue(value.ToString());
                    dataType = new EnumValue<CellValues>(CellValues.Number);
                }
                else
                {
                    // Default to string if no specific type is matched
                    cellVal = new CellValue(value.ToString());
                    dataType = new EnumValue<CellValues>(CellValues.String);
                }
            }


            return new Cell()
            {
                CellValue = cellVal,
                DataType = dataType,
                StyleIndex = styleIndex
            };
        }

        public Task<List<T>> ReadExcelFileAsync<T>(Stream filestream, ImportTemplateSettings settings) where T : class
        {
            var allColumns = typeof(T).GetProperties()
                        .Select(prop => new
                        {
                            Property = prop,
                            Attribute = prop.GetCustomAttribute<ImportTemplateColumnAttribute>()
                        })
                        .Where(x => x.Attribute != null)
                        .ToArray();

            var list = new List<T>();
            using (var document = SpreadsheetDocument.Open(filestream, false))
            {
                var workbookPart = document.WorkbookPart;
                var sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>();
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                var rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

                int currentRow = 0;

                foreach (var row in rows)
                {
                    if (currentRow >= settings.StartRow)
                    {
                        bool canAdd = false;
                        var model = Activator.CreateInstance<T>();

                        foreach (var column in allColumns)
                        {
                            if (column.Property.PropertyType == typeof(string))
                            {
                                string value = GetCellValue(workbookPart, row, column.Attribute.ColumnIndex);
                                column.Property.SetValue(model, value);

                                if (!string.IsNullOrWhiteSpace(value))
                                {
                                    canAdd = true;
                                }
                            }
                        }

                        if (canAdd)
                        {
                            list.Add(model);
                        }
                    }

                    currentRow++;
                }
            }

            return Task.FromResult(list);
        }

        private string GetCellValue(WorkbookPart workbookPart, Row row, uint columnIndex)
        {
            var cell = row.Elements<Cell>().ElementAtOrDefault((int)columnIndex);
            if (cell == null) return string.Empty;

            var value = cell.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
            }
            return value;
        }
    }
}
