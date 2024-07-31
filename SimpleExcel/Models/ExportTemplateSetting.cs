namespace SimpleExcel.Models
{
    public class ExportTemplateSetting<T> where T : class
    {
        public required List<T> Data { get; set; }

        public string SheetName { get; set; } = "Sheet1";

        public uint StartRow { get; set; } = 1;

        public uint OffsetColumns { get; set; } = 0;

        public bool Autofit { get; set; } = true;

        public uint SpaceAfter { get; set; } = 2;

        public Dictionary<string, CellTemplateQuery<T>>? FormatCellQueries { get; set; }

        public StyleSettings DefaultStyle { get; set; } = new StyleSettings();

        public StyleSettings HeaderStyle { get; set; } = new StyleSettings { TextBold = true};

    }

    public class CellTemplateQuery<T> where T: class
    {
        public required Func<T, bool> Query { get; set; }

        public required StyleSettings Style { get; set; }
    }

    public class StyleSettings
    {
        public uint? Id { get; set; }

        public string FontName { get; set; } = "Calibri";

        public uint FontSize { get; set; } = 11;

        // Eg: FF0000
        public string? ForeColorRgbHex { get; set; }

        public bool TextBold { get; set; } = false;
    }
}
