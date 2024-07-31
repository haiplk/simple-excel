namespace SimpleExcel.Models.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ImportTemplateColumnAttribute : Attribute
    {
        public ImportTemplateColumnAttribute(uint columnIndex)
        {
            ColumnIndex = columnIndex;
        }

        public uint ColumnIndex { get; set; }

        public bool IgnoreError { get; set; } = true;

    }
}
