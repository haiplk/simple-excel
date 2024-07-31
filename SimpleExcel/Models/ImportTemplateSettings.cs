namespace SimpleExcel.Models
{
    public class ImportTemplateSettings
    {
        public bool IsSkipBlankRowa { get; set; } = true;

        public uint StartRow { get; set; } = 1; //Skip header

    }
}
