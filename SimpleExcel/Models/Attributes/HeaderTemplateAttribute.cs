namespace SimpleExcel.Models.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class HeaderTemplateAttribute : Attribute
    {
        public string DisplayName { get; set; }

        public HeaderTemplateAttribute(string displayName)
        {
            DisplayName = displayName;
        }

    }
}
