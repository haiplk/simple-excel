using SimpleExcel.Services;

namespace SimpleExcel
{
    public static class SimpleExcelFactory
    {
        public static ISpreadsheetService CreateInstance()
        {
            return new SpreadsheetService();
        }
    }
}
