using SimpleExcel.Models;

namespace SimpleExcel
{
    public interface ISpreadsheetService
    {
        Task<Stream> ExportExcelAsync<T>(ExportTemplateSetting<T> settings) where T : class;

        Task<List<T>> ReadExcelFileAsync<T>(Stream filestream, ImportTemplateSettings settings) where T : class;
    }
}
