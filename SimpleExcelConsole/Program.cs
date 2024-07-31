using SimpleExcel;

namespace SimpleExcelConsole
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var students = GenerateStudentList(20);

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
        }

        static List<StudentModel> GenerateStudentList(int count)
        {
            var students = new List<StudentModel>();
            var random = new Random();
            var firstNames = new[] { "John", "Jane", "Alex", "Emily", "Chris", "Katie", "Mike", "Sara", "Tom", "Laura" };
            var lastNames = new[] { "Smith", "Johnson", "Brown", "Williams", "Jones", "Davis", "Garcia", "Miller", "Wilson", "Moore" };

            for (int i = 0; i < count; i++)
            {
                var student = new StudentModel
                {
                    Id = Guid.NewGuid(),
                    FirstName = firstNames[random.Next(firstNames.Length)],
                    LastName = lastNames[random.Next(lastNames.Length)],
                    Email = $"student{i}@example.com",
                    YoB = random.Next(1990, 2010)
                };

                students.Add(student);
            }

            return students;
        }


    }
}
