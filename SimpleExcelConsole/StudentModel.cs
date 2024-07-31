using SimpleExcel.Models.Attributes;

namespace SimpleExcelConsole
{
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
}
