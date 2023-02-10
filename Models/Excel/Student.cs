using System.ComponentModel.DataAnnotations;

namespace StudentApplication.Models.Excel
{
    public class Student
    {
        [Key]
        public int Id { get; set; }

        public int ExternalStudentID { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string DOB { get; set; }

        public string SSN { get; set; }

        public string Adddress { get; set; }

        public string City { get; set; }

        public string State { get; set; }

        public string Email { get; set; }

        public string MaritalStatus { get; set; }
    }
}
