using Excelify.Services.Utility.Attributes;

namespace Excelify.API.Models
{
    public class TeacherDTO
    {
        [Excelify("id")]
        public int Id { get; set; }

        [Excelify("first_name")]
        public string FirstName { get; set; }

        [Excelify("last_name")]
        public string LastName { get; set; }

        [Excelify("title")]
        public string Title { get; set; }

        [Excelify("phone_number")]
        public int PhoneNumber { get; set; }

        [Excelify("email")]
        public string Email { get; set; }

        [Excelify("start_date")]
        public DateTime StartDate { get; set; }
    }
}
