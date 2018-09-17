using System;

namespace webApp.Models
{
    public class Student
    {
        public int StudentId { get; set; }
        public int RoleNumber { get; set; }
        public string Name { get; set; }
        public bool Active { get; set; }
        public DateTime Dob { get; set; }
    }
}
