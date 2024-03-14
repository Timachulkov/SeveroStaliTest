using System;

namespace SeveroStaliTest
{
    class Employees
    {
        public ulong TableID { get; set; }
        public string Surname { get; set; }
        public string Name { get; set; }
        public string Patronymic { get; set; }
        public DateTime BirthDate { get; set; }
        public ulong DepartmentID { get; set; }
    }
}
