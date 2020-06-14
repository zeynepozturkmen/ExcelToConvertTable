using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelToConvertTable.Models
{
    public class PersonModel
    {
        public List<PersonTable> PersonList { get; set; }
        public PersonTable Person { get; set; }
    }
    public class PersonTable
    {
        public string FullName { get; set; }
        public string Email { get; set; }
        public string Address { get; set; }
        public string PhoneNumber { get; set; }
        public string TC { get; set; }
    }
}