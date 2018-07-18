using System;
using System.Collections.Generic;
using System.Text;

namespace ProcessExcel
{
    public class CourseModule
    {

        public string Module { get; set; }
        public float Points { get; set; }
        public int Classes { get; set; }
        public int Labs { get; set; }
        public string Instructor { get; set; }
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public bool? Weekend { get; set; }

    }
}
