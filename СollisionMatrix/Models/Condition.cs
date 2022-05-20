using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Models
{
    public class Condition
    {
        public static string NameOfClass = "condition";
        public string Tag_test { get; set; }
        public string Tag_flags { get; set; }
        public Category Category { get; set; }
        public Property Property { get; set; }
        public Value Value { get; set; }
    }
}
