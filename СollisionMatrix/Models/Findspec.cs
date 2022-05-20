using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Models
{
    public class Findspec
    {
        public static string NameOfClass = "findspec";
        public string Tag_mode { get; set; }
        public string Tag_disjoint { get; set; }
        public Conditions Conditions { get; set; }
        public Locator Locator { get; set; }
    }
}
