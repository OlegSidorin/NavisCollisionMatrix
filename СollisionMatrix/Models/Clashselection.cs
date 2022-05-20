using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Models
{
    public class Clashselection
    {
        public static string NameOfClass = "clashselection";
        public string Tag_selfintersect { get; set; }
        public string Tag_primtypes { get; set; }
        public Locator Locator { get; set; }
    }
}
