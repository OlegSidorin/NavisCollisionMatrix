using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Models
{
    public class Clashtest
    {
        public static string NameOfClass = "clashtest";
        public Linkage Linkage { get; set; }
        public Left Left { get; set; }
        public Right Right { get; set; }
        public Rules Rules { get; set; }
    }
}
