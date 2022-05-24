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
        public string Tag_name { get; set; }
        public string Tag_test_type { get; set; }
        public string Tag_status { get; set; }
        public string Tag_tolerance { get; set; }
        public string Tag_tolerance_in_mm { get; set; }
        public string Tag_merge_composites { get; set; }
        public Linkage Linkage { get; set; }
        public Left Left { get; set; }
        public Right Right { get; set; }
        public Rules Rules { get; set; }
        public Summary Summary { get; set; }
    }
}
