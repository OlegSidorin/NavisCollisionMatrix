using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Models
{
    public class Summary
    {
        public static string NameOfClass = "summary";
        public string Tag_total { get; set; }
        public string Tag_new { get; set; }
        public string Tag_active { get; set; }
        public string Tag_reviewed { get; set; }
        public string Tag_approved { get; set; }
        public string Tag_resolved { get; set; }
        public int Tag_total_int { get; set; }
        public int Tag_new_int { get; set; }
        public int Tag_active_int { get; set; }
        public int Tag_reviewed_int { get; set; }
        public int Tag_approved_int { get; set; }
        public int Tag_resolved_int { get; set; }
        public Testtype Testtype { get; set; }
        public Teststatus Teststatus { get; set; }
    }
}
