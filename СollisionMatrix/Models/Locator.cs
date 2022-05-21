using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Models
{
    public class Locator
    {
        public static string NameOfClass = "locator";
        public string Tag_inner_text { get; set; }
        public string Tag_inner_text_selection_name { get; set; }
        public string Tag_inner_text_draft_name { get; set; }
        public List<string> Tag_inner_text_folders { get; set; }
    }
}
