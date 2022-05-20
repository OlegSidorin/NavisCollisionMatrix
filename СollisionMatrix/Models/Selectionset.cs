using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Models
{
    public class Selectionset
    {
        public static string NameOfClass = "selectionset";
        public string Tag_name { get; set; }
        public string Tag_guid { get; set; }
        public Findspec Findspec { get; set; }
    }
}
