using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix.Models
{
    public class Conditions
    {
        public static string NameOfClass = "conditions";
        public List<Condition> Conditions_list { get; set; }

        public Conditions()
        {
            Conditions_list = new List<Condition>();
        }
    }
}
