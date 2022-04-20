using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix
{
    public class SelectionSet : ObservableObject
    {
        public string Name { get; set; }
        public List<string> Selections { get; set; }
    }
}
