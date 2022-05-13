using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix
{
    public class MatrixSelectionLineModel
    {
        public string NameOfSelection { get; set; }
        public List<string> SelectionIntersectionTolerance { get; set; }

        public MatrixSelectionLineModel()
        {
            SelectionIntersectionTolerance = new List<string>();
        }
    }
}
