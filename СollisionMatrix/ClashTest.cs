using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix
{
    public class ClashTest
    {
        public string Name { get; set; }
        public string SelectionLeftName { get; set; }
        public string SelectionRightName { get; set; }
        public string Tolerance { get; set; }
        public string SummaryTestType { get; set; } 
        public string SummaryTotal { get; set; } 
        public string SummaryNew { get; set; }
        public string SummaryActive { get; set; }
        public string SummaryReviewed { get; set; }
        public string SummaryApproved { get; set; }
        public string SummaryResolved { get; set; }
    }

    /*
     <clashtest name="АР_Вертикальные конструкции - КР_Каркас несущий" test_type="hard" status="ok" tolerance="0.005" merge_composites="0">
        <linkage mode="none"/>
        <left>
          <clashselection selfintersect="0" primtypes="1">
            <locator>lcop_selection_set_tree/АР_Вертикальные конструкции</locator>
          </clashselection>
        </left>
        <right>
          <clashselection selfintersect="0" primtypes="1">
            <locator>lcop_selection_set_tree/КР_Каркас несущий</locator>
          </clashselection>
        </right>
        <rules/>
        <summary total="68" new="0" active="0" reviewed="0" approved="0" resolved="68">
          <testtype>По пересечению</testtype>
          <teststatus>ОК</teststatus>
        </summary>
        <clashresults/>
      </clashtest>
     
     */
}
