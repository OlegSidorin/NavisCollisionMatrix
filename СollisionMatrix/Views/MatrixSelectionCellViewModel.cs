﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace СollisionMatrix
{
    public class MatrixSelectionCellViewModel : ObservableObject
    {
        private Models.Clashtest _clashtest;
        public Models.Clashtest Clashtest
        {
            get { return _clashtest; }
            set
            {
                if (_clashtest != value)
                {
                    _clashtest = value;
                    OnPropertyChanged();
                }
            }
        }

        //JS.Clashtest jclashtest;
        //public JS.Clashtest JClashtest { get { return jclashtest; } set { jclashtest = value; OnPropertyChanged(); }  }

        private string tolerance;
        public string Tolerance { get { return tolerance; } set { tolerance = value; OnPropertyChanged(); } }

        //public string Tolerance
        //{
        //    get 
        //    {
        //        bool ok = double.TryParse(jclashtest.Tolerance, out double result);
        //        if (ok)
        //        {
        //            return ((int)(result * 304.8)).ToString();
        //        }
        //        else
        //        {
        //            return jclashtest.Tolerance;
        //        }
        //    }
        //    set
        //    {
        //        bool ok = int.TryParse(value, out int result);
        //        if (ok)
        //        {
        //            tolerance = (result / 304.8).ToString();
        //        }
        //        else
        //        {
        //            tolerance = value;
        //        }
        //    }
        //}
    }
}
