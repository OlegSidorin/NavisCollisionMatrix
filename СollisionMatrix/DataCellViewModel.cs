using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace СollisionMatrix
{
    public class DataCellViewModel : ObservableObject
    {
        private ClashTest _clashTest;
        public ClashTest ClashTest
        {
            get { return _clashTest; }
            set
            {
                if (_clashTest != value)
                {
                    _clashTest = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _selection1Name;
        public string Selection1Name
        {
            get { return _selection1Name; }
            set
            {
                if (_selection1Name != value)
                {
                    _selection1Name = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _selection2Name;
        public string Selection2Name
        {
            get { return _selection2Name; }
            set
            {
                if (_selection2Name != value)
                {
                    _selection2Name = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _collisionsNumber;
        public string CollisionsNumber
        {
            get { return _collisionsNumber; }
            set
            {
                if (_collisionsNumber != value)
                {
                    _collisionsNumber = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}
