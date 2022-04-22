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

        private string _collisionsTotalNumber;
        public string CollisionsTotalNumber
        {
            get { return _collisionsTotalNumber; }
            set
            {
                if (_collisionsTotalNumber != value)
                {
                    _collisionsTotalNumber = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _collisionsActiveNumber;
        public string CollisionsActiveNumber
        {
            get { return _collisionsActiveNumber; }
            set
            {
                if (_collisionsActiveNumber != value)
                {
                    _collisionsActiveNumber = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _collisionsResolvedNumber;
        public string CollisionsResolvedNumber
        {
            get { return _collisionsResolvedNumber; }
            set
            {
                if (_collisionsResolvedNumber != value)
                {
                    _collisionsResolvedNumber = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _collisionsReviewedNumber;
        public string CollisionsReviewedNumber
        {
            get { return _collisionsReviewedNumber; }
            set
            {
                if (_collisionsReviewedNumber != value)
                {
                    _collisionsReviewedNumber = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _collisionsApprovedNumber;
        public string CollisionsApprovedNumber
        {
            get { return _collisionsApprovedNumber; }
            set
            {
                if (_collisionsApprovedNumber != value)
                {
                    _collisionsApprovedNumber = value;
                    OnPropertyChanged();
                }
            }
        }
    }
}
