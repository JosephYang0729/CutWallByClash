using System;
using System.ComponentModel;

namespace CutWallByClash.Models
{
    public enum OpeningType
    {
        Rectangular,
        Circular
    }

    public enum MEPCategory
    {
        Pipes,
        Ducts,
        CableTray,
        Conduit,
        FlexPipe
    }

    public class OpeningParameters : INotifyPropertyChanged
    {
        private OpeningType _openingType = OpeningType.Rectangular;
        private double _rectangularWidth = 200; // mm
        private double _rectangularHeight = 200; // mm
        private double _circularDiameter = 200; // mm
        private double _wallThickness = 50; // mm
        private double _mergeDistance = 100; // mm

        public OpeningType OpeningType
        {
            get => _openingType;
            set
            {
                _openingType = value;
                OnPropertyChanged(nameof(OpeningType));
            }
        }

        public double RectangularWidth
        {
            get => _rectangularWidth;
            set
            {
                _rectangularWidth = value;
                OnPropertyChanged(nameof(RectangularWidth));
            }
        }

        public double RectangularHeight
        {
            get => _rectangularHeight;
            set
            {
                _rectangularHeight = value;
                OnPropertyChanged(nameof(RectangularHeight));
            }
        }

        public double CircularDiameter
        {
            get => _circularDiameter;
            set
            {
                _circularDiameter = value;
                OnPropertyChanged(nameof(CircularDiameter));
            }
        }

        public double WallThickness
        {
            get => _wallThickness;
            set
            {
                _wallThickness = value;
                OnPropertyChanged(nameof(WallThickness));
            }
        }

        public double MergeDistance
        {
            get => _mergeDistance;
            set
            {
                _mergeDistance = value;
                OnPropertyChanged(nameof(MergeDistance));
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}