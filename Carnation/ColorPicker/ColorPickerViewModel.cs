using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media;
using Carnation.Helpers;
using Carnation.Models;

namespace Carnation
{
    internal class ColorPickerViewModel : NotifyPropertyBase
    {
        public ColorPickerViewModel()
        {
            ForegroundColor = new ObservableColor(Colors.Transparent);
            BackgroundColor = new ObservableColor(Colors.Transparent);
            CurrentEditorColor = ForegroundColor;
        }

        private ObservableColor _currentEditorColor;
        public ObservableColor CurrentEditorColor
        {
            get => _currentEditorColor;
            set
            {
                if (SetProperty(ref _currentEditorColor, value))
                {
                    NotifyPropertyChanged(nameof(IsForegroundBeingEdited));
                }
            }
        }

        private ObservableColor _backgroundColor;
        public ObservableColor BackgroundColor
        {
            get => _backgroundColor;
            set => SetProperty(ref _backgroundColor, value);
        }

        private ObservableColor _foregroundColor;
        public ObservableColor ForegroundColor
        {
            get => _foregroundColor;
            set => SetProperty(ref _foregroundColor, value);
        }

        private FontFamily _sampleTextFontFamily;
        public FontFamily SampleTextFontFamily
        {
            get => _sampleTextFontFamily;
            set => SetProperty(ref _sampleTextFontFamily, value);
        }

        private double _sampleTextFontSize;
        public double SampleTextFontSize
        {
            get => _sampleTextFontSize;
            set => SetProperty(ref _sampleTextFontSize, value);
        }

        public void SetForegroundColor(Color color)
        {
            ForegroundColor.Color = color;
        }

        public void SetBackgroundColor(Color color)
        {
            BackgroundColor.Color = color;
        }

        public bool IsForegroundBeingEdited => CurrentEditorColor == ForegroundColor;
    }
}
