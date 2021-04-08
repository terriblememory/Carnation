using System.Globalization;
using System.Windows.Media;

namespace Carnation.Models
{
    internal class ObservableColor : NotifyPropertyBase
    {
        private enum UpdateBehavior
        {
            None,
            FromColor,
            FromComponent
        }

        private UpdateBehavior _updateBehavior = UpdateBehavior.None;

        public ObservableColor(Color color)
        {
            Color = color;
        }

        private Color _color;
        public Color Color
        {
            get => _color;
            set
            {
                if (!SetProperty(ref _color, value)) return;
                if (_updateBehavior != UpdateBehavior.FromComponent) UpdateColorComponents();
            }
        }

        private string _hex;
        public string Hex
        {
            get => _hex;
            set
            {
                if (value.StartsWith("#")) value = value.Substring(1); // remove leading #
                if (value.Length == 8) value = value.Substring(2); // remove alpha
                if (!SetProperty(ref _hex, value)) return; // set the string property
                if (value.Length != 6) return; // don't parse further unless six chars
                if (_updateBehavior != UpdateBehavior.None) return;
                var argb = uint.Parse(value, NumberStyles.HexNumber, CultureInfo.CurrentCulture.NumberFormat);
                Color = ColorHelpers.ToColor(argb);
            }
        }

        private double _hue;
        public double Hue
        {
            get => _hue;
            set
            {
                if (!SetProperty(ref _hue, value)) return;
                if (_updateBehavior == UpdateBehavior.None) UpdateColorFromHSB();
            }
        }

        private double _saturation;
        public double Saturation
        {
            get => _saturation;
            set
            {
                if (!SetProperty(ref _saturation, value)) return;
                if (_updateBehavior == UpdateBehavior.None) UpdateColorFromHSB();
            }
        }

        private double _brightness;
        public double Brightness
        {
            get => _brightness;
            set
            {
                if (!SetProperty(ref _brightness, value)) return;
                if (_updateBehavior == UpdateBehavior.None) UpdateColorFromHSB();
            }
        }

        private byte _red;
        public byte Red
        {
            get => _red;
            set
            {
                if (!SetProperty(ref _red, value)) return;
                if (_updateBehavior == UpdateBehavior.None) UpdateColorFromRGB();
            }
        }

        private byte _green;
        public byte Green
        {
            get => _green;
            set
            {
                if (!SetProperty(ref _green, value)) return;
                if (_updateBehavior == UpdateBehavior.None) UpdateColorFromRGB();
            }
        }

        private byte _blue;
        public byte Blue
        {
            get => _blue;
            set
            {
                if (!SetProperty(ref _blue, value)) return;
                if (_updateBehavior == UpdateBehavior.None) UpdateColorFromRGB();
            }
        }

        private void UpdateColorFromRGB()
        {
            _updateBehavior = UpdateBehavior.FromComponent;

            Color = Color.FromRgb(Red, Green, Blue);
            Hue = ColorHelpers.GetHue(Color);
            Saturation = ColorHelpers.GetSaturation(Color);
            Brightness = ColorHelpers.GetBrightness(Color);
            Hex = Color.ToString();

            _updateBehavior = UpdateBehavior.None;
        }

        private void UpdateColorFromHSB()
        {
            _updateBehavior = UpdateBehavior.FromComponent;

            Color = ColorHelpers.HSVToColor(Hue, Saturation, Brightness);
            Red = Color.R;
            Green = Color.G;
            Blue = Color.B;
            Hex = Color.ToString();

            _updateBehavior = UpdateBehavior.None;
        }

        private void UpdateColorComponents()
        {
            _updateBehavior = UpdateBehavior.FromColor;

            Red = Color.R;
            Green = Color.G;
            Blue = Color.B;
            Hue = ColorHelpers.GetHue(Color);
            Saturation = ColorHelpers.GetSaturation(Color);
            Brightness = ColorHelpers.GetBrightness(Color);
            Hex = Color.ToString();

            _updateBehavior = UpdateBehavior.None;
        }
    }
}
