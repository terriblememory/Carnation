using System.Diagnostics.Contracts;
using System.Windows.Media;
using static Microsoft.VisualStudio.Shell.ThreadHelper;

namespace Carnation
{
    internal partial class ClassificationProvider
    {
        internal abstract class ColorItemBase : NotifyPropertyBase
        {
            // Foreground

            private uint _foregroundColorRef;
            public uint ForegroundColorRef
            {
                get => _foregroundColorRef;
                set => SetProperty(ref _foregroundColorRef, value, nameof(Foreground));
            }

            public Color Foreground
            {
                get
                {
                    ThrowIfNotOnUIThread();
                    if (!IsForegroundEditable) return Colors.Transparent;
                    return FontsAndColorsHelper.TryGetColor(ForegroundColorRef) ?? DefaultForeground;
                }
                set
                {
                    ThrowIfNotOnUIThread();
                    Contract.Assert(IsUpdating || IsForegroundEditable);
                    ForegroundColorRef = FontsAndColorsHelper.GetColorRef(value, DefaultForeground);
                }
            }

            private readonly uint _autoForegroundColorRef;
            public uint AutoForegroundColorRef
            {
                get => _autoForegroundColorRef;
                set => SetProperty(ref _autoBackgroundColorRef, value, nameof(DefaultForeground));
            }

            public Color DefaultForeground
            {
                get
                {
                    ThrowIfNotOnUIThread();
                    return FontsAndColorsHelper.TryGetColor(AutoForegroundColorRef) ?? PlainTextForeground;
                }
            }

            private bool _isForegroundEditable = true;
            public bool IsForegroundEditable
            {
                get => _isForegroundEditable;
                set => SetProperty(ref _isForegroundEditable, value);
            }

            // Background

            private uint _backgroundColorRef;
            public uint BackgroundColorRef
            {
                get => _backgroundColorRef;
                set => SetProperty(ref _backgroundColorRef, value, nameof(Background));
            }

            public Color Background
            {
                get
                {
                    ThrowIfNotOnUIThread();
                    if (!IsBackgroundEditable) return Colors.Transparent;
                    return FontsAndColorsHelper.TryGetColor(BackgroundColorRef) ?? DefaultBackground;
                }
                set
                {
                    ThrowIfNotOnUIThread();
                    Contract.Assert(IsUpdating || IsBackgroundEditable);
                    BackgroundColorRef = FontsAndColorsHelper.GetColorRef(value, DefaultBackground);
                }
            }

            private uint _autoBackgroundColorRef;
            public uint AutoBackgroundColorRef
            {
                get => _autoBackgroundColorRef;
                set => SetProperty(ref _autoBackgroundColorRef, value, nameof(DefaultBackground));
            }

            public Color DefaultBackground
            {
                get
                {
                    ThrowIfNotOnUIThread();
                    return FontsAndColorsHelper.TryGetColor(AutoBackgroundColorRef) ?? PlainTextBackground;
                }
            }

            private bool _isBackgroundEditable = true;
            public bool IsBackgroundEditable
            {
                get => _isBackgroundEditable;
                set => SetProperty(ref _isBackgroundEditable, value);
            }

            // Bold

            private bool _isBold;
            public bool IsBold
            {
                get => _isBold;
                set
                {
                    Contract.Assert(IsUpdating || IsBoldEditable);
                    SetProperty(ref _isBold, value);
                }
            }

            private bool _isBoldEditable = true;
            public bool IsBoldEditable
            {
                get => _isBoldEditable;
                set => SetProperty(ref _isBoldEditable, value);
            }

            // Contrast ratio

            private double _contrastRatio;
            public double ContrastRatio
            {
                get => _contrastRatio;
                set => SetProperty(ref _contrastRatio, value);
            }

            protected ColorItemBase(
                uint foregroundColorRef,
                uint backgroundColorRef,
                uint autoForegroundColorRef,
                uint autoBackgroundColorRef,
                bool isBold,
                bool isForegroundEditable,
                bool isBackgroundEditable,
                bool isBoldEditable)
            {
                ThrowIfNotOnUIThread();

                _foregroundColorRef = foregroundColorRef;
                _backgroundColorRef = backgroundColorRef;
                _autoForegroundColorRef = autoForegroundColorRef;
                _autoBackgroundColorRef = autoBackgroundColorRef;
                _isBold = isBold;
                _isForegroundEditable = isForegroundEditable;
                _isBackgroundEditable = isBackgroundEditable;
                _isBoldEditable = isBoldEditable;

                ComputeContrastRatio();

                PropertyChanged += (s, o) =>
                {
                    switch (o.PropertyName)
                    {
                        case nameof(Foreground):
                        case nameof(Background):
                            ComputeContrastRatio();
                            break;
                    }
                };
            }

            private void ComputeContrastRatio()
            {
                ThrowIfNotOnUIThread();
                if (!IsForegroundEditable || !IsBackgroundEditable)
                {
                    ContrastRatio = 0.0;
                    return;
                }
                var contrast = ColorHelpers.GetContrast(Foreground, Background);
                ContrastRatio = contrast;
            }
        }
    }
}
