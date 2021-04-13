using System;
using System.Diagnostics.Contracts;
using System.Windows.Media;
using Microsoft.VisualStudio.Shell;
using static Microsoft.VisualStudio.Shell.ThreadHelper;

namespace Carnation
{
    public class GridItem : NotifyPropertyBase
    {
        private Guid _category;
        public Guid Category
        {
            get => _category;
            set => SetProperty(ref _category, value);
        }

        private string _definitionName;
        public string DefinitionName
        {
            get => _definitionName;
            set => SetProperty(ref _definitionName, value);
        }

        private string _definitionLocalizedName;
        public string DefinitionLocalizedName
        {
            get => _definitionLocalizedName;
            set => SetProperty(ref _definitionLocalizedName, value);
        }

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
                return FontsAndColorsHelper.TryGetColorFromColorRef(ForegroundColorRef) ?? DefaultForeground;
            }
            set
            {
                ThrowIfNotOnUIThread();
                Contract.Assert(ClassificationProvider.IsUpdating || IsForegroundEditable);
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
                return FontsAndColorsHelper.TryGetColorFromColorRef(AutoForegroundColorRef) ?? ClassificationProvider.PlainTextForeground;
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
                return FontsAndColorsHelper.TryGetColorFromColorRef(BackgroundColorRef) ?? DefaultBackground;
            }
            set
            {
                ThrowIfNotOnUIThread();
                Contract.Assert(ClassificationProvider.IsUpdating || IsBackgroundEditable);
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
                return FontsAndColorsHelper.TryGetColorFromColorRef(AutoBackgroundColorRef) ?? ClassificationProvider.PlainTextBackground;
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
                Contract.Assert(ClassificationProvider.IsUpdating || IsBoldEditable);
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

        private bool _hasContrastWarning;
        public bool HasContrastWarning
        {
            get => _hasContrastWarning;
            set => SetProperty(ref _hasContrastWarning, value);
        }

        // Classification

        public string Classification => ClassificationMap.GetClassificationNameForItemName(DefinitionName);

        public override string ToString()
            => Classification;

        public string Sample => "Sample Text";

        public GridItem(
            Guid category,
            string definitionName,
            string definitionLocalizedName,
            uint foregroundColorRef,
            uint backgroundColorRef,
            uint autoForegroundColorRef,
            uint autoBackgroundColorRef,
            bool isBold,
            bool isForegroundEditable,
            bool isBackgroundEditable,
            bool isBoldEditable)
        {
            _category = category;
            _definitionName = definitionName;
            _definitionLocalizedName = definitionLocalizedName;
            _foregroundColorRef = foregroundColorRef;
            _backgroundColorRef = backgroundColorRef;
            _autoForegroundColorRef = autoForegroundColorRef;
            _autoBackgroundColorRef = autoBackgroundColorRef;
            _isBold = isBold;
            _isForegroundEditable = isForegroundEditable;
            _isBackgroundEditable = isBackgroundEditable;
            _isBoldEditable = isBoldEditable;

            PropertyChanged += (s, o) =>
            {
                if (ClassificationProvider.IsUpdating)
                {
                    return;
                }

                ThreadHelper.ThrowIfNotOnUIThread();

                switch (o.PropertyName)
                {
                    case nameof(Foreground):
                    case nameof(Background):
                    case nameof(IsBold):
                        FontsAndColorsHelper.SaveGridItem(this);
                        break;
                }
            };
        }
    }
}
