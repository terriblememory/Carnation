using System;
using Microsoft.VisualStudio.Shell;

namespace Carnation
{
    internal partial class ClassificationProvider
    {
        public class ClassificationGridItem : ColorItemBase
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

            private bool _hasContrastWarning;
            public bool HasContrastWarning
            {
                get => _hasContrastWarning;
                set => SetProperty(ref _hasContrastWarning, value);
            }

            public string Classification => ClassificationNameMap.ContainsKey(DefinitionName)
                ? ClassificationNameMap[DefinitionName]
                : DefinitionName;

            public string Sample => "Sample Text";

            public ClassificationGridItem(
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
                : base(foregroundColorRef,
                       backgroundColorRef,
                       autoForegroundColorRef,
                       autoBackgroundColorRef,
                       isBold,
                       isForegroundEditable,
                       isBackgroundEditable,
                       isBoldEditable)
            {
                _category = category;
                _definitionName = definitionName;
                _definitionLocalizedName = definitionLocalizedName;

                PropertyChanged += (s, o) =>
                {
                    if (IsUpdating)
                    { 
                        return;
                    }

                    ThreadHelper.ThrowIfNotOnUIThread();

                    switch (o.PropertyName)
                    {
                        case nameof(Foreground):
                        case nameof(Background):
                        case nameof(IsBold):
                            FontsAndColorsHelper.SaveClassificationItem(this);
                            break;
                    }
                };
            }

            public override string ToString()
                => Classification;
        }
    }
}
