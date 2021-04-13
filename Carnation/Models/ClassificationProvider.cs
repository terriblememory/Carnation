using System.Collections.Immutable;
using System.Linq;
using System.Windows.Media;
using static Microsoft.VisualStudio.Shell.ThreadHelper;

namespace Carnation
{
    public class ClassificationProvider
    {
        public static Color PlainTextForeground { get; private set; }
        public static Color PlainTextBackground { get; private set; }
        public static bool IsUpdating { get; private set; }

        static ClassificationProvider()
        {
            ThrowIfNotOnUIThread();
            (PlainTextForeground, PlainTextBackground) = FontsAndColorsHelper.GetPlainTextColors();
        }

        public ImmutableArray<GridItem> GridItems { get; }

        public ClassificationProvider()
        {
            ThrowIfNotOnUIThread();
            var builder = ImmutableArray.CreateBuilder<GridItem>();
            var items = FontsAndColorsHelper.GetColorItems();
            foreach (var item in items)
            {
                var ci = FontsAndColorsHelper.TryGetGridItemForColorItem(item);
                builder.Add(ci);
            }
            GridItems = builder.ToImmutable();
        }

        public void Refresh(ILookup<string, string> definitionNames)
        {
            ThrowIfNotOnUIThread();

            try
            {
                IsUpdating = true;

                var colorItems = FontsAndColorsHelper.GetColorItems();

                foreach (var classificationItem in GridItems)
                {
                    if (definitionNames.Contains(classificationItem.DefinitionName))
                    {
                        foreach (var colorItem in colorItems)
                        {
                            if (colorItem.AllColorableItemInfo.bstrName == classificationItem.DefinitionName)
                            {
                                FontsAndColorsHelper.RefreshGridItemFromColorItem(classificationItem, colorItem);
                            }
                        }
                    }
                }
            }
            finally
            {
                IsUpdating = false;
            }
        }
    }
}
