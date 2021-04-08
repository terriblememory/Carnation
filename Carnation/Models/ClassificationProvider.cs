using System.Collections.Immutable;
using System.Linq;
using System.Windows.Media;
using static Microsoft.VisualStudio.Shell.ThreadHelper;

namespace Carnation
{
    internal partial class ClassificationProvider
    {
        public static Color PlainTextForeground { get; private set; }
        public static Color PlainTextBackground { get; private set; }
        public static ImmutableDictionary<string, string> ClassificationNameMap { get; private set; }
        public static bool IsUpdating { get; private set; }

        public ImmutableArray<ClassificationGridItem> GridItems { get; }

        public ClassificationProvider()
        {
            ThrowIfNotOnUIThread();

            var infos = FontsAndColorsHelper.GetTextEditorInfos();

            GridItems =
                infos.Keys
                .SelectMany(category => infos[category].Select(info => FontsAndColorsHelper.TryGetClassificationItemForInfo(category, info)))
                .OfType<ClassificationGridItem>()
                .ToImmutableArray();
        }

        public void Refresh(ILookup<string, string> definitionNames)
        {
            ThrowIfNotOnUIThread();

            try
            {
                IsUpdating = true;

                var infos = FontsAndColorsHelper.GetTextEditorInfos()
                    .ToImmutableDictionary(
                        kvp => kvp.Key,
                        kvp => kvp.Value
                            .Where(info => definitionNames.Contains(info.bstrName))
                            .ToImmutableDictionary(info => info.bstrName));

                foreach (var item in GridItems.Where(item => definitionNames.Contains(item.DefinitionName)))
                {
                    FontsAndColorsHelper.RefreshClassificationItem(item, infos[item.Category][item.DefinitionName]);
                }
            }
            finally
            {
                IsUpdating = false;
            }
        }
    }
}
