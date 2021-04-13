using System.Collections.Generic;
using System.Collections.Immutable;
using System.Linq;
using System.Reflection;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Language.StandardClassification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;

namespace Carnation
{
    public static class ClassificationMap
    {
        private static ImmutableDictionary<string, string> itemNameToClassification;

        static ClassificationMap()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var builder = ImmutableDictionary.CreateBuilder<string, string>();
            var efds = VSServiceHelpers.GetMefExports<EditorFormatDefinition>().ToArray();
            foreach (var efd in efds)
            {
                var type = efd.GetType();
                var uv = type.GetCustomAttribute<UserVisibleAttribute>();
                if (uv?.UserVisible != true) continue;
                var name = type.GetCustomAttribute<NameAttribute>()?.Name;
                var ctns = type.GetCustomAttribute<ClassificationTypeAttribute>()?.ClassificationTypeNames;
                if (string.IsNullOrEmpty(name)) continue;
                builder.Add(name, ctns ?? name);
            }

            itemNameToClassification = builder.ToImmutable();
        }

        public static string GetClassificationNameForItemName(string itemName)
        {
            return itemNameToClassification.TryGetValue(itemName, out var classificationName) ? classificationName : itemName;
        }
    }

    internal static class ClassificationHelpers
    {
        public static ImmutableDictionary<string, string> GetClassificationNameMap()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // This results in a dictionary where the keys are font/color item
            // names and the values are VS editor format classification names.

            var builder = ImmutableDictionary.CreateBuilder<string, string>();
            var efds = VSServiceHelpers.GetMefExports<EditorFormatDefinition>().ToArray();
            foreach (var efd in efds)
            {
                var type = efd.GetType();
                var uv = type.GetCustomAttribute<UserVisibleAttribute>();
                if (uv?.UserVisible != true) continue;
                var name = type.GetCustomAttribute<NameAttribute>()?.Name;
                var ctns = type.GetCustomAttribute<ClassificationTypeAttribute>()?.ClassificationTypeNames;
                if (string.IsNullOrEmpty(name)) continue;
                builder.Add(name, ctns ?? name);
            }

            return builder.ToImmutable();
        }

        public static ImmutableArray<string> GetClassificationsForSpan(IWpfTextView view, Span span)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // If the cursor is after the last character return empty array.
            if (span.Start == view.TextBuffer.CurrentSnapshot.Length)
                return ImmutableArray<string>.Empty;

            // The span to classify must have a length.
            var snapshotSpan = span.Length == 0
                ? new SnapshotSpan(view.TextSnapshot, span.Start, 1)
                : new SnapshotSpan(view.TextSnapshot, span);

            // Get the classifications that apply to the span.
            var classificationService = VSServiceHelpers.GetMefExport<IClassifierAggregatorService>();
            var classifier = (IAccurateClassifier)classificationService.GetClassifier(view.TextBuffer);
            var classifiedSpans = classifier.GetAllClassificationSpans(snapshotSpan, CancellationToken.None);
            var builder = ImmutableArray.CreateBuilder<string>();
            foreach (var cs in classifiedSpans) CollectClassifications(cs.ClassificationType, builder);
            return builder.ToImmutable();
        }

        // There's no define for the transient classification. Oh well.
        private static readonly HashSet<string> s_ignoredClassifications = new HashSet<string>
        {
            "(TRANSIENT)",
            PredefinedClassificationTypeNames.FormalLanguage
        };

        private static void CollectClassifications(IClassificationType classificationType, ImmutableArray<string>.Builder builder)
        {
            var baseClassificationTypes = classificationType.BaseTypes.ToArray();

            if (baseClassificationTypes.Length > 1)
            {
                // This is a compound classification - break it into its base
                // classifications and process those.
                foreach (var baseClassificationType in baseClassificationTypes)
                {
                    CollectClassifications(baseClassificationType, builder);
                }
            }
            else
            {
                // If this is one of the ignored classifications (we ignore
                // some because they're not useful for our purposes and would
                // show up all the time) don't add it.
                if (s_ignoredClassifications.Contains(classificationType.Classification)) return;

                // If we've already added it (because it was part of some other
                // classification) don't add it.
                if (builder.Contains(classificationType.Classification)) return;

                // Add the classification.
                builder.Add(classificationType.Classification);
            }
        }
    }
}
