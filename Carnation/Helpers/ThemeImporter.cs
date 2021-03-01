﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Microsoft.VisualStudio.Shell;
using static Carnation.ClassificationProvider;

namespace Carnation.Helpers
{
    internal static class ThemeImporter
    {
        private const string FontsAndColorsCategoryId = "{1EDA5DD4-927A-43a7-810E-7FD247D0DA1D}";
        private const string TextEditorCategoryId = "{A27B4E24-A735-4D1D-B8E7-9716E1E3D8E0}";

        public static void Import(string fileName, IEnumerable<ClassificationGridItem> items)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Import(XDocument.Load(fileName), items);
        }

        public static void Import(XDocument settings, IEnumerable<ClassificationGridItem> items)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var classificationsByCategoryId = items.GroupBy(item => item.Category)
                .ToDictionary(group => group.Key, group => group.ToDictionary(item => item.DefinitionName));

            var allCategories = settings.Descendants("Category");

            var fontsAndColorsCategory = allCategories
                .SingleOrDefault(category => FontsAndColorsCategoryId.Equals(category.Attribute("Category")?.Value, StringComparison.OrdinalIgnoreCase));
            if (fontsAndColorsCategory is null)
            {
                return;
            }

            var fontsAndColorsNode = fontsAndColorsCategory.Descendants("FontsAndColors").SingleOrDefault();
            if (fontsAndColorsNode?.Attribute("Version")?.Value != "2.0")
            {
                return;
            }

            FontsAndColorsHelper.ResetAllClassificationItems();

            var categories = fontsAndColorsNode.Descendants("Category");
            foreach (var category in categories)
            {
                // Check guid
                var guid = category.Attribute("GUID")?.Value;
                if (guid is null)
                {
                    continue;
                }

                if (TextEditorCategoryId.Equals(guid, StringComparison.OrdinalIgnoreCase))
                {
                    ImportTextEditorCategory(category, classificationsByCategoryId);
                }

                if (classificationsByCategoryId.TryGetValue(new Guid(guid), out var classificationsByName))
                {
                    ImportCategory(category, classificationsByName);
                }
            }
        }

        private static void ImportTextEditorCategory(XElement category, Dictionary<Guid, Dictionary<string, ClassificationGridItem>> classificationsByCategoryId)
        {
            foreach (var item in category.Descendants("Item"))
            {
                var name = item.Attribute("Name")?.Value;
                if (name is null)
                {
                    continue;
                }

                foreach (var classificationsByName in classificationsByCategoryId.Values)
                {
                    if (classificationsByName.TryGetValue(name, out var classificationItem))
                    {
                        ImportItem(item, classificationItem);
                        break;
                    }

                }
            }
        }

        private static void ImportCategory(XElement category, Dictionary<string, ClassificationGridItem> classificationsByName)
        {
            foreach (var item in category.Descendants("Item"))
            {
                var name = item.Attribute("Name")?.Value;
                if (name is null ||
                    !classificationsByName.TryGetValue(name, out var classificationItem))
                {
                    continue;
                }

                ImportItem(item, classificationItem);
            }
        }

        private static void ImportItem(XElement item, ClassificationGridItem classificationItem)
        {
            var foreground = item.Attribute("Foreground")?.Value;
            if (classificationItem.IsForegroundEditable &&
                foreground is not null &&
                uint.TryParse(foreground.Substring(2), NumberStyles.HexNumber, provider: null, out var foregroundColorRef))
            {
                classificationItem.ForegroundColorRef = foregroundColorRef;
            }

            var background = item.Attribute("Background")?.Value;
            if (classificationItem.IsBackgroundEditable &&
                background is not null &&
                uint.TryParse(background.Substring(2), NumberStyles.HexNumber, provider: null, out var backgroundColorRef))
            {
                classificationItem.BackgroundColorRef = backgroundColorRef;
            }

            var boldFont = item.Attribute("BoldFont")?.Value;
            if (classificationItem.IsBoldEditable &&
                boldFont is not null)
            {
                classificationItem.IsBold = boldFont.Equals("Yes", StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
