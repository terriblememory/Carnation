using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using static Carnation.ClassificationProvider;
using static Microsoft.VisualStudio.Shell.ThreadHelper;

namespace Carnation.Helpers
{
    internal static class ThemeImporter
    {
        private const string FontsAndColorsCategoryId = "{1EDA5DD4-927A-43a7-810E-7FD247D0DA1D}";

        public static void Import(string fileName, IEnumerable<GridItem> items)
        {
            ThrowIfNotOnUIThread();

            Import(XDocument.Load(fileName), items);
        }

        public static void Import(XDocument settings, IEnumerable<GridItem> items)
        {
            ThrowIfNotOnUIThread();

            var classificationsByCategoryId =
                items.GroupBy(item => item.Category)
                .ToDictionary(group => group.Key, group => group.ToDictionary(item => item.DefinitionName));

            var allCategories = settings.Descendants("Category");

            var fontsAndColorsCategory =
                allCategories
                .SingleOrDefault(category => category.Attribute("Category")?.Value == FontsAndColorsCategoryId);

            if (fontsAndColorsCategory is null)
            {
                return;
            }

            var fontsAndColorsNode = fontsAndColorsCategory.Descendants("FontsAndColors").SingleOrDefault();

            if (fontsAndColorsNode?.Attribute("Version")?.Value != "2.0")
            {
                return;
            }

            FontsAndColorsHelper.ResetAllGridItems();

            var categories = fontsAndColorsNode.Descendants("Category");

            foreach (var category in categories)
            {
                // Check guid
                var guid = category.Attribute("GUID")?.Value;
                if (guid is null) continue;
                if (!classificationsByCategoryId.TryGetValue(new Guid(guid), out var classificationsByName)) continue;

                foreach (var item in category.Descendants("Item"))
                {
                    // Check name
                    var name = item.Attribute("Name")?.Value;
                    if (name is null) continue;
                    if (!classificationsByName.TryGetValue(name, out var classificationItem)) continue;

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
    }
}

/*

<UserSettings>

    <ApplicationIdentity version="16.0"/>

    <ToolsOptions>
        <ToolsOptionsCategory name="Environment" RegisteredName="Environment"/>
    </ToolsOptions>

    <!-- Level 1 Category -->
    <Category name="Environment_Group" RegisteredName="Environment_Group">

        <!-- Level 2 Category -->
        <Category
            name="Environment_FontsAndColors"
            Category="{1EDA5DD4-927A-43a7-810E-7FD247D0DA1D}"
            Package="{DA9FB551-C724-11d0-AE1F-00A0C90FFFC3}"
            RegisteredName="Environment_FontsAndColors"
            PackageName="Visual Studio Environment Package">

            <PropertyValue name="Version">2</PropertyValue>
            <FontsAndColors Version="2.0">
                <Theme Id="{297B4000-C797-45C1-A76F-6BCDE49C72D9}"/>
                <Categories>

                    <!-- Level 3 Category -->
                    <Category GUID="{FA937F7B-C0D2-46B8-9F10-A7A92642B384}" FontIsDefault="Yes">
                        <Items>
                            <Item Name="Artboard Background" Foreground="0x02000000" Background="0x02000000"/>
                        </Items>
                    </Category>

                    <Category GUID="{B36B0228-DBAD-4DB0-B9C7-2AD3E572010F}" FontName="Segoe UI" FontSize="9" CharSet="1" FontIsDefault="No">
                        <Items>
                            <Item Name="Odd Row Items" Foreground="0x00000000" Background="0x00FFFFFF" BoldFont="No"/>
                            <Item Name="Even Row Items" Foreground="0x00000000" Background="0x00FFFFFF" BoldFont="No"/>
                            <!-- etc. -->
                        </Items>
                    </Category>

                    <!-- TextEditorMEFItemsCategory -->
                    <Category GUID="{75A05685-00A8-4DED-BAE5-E7A50BFA929A}" FontName="Consolas_4.0" FontSize="11" CharSet="1" FontIsDefault="No">
                        <Items>
                            <Item Name="MarkerFormatDefinition/ScopeHighlight" Foreground="0x00F7EBE7" Background="0x0000FF00" BoldFont="No"/>
                            <Item Name="CppStringDelimiterCharacterSyntacticTokenFormat" Foreground="0x00FFFFFF" Background="0x01000001" BoldFont="No"/>
                            <Item Name="CppControlKeywordSyntacticTokenFormat" Foreground="0x00FFFFFF" Background="0x01000001" BoldFont="No"/>
                            <Item Name="C/C++ User Keywords" Foreground="0x00FFFFFF" Background="0x02000000" BoldFont="No"/>
                            <Item Name="String" Foreground="0x00FFFFFF" Background="0x01000001" BoldFont="No"/>
                            <!-- etc. -->
                        </Items>
                    </Category>

                    <!-- TextEditorLanguageServiceCategory -->
                    <Category GUID="{E0187991-B458-4F7E-8CA9-42C9A573B56C}" FontName="Consolas_4.0" FontSize="11" CharSet="1" FontIsDefault="No">
                        <Items>
                            <Item Name="Number" Foreground="0x02000000" Background="0x02000000" BoldFont="No"/>
                            <Item Name="String" Foreground="0x00FFFFFF" Background="0x02000000" BoldFont="No"/>
                            <!-- etc. -->
                        </Items>
                    </Category>

                    <!-- TextEditorManagerCategory -->
                    <Category GUID="{58E96763-1D3B-4E05-B6BA-FF7115FD0B7B}" FontName="Consolas_4.0" FontSize="11" CharSet="1" FontIsDefault="No">
                        <Items>
                            <Item Name="Indicator Margin" Foreground="0x01000017" Background="0x00221813" BoldFont="No"/>
                            <Item Name="Plain Text" Foreground="0x00FFFFFF" Background="0x00221813" BoldFont="No"/>
                        </Items>
                    </Category>

                    <!-- TextEditorMarkerCategory -->
                    <Category GUID="{FF349800-EA43-46C1-8C98-878E78F46501}" FontName="Consolas_4.0" FontSize="11" CharSet="1" FontIsDefault="No">
                        <Items>
                            <Item Name="Breakpoint (Enabled)" Foreground="0x00FFFFFF" Background="0x000000FF" BoldFont="No"/>
                            <Item Name="Breakpoint (Warning)" Foreground="0x00FFFFFF" Background="0x000000FF" BoldFont="No"/>
                            <!-- etc. -->
                        </Items>
                    </Category>

                </Categories>
            </FontsAndColors>
        </Category>
    </Category>
</UserSettings> 

*/
