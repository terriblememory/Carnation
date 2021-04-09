using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Windows.Media;
using Microsoft;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using static Microsoft.VisualStudio.VSConstants;
using static Microsoft.VisualStudio.Shell.ThreadHelper;
using Carnation.Helpers;

namespace Carnation
{
    public class ColorItem
    {
        internal Guid Category;
        internal string CategoryName;
        internal int Priority;
        internal AllColorableItemInfo AllColorableItemInfo;
    }

    internal static class FontsAndColorsHelper
    {
        internal static readonly Guid TextEditorPackageGuid = new Guid("daf27b38-80b3-4c58-8133-afd41c36c79a");
        internal static readonly Guid TextEditorCategoryGuid = new Guid(FontsAndColorsCategory.TextEditor);

        internal static readonly (FontFamily FontFamily, double FontSize) DefaultFontInfo = (new FontFamily("Consolas"), 13.0);
        internal static readonly (Color Foreground, Color Background) DefaultTextColors = (Colors.Black, Colors.White);

        private static readonly IVsFontAndColorStorage s_fontsAndColorStorage;
        private static readonly IVsFontAndColorUtilities s_fontAndColorUtilities;
        private static readonly IVsFontAndColorDefaultsProvider s_fontsAndColorDefaultsProvider;
        private static readonly IVsUIShell2 s_vsUIShell2;

        private const uint InvalidColorRef = 0xff000000;

        /// <summary>
        /// Static constructor - get services, etc.
        /// </summary>
        static FontsAndColorsHelper()
        {
            ThrowIfNotOnUIThread();

            s_fontsAndColorStorage = ServiceProvider.GlobalProvider.GetService<SVsFontAndColorStorage, IVsFontAndColorStorage>();
            Assumes.Present(s_fontsAndColorStorage);

            s_fontAndColorUtilities = (IVsFontAndColorUtilities)s_fontsAndColorStorage;
            Assumes.Present(s_fontAndColorUtilities);

            s_vsUIShell2 = ServiceProvider.GlobalProvider.GetService<SVsUIShell, IVsUIShell2>();
            Assumes.Present(s_vsUIShell2);

            s_fontsAndColorDefaultsProvider = (IVsFontAndColorDefaultsProvider)ServiceProvider.GlobalProvider.GetService(TextEditorPackageGuid);
            Assumes.Present(s_fontsAndColorDefaultsProvider);
        }

        /// <summary>
        /// Get the complete list of ColorItems.
        /// </summary>
        /// <remarks>
        /// This list is basically a union of subcategories with duplicates
        /// removed, where a duplicate is considered to have the same non-
        /// localized name. The item from the group with the highest priority
        /// is retained. I believe this is the behavior expected by VS based
        /// on observation, but I don't see it clearly documented anywhere.
        /// </remarks>
        public static ImmutableArray<ColorItem> GetColorItems()
        {
            ThrowIfNotOnUIThread();
            var items = new Dictionary<string, ColorItem>();
            AppendColorItemsForCategory(items, TextEditorCategoryGuid);
            var builder = ImmutableArray.CreateBuilder<ColorItem>();
            builder.AddRange(items.Values);
            builder.Sort((lhs, rhs) => string.Compare(lhs.AllColorableItemInfo.bstrName, rhs.AllColorableItemInfo.bstrName));
            return builder.ToImmutable();
        }

        /// <summary>
        /// Helper - append the ColorItems for a category/subcategory to the builder.
        /// </summary>
        private static void AppendColorItemsForCategory(Dictionary<string, ColorItem> items, Guid category)
        {
            ThrowIfNotOnUIThread();
            Assumes.True(s_fontsAndColorDefaultsProvider.GetObject(category, out var obj) == S_OK);
            if (obj is IVsFontAndColorGroup group)
            {
                var index = 0;
                while (group.GetCategory(index, out var subcategory) == S_OK)
                {
                    if (subcategory == Guid.Empty) break;
                    AppendColorItemsForCategory(items, subcategory);
                    ++index;
                }
            }
            if (obj is IVsFontAndColorDefaults fontAndColorDefaults)
            {
                fontAndColorDefaults.GetCategoryName(out var categoryName);
                fontAndColorDefaults.GetPriority(out var priority);
                Assumes.True(fontAndColorDefaults.GetItemCount(out var count) == S_OK);
                var item = new AllColorableItemInfo[1];
                for (var index = 0; index < count; index++)
                {
                    Assumes.True(fontAndColorDefaults.GetItem(index, item) == S_OK);
                    var ci = new ColorItem() { Category = category, CategoryName = categoryName, Priority = priority, AllColorableItemInfo = item[0] };
                    if (items.TryGetValue(item[0].bstrName, out var existing) && existing.Priority < ci.Priority) continue;
                    items[item[0].bstrName] = ci;
                }
            }
        }

        /// <summary>
        /// Get a GridItem for a ColorItem.
        /// </summary>
        public static GridItem TryGetGridItemForColorItem(ColorItem colorItem)
        {
            ThrowIfNotOnUIThread();

            var openFlags = (uint)(__FCSTORAGEFLAGS.FCSF_READONLY | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS);
            Assumes.True(s_fontsAndColorStorage.OpenCategory(colorItem.Category, openFlags) == S_OK);

            try
            {
                var definitionName = colorItem.AllColorableItemInfo.bstrName;
                var colorableItemInfos = new ColorableItemInfo[1];
                Assumes.True(s_fontsAndColorStorage.GetItem(definitionName, colorableItemInfos) == S_OK);
                var colorItemInfo = colorableItemInfos[0];
                var isBold = ((FONTFLAGS)colorItemInfo.dwFontFlags).HasFlag(FONTFLAGS.FF_BOLD);
                var flags = (__FCITEMFLAGS)colorItem.AllColorableItemInfo.fFlags;
                var isForegroundEditable = flags.HasFlag(__FCITEMFLAGS.FCIF_ALLOWFGCHANGE);
                var isBackgroundEditable = flags.HasFlag(__FCITEMFLAGS.FCIF_ALLOWBGCHANGE);
                var isBoldEditable = flags.HasFlag(__FCITEMFLAGS.FCIF_ALLOWBOLDCHANGE);

                return new GridItem(
                    colorItem.Category,
                    definitionName,
                    $"{colorItem.AllColorableItemInfo.bstrLocalizedName} ({definitionName} from {colorItem.CategoryName})",
                    colorItemInfo.crForeground,
                    colorItemInfo.crBackground,
                    colorItem.AllColorableItemInfo.crAutoForeground,
                    colorItem.AllColorableItemInfo.crAutoBackground,
                    isBold,
                    isForegroundEditable,
                    isBackgroundEditable,
                    isBoldEditable);
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        /// <summary>
        /// Update a GridItem from a ColorItem.
        /// </summary>
        internal static void RefreshGridItemFromColorItem(GridItem gridItem, ColorItem colorItem)
        {
            ThrowIfNotOnUIThread();

            var openFlags = (uint)(__FCSTORAGEFLAGS.FCSF_READONLY | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS);
            Assumes.True(s_fontsAndColorStorage.OpenCategory(gridItem.Category, openFlags) == S_OK);

            try
            {
                var colorableItemInfos = new ColorableItemInfo[1];
                Assumes.True(s_fontsAndColorStorage.GetItem(gridItem.DefinitionName, colorableItemInfos) == S_OK);
                var colorableItemInfo = colorableItemInfos[0];
                gridItem.AutoForegroundColorRef = colorItem.AllColorableItemInfo.crAutoForeground;
                gridItem.AutoBackgroundColorRef = colorItem.AllColorableItemInfo.crAutoBackground;
                gridItem.ForegroundColorRef = colorableItemInfo.crForeground;
                gridItem.BackgroundColorRef = colorableItemInfo.crBackground;
                gridItem.IsBold = ((FONTFLAGS)colorableItemInfo.dwFontFlags).HasFlag(FONTFLAGS.FF_BOLD);
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        /// <summary>
        /// Save a GridItem.
        /// </summary>
        internal static void SaveGridItem(GridItem item)
        {
            ThrowIfNotOnUIThread();

            var openFlags = (uint)(__FCSTORAGEFLAGS.FCSF_PROPAGATECHANGES | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS);
            Assumes.True(s_fontsAndColorStorage.OpenCategory(item.Category, openFlags) == S_OK);

            try
            {
                var colorItems = new ColorableItemInfo[1];
                Assumes.True(s_fontsAndColorStorage.GetItem(item.DefinitionName, colorItems) == S_OK);
                var colorItem = colorItems[0];
                colorItem.crForeground = item.ForegroundColorRef;
                colorItem.crBackground = item.BackgroundColorRef;
                colorItem.dwFontFlags = item.IsBold ? (uint)FONTFLAGS.FF_BOLD : (uint)FONTFLAGS.FF_DEFAULT;
                Assumes.True(s_fontsAndColorStorage.SetItem(item.DefinitionName, new[] { colorItem }) == S_OK);
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        /// <summary>
        /// Reset a GridItem to its defaults.
        /// </summary>
        internal static void ResetGridItem(GridItem item)
        {
            ThrowIfNotOnUIThread();

            var openFlags = (uint)(__FCSTORAGEFLAGS.FCSF_PROPAGATECHANGES | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS);
            Assumes.True(s_fontsAndColorStorage.OpenCategory(item.Category, openFlags) == S_OK);

            try
            {
                var colorItems = new ColorableItemInfo[1];
                Assumes.True(s_fontsAndColorStorage.GetItem(item.DefinitionName, colorItems) == S_OK);
                Assumes.True(((IVsFontAndColorStorage2)s_fontsAndColorStorage).RevertItemToDefault(item.DefinitionName) == S_OK);
                Assumes.True(s_fontsAndColorStorage.GetItem(item.DefinitionName, colorItems) == S_OK);
                var colorItem = colorItems[0];
                if (item.IsForegroundEditable) item.ForegroundColorRef = colorItem.crForeground;
                if (item.IsBackgroundEditable) item.BackgroundColorRef = colorItem.crBackground;
                if (item.IsBoldEditable) item.IsBold = ((FONTFLAGS)colorItem.dwFontFlags).HasFlag(FONTFLAGS.FF_BOLD);
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        /// <summary>
        /// Reset all TextEditor GridItems to defaults.
        /// </summary>
        internal static void ResetAllGridItems()
        {
            ThrowIfNotOnUIThread();

            ResetCategoryItems(TextEditorCategoryGuid);
        }

        /// <summary>
        /// Reset GridItems for a category to defaults.
        /// </summary>
        internal static void ResetCategoryItems(Guid category)
        {
            ThrowIfNotOnUIThread();

            var openFlags = (uint)(__FCSTORAGEFLAGS.FCSF_PROPAGATECHANGES);
            Assumes.True(s_fontsAndColorStorage.OpenCategory(category, openFlags) == S_OK);

            try
            {
                Assumes.True(((IVsFontAndColorStorage2)s_fontsAndColorStorage).RevertAllItemsToDefault() == S_OK);
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        /// <summary>
        /// Returns the Text Editor plain text colors.
        /// </summary>
        public static (Color Foreground, Color Background) GetPlainTextColors()
        {
            ThrowIfNotOnUIThread();

            var openFlags = (uint)(__FCSTORAGEFLAGS.FCSF_READONLY | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS);
            Assumes.True(s_fontsAndColorStorage.OpenCategory(TextEditorCategoryGuid, openFlags) == S_OK);

            try
            {
                var colorItems = new ColorableItemInfo[1];
                Assumes.True(s_fontsAndColorStorage.GetItem("Plain Text", colorItems) == S_OK);
                var colorItem = colorItems[0];
                var foreground = TryGetColorFromColorRef(colorItem.crForeground) ?? DefaultTextColors.Foreground;
                var background = TryGetColorFromColorRef(colorItem.crBackground) ?? DefaultTextColors.Background;
                return (foreground, background);
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        /// <summary>
        /// Returns the Text Editor font information.
        /// </summary>
        public static (FontFamily FontFamily, double FontSize) GetEditorFontInfo(bool scaleFontSize = true)
        {
            ThrowIfNotOnUIThread();

            var openFlags = (uint)(__FCSTORAGEFLAGS.FCSF_READONLY | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS);
            Assumes.True(s_fontsAndColorStorage.OpenCategory(TextEditorCategoryGuid, openFlags) == S_OK);

            try
            {
                var logFont = new LOGFONTW[1];
                var fontInfo = new FontInfo[1];
                Assumes.True(s_fontsAndColorStorage.GetFont(logFont, fontInfo) == S_OK);
                var fontFamily = fontInfo[0].bFaceNameValid != 0 ? new FontFamily(fontInfo[0].bstrFaceName) : DefaultFontInfo.FontFamily;
                var fontSize = DefaultFontInfo.FontSize;

                if (fontInfo[0].bPointSizeValid != 0)
                {
                    if (scaleFontSize)
                    {
                        fontSize = Math.Abs(logFont[0].lfHeight) / GetDipsPerPixel();
                    }
                    else
                    {
                        fontSize = fontInfo[0].wPointSize;
                    }
                }

                return (fontFamily, fontSize);
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        private static double GetDipsPerPixel()
        {
            var dc = UnsafeNativeMethods.GetDC(IntPtr.Zero);

            if (dc != IntPtr.Zero)
            {
                // Getting the DPI from the desktop is bad, but some callers just have no context for what monitor they are on.
                double fallbackDpi = UnsafeNativeMethods.GetDeviceCaps(dc, UnsafeNativeMethods.LOGPIXELSX);
                UnsafeNativeMethods.ReleaseDC(IntPtr.Zero, dc);
                return fallbackDpi / 96.0;
            }

            return 1;
        }

        public static Color? TryGetColorFromColorRef(uint colorRef)
        {
            ThrowIfNotOnUIThread();

            Assumes.True(s_fontAndColorUtilities.GetColorType(colorRef, out var colorType) == S_OK);

            uint? win32Color = null;

            if (colorType == (int)__VSCOLORTYPE.CT_INVALID)
            {
                return null;
            }
            else if (colorType == (int)__VSCOLORTYPE.CT_AUTOMATIC)
            {
                return null; 
            }
            else if (colorType == (int)__VSCOLORTYPE.CT_RAW)
            {
                win32Color = colorRef;
            }
            else if (colorType == (int)__VSCOLORTYPE.CT_COLORINDEX)
            {
                var encodedIndex = new COLORINDEX[1];
                if (s_fontAndColorUtilities.GetEncodedIndex(colorRef, encodedIndex) == S_OK &&
                    s_fontAndColorUtilities.GetRGBOfIndex(encodedIndex[0], out var decoded) == S_OK)
                {
                    if (encodedIndex[0] == COLORINDEX.CI_SYSTEXT_BK ||
                        encodedIndex[0] == COLORINDEX.CI_SYSTEXT_FG)
                    {
                        return null;
                    }

                    win32Color = encodedIndex[0] == COLORINDEX.CI_USERTEXT_BK
                        ? decoded & 0x00ffffff
                        : decoded | 0xff000000;
                }
            }
            else if (colorType == (int)__VSCOLORTYPE.CT_SYSCOLOR)
            {
                if (s_fontAndColorUtilities.GetEncodedSysColor(colorRef, out var sysColor) == S_OK)
                {
                    win32Color = (uint)sysColor;
                }
            }
            else if (colorType == (int)__VSCOLORTYPE.CT_VSCOLOR)
            {
                if (s_fontAndColorUtilities.GetEncodedVSColor(colorRef, out var vsSysColor) == S_OK &&
                    s_vsUIShell2.GetVSSysColorEx(vsSysColor, out var rgbColor) == S_OK)
                {
                    win32Color = rgbColor;
                }
            }

            if (!win32Color.HasValue) return null;
            var drawingColor = System.Drawing.ColorTranslator.FromWin32((int)win32Color.Value);
            return Color.FromRgb(drawingColor.R, drawingColor.G, drawingColor.B);
        }

        internal static uint GetColorRef(Color color, Color defaultColor)
        {
            if (color == defaultColor) return InvalidColorRef;
            return (uint)(color.R | color.G << 8 | color.B << 16);
        }
    }
}
