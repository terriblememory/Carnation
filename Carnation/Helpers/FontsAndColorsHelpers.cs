using System;
using System.Collections.Immutable;
using System.Windows.Media;
using Carnation.Helpers;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using static Microsoft.VisualStudio.VSConstants;
using static Carnation.ClassificationProvider;
using static Microsoft.VisualStudio.Shell.ThreadHelper;
using Microsoft;

namespace Carnation
{
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
        /// Static constructor.
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

        public static ImmutableDictionary<Guid, ImmutableArray<AllColorableItemInfo>> GetTextEditorInfos()
        {
            var builder = ImmutableDictionary.CreateBuilder<Guid, ImmutableArray<AllColorableItemInfo>>();
            AppendAllColorableItemInfos(builder, TextEditorCategoryGuid);
            return builder.ToImmutable();
        }

        private static void AppendAllColorableItemInfos(ImmutableDictionary<Guid, ImmutableArray<AllColorableItemInfo>>.Builder dbuilder, Guid category)
        {
            ThrowIfNotOnUIThread();

            Assumes.True(s_fontsAndColorDefaultsProvider.GetObject(category, out var obj) == S_OK);

            if (obj is IVsFontAndColorGroup group)
            {
                var index = 0;
                while (group.GetCategory(index, out var categoryGuid) == S_OK)
                {
                    if (categoryGuid == Guid.Empty) break;
                    AppendAllColorableItemInfos(dbuilder, categoryGuid);
                    ++index;
                }
            }
            else if (obj is IVsFontAndColorDefaults fontAndColorDefaults)
            {
                Assumes.True(fontAndColorDefaults.GetItemCount(out var count) == S_OK);
                var abuilder = ImmutableArray.CreateBuilder<AllColorableItemInfo>();
                var items = new AllColorableItemInfo[1];
                for (var index = 0; index < count; index++)
                {
                    Assumes.True(fontAndColorDefaults.GetItem(index, items) == S_OK);
                    abuilder.Add(items[0]);
                }
                dbuilder.Add(category, abuilder.ToImmutable());
            }
        }

        public static (Color Foreground, Color Background) GetPlainTextColors()
        {
            ThrowIfNotOnUIThread();

            if (s_fontsAndColorStorage.OpenCategory(TextEditorCategoryGuid, (uint)(__FCSTORAGEFLAGS.FCSF_READONLY | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS)) != S_OK)
            {
                return DefaultTextColors;
            }

            try
            {
                var colorItems = new ColorableItemInfo[1];

                if (s_fontsAndColorStorage.GetItem("Plain Text", colorItems) != S_OK)
                {
                    return DefaultTextColors;
                }

                var colorItem = colorItems[0];
                var foreground = TryGetColor(colorItem.crForeground) ?? DefaultTextColors.Foreground;
                var background = TryGetColor(colorItem.crBackground) ?? DefaultTextColors.Background;

                return (foreground, background);
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        public static (FontFamily FontFamily, double FontSize) GetEditorFontInfo(bool scaleFontSize = true)
        {
            ThrowIfNotOnUIThread();

            if (s_fontsAndColorStorage.OpenCategory(TextEditorCategoryGuid, (uint)(__FCSTORAGEFLAGS.FCSF_READONLY | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS)) != S_OK)
            {
                return DefaultFontInfo;
            }

            try
            {
                var logFont = new LOGFONTW[1];
                var fontInfo = new FontInfo[1];

                if (s_fontsAndColorStorage.GetFont(logFont, fontInfo) != S_OK)
                {
                    return DefaultFontInfo;
                }

                var fontFamily = fontInfo[0].bFaceNameValid != 0
                    ? new FontFamily(fontInfo[0].bstrFaceName)
                    : DefaultFontInfo.FontFamily;

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

        public static ClassificationGridItem TryGetClassificationItemForInfo(Guid category, AllColorableItemInfo allColorableItemInfo)
        {
            ThrowIfNotOnUIThread();

            var flags = (uint)(__FCSTORAGEFLAGS.FCSF_READONLY | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS);

            if (s_fontsAndColorStorage.OpenCategory(category, flags) != S_OK)
            {
                // We were unable to access color information.
                return null;
            }

            try
            {
                var definitionName = allColorableItemInfo.bstrName;

                var colorItems = new ColorableItemInfo[1];

                if (s_fontsAndColorStorage.GetItem(definitionName, colorItems) != S_OK)
                {
                    return null;
                }

                var colorItem = colorItems[0];
                var isBold = ((FONTFLAGS)colorItem.dwFontFlags).HasFlag(FONTFLAGS.FF_BOLD);
                var isForegroundEditable = ((__FCITEMFLAGS)allColorableItemInfo.fFlags).HasFlag(__FCITEMFLAGS.FCIF_ALLOWFGCHANGE);
                var isBackgroundEditable = ((__FCITEMFLAGS)allColorableItemInfo.fFlags).HasFlag(__FCITEMFLAGS.FCIF_ALLOWBGCHANGE);
                var isBoldEditable = ((__FCITEMFLAGS)allColorableItemInfo.fFlags).HasFlag(__FCITEMFLAGS.FCIF_ALLOWBOLDCHANGE);

                return new ClassificationGridItem(
                    category,
                    definitionName,
                    allColorableItemInfo.bstrLocalizedName,
                    colorItem.crForeground,
                    colorItem.crBackground,
                    allColorableItemInfo.crAutoForeground,
                    allColorableItemInfo.crAutoBackground,
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

        public static Color? TryGetColor(uint colorRef)
        {
            ThrowIfNotOnUIThread();

            if (s_fontAndColorUtilities.GetColorType(colorRef, out var colorType) != S_OK)
            {
                return null;
            }

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

        internal static void SaveClassificationItem(ClassificationGridItem item)
        {
            ThrowIfNotOnUIThread();

            // Make sure LOADDEFAULTS is passed so any default values can be modified as well.
            if (s_fontsAndColorStorage.OpenCategory(item.Category, (uint)(__FCSTORAGEFLAGS.FCSF_PROPAGATECHANGES | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS)) != S_OK)
            {
                // We were unable to access color information.
                return;
            }

            try
            {
                var colorItems = new ColorableItemInfo[1];
                if (s_fontsAndColorStorage.GetItem(item.DefinitionName, colorItems) != S_OK)
                {
                    return;
                }

                var colorItem = colorItems[0];

                colorItem.crForeground = item.ForegroundColorRef;
                colorItem.crBackground = item.BackgroundColorRef;

                colorItem.dwFontFlags = item.IsBold
                    ? (uint)FONTFLAGS.FF_BOLD
                    : (uint)FONTFLAGS.FF_DEFAULT;

                if (s_fontsAndColorStorage.SetItem(item.DefinitionName, new[] { colorItem }) != S_OK)
                {
                    throw new Exception();
                }
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        internal static uint GetColorRef(Color color, Color defaultColor)
        {
            if (color == defaultColor) return InvalidColorRef;
            return (uint)(color.R | color.G << 8 | color.B << 16);
        }

        internal static void RefreshClassificationItem(ClassificationGridItem item, AllColorableItemInfo allColorableItemInfo)
        {
            ThrowIfNotOnUIThread();

            if (s_fontsAndColorStorage.OpenCategory(item.Category, (uint)(__FCSTORAGEFLAGS.FCSF_READONLY | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS)) != S_OK)
            {
                return;
            }

            try
            {
                var colorItems = new ColorableItemInfo[1];

                if (s_fontsAndColorStorage.GetItem(item.DefinitionName, colorItems) != S_OK)
                {
                    return;
                }

                var colorItem = colorItems[0];

                item.AutoForegroundColorRef = allColorableItemInfo.crAutoForeground;
                item.AutoBackgroundColorRef = allColorableItemInfo.crAutoBackground;
                item.ForegroundColorRef = colorItem.crForeground;
                item.BackgroundColorRef = colorItem.crBackground;
                item.IsBold = ((FONTFLAGS)colorItem.dwFontFlags).HasFlag(FONTFLAGS.FF_BOLD);
            }
            finally
            {
                s_fontsAndColorStorage.CloseCategory();
            }
        }

        internal static void ResetClassificationItem(ClassificationGridItem item)
        {
            ThrowIfNotOnUIThread();

            if (s_fontsAndColorStorage.OpenCategory(item.Category, (uint)(__FCSTORAGEFLAGS.FCSF_PROPAGATECHANGES | __FCSTORAGEFLAGS.FCSF_LOADDEFAULTS)) != S_OK)
            {
                return;
            }

            try
            {
                var colorItems = new ColorableItemInfo[1];

                if (s_fontsAndColorStorage.GetItem(item.DefinitionName, colorItems) != S_OK)
                {
                    return;
                }

                if (((IVsFontAndColorStorage2)s_fontsAndColorStorage).RevertItemToDefault(item.DefinitionName) != S_OK)
                {
                    throw new Exception();
                }

                if (s_fontsAndColorStorage.GetItem(item.DefinitionName, colorItems) != S_OK)
                {
                    return;
                }

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

        internal static void ResetAllClassificationItems()
        {
            ThrowIfNotOnUIThread();

            ResetCategoryItems(TextEditorCategoryGuid);

            return;

            static void ResetCategoryItems(Guid category)
            {
                if (s_fontsAndColorStorage.OpenCategory(category, (uint)(__FCSTORAGEFLAGS.FCSF_PROPAGATECHANGES)) != S_OK)
                {
                    return;
                }

                try
                {
                    if (((IVsFontAndColorStorage2)s_fontsAndColorStorage).RevertAllItemsToDefault() != S_OK)
                    {
                        throw new Exception();
                    }
                }
                finally
                {
                    s_fontsAndColorStorage.CloseCategory();
                }
            }
        }
    }
}
