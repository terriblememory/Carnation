using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using Carnation.Helpers;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.Utilities;
using Microsoft.Win32;
using static Carnation.ClassificationProvider;
using static Microsoft.VisualStudio.Shell.ThreadHelper;

namespace Carnation
{
    internal class MainWindowControlViewModel : NotifyPropertyBase
    {
        public MainWindowControlViewModel()
        {
            ThrowIfNotOnUIThread();

            var ss = ThreadHelper.JoinableTaskFactory.Run(() => OptionsHelper.GetWritableSettingsStoreAsync());

            if (ss.TryGetBoolean(
                OptionsHelper.GeneralSettingsCollectionName,
                nameof(UseExtraContrastSuggestions),
                out var useExtraContrastSuggestions))
            {
                UseExtraContrastSuggestions = useExtraContrastSuggestions;
            }

            PropertyChanged += OnPropertyChanged;

            EditForegroundCommand = new RelayCommand<GridItem>(OnEditForeground);
            EditBackgroundCommand = new RelayCommand<GridItem>(OnEditBackground);
            ToggleIsBoldCommand = new RelayCommand<GridItem>(OnToggleIsBold);
            ResetToDefaultsCommand = new RelayCommand<GridItem>(OnResetToDefaults);
            UseForegroundSuggestionCommand = new RelayCommand<GridItem>(OnUseForegroundSuggestion);
            ResetAllToDefaultsCommand = new RelayCommand(OnResetAllToDefaults);
            UseAllForegroundSuggestionsCommand = new RelayCommand(OnUseAllForegroundSuggestions);
            ExportThemeCommand = new RelayCommand(OnExportTheme);
            ImportThemeCommand = new RelayCommand(OnImportTheme);

            foreach (var item in ClassificationProvider.GridItems) ClassificationGridItems.Add(item);
            UpdateContrastWarnings();

            ClassificationGridView = CollectionViewSource.GetDefaultView(ClassificationGridItems);
            ClassificationGridView.Filter = o => FilterClassification((GridItem)o);
            ClassificationGridView.SortDescriptions.Clear();
            ClassificationGridView.SortDescriptions.Add(new SortDescription(nameof(GridItem.Classification), ListSortDirection.Ascending));

            (FontFamily, FontSize) = FontsAndColorsHelper.GetEditorFontInfo();
        }

        public ClassificationProvider ClassificationProvider { get; } = new ClassificationProvider();
        public ObservableCollection<GridItem> ClassificationGridItems { get; } = new ObservableCollection<GridItem>();

        private GridItem _selectedClassification;
        public GridItem SelectedClassification
        {
            get => _selectedClassification;
            set => SetProperty(ref _selectedClassification, value);
        }

        private bool _followCursorSelected;
        public bool FollowCursorSelected
        {
            get => _followCursorSelected;
            set => SetProperty(ref _followCursorSelected, value);
        }

        private string _searchText;
        public string SearchText
        {
            get => _searchText;
            set => SetProperty(ref _searchText, value);
        }

        private bool _searchTextEnabled;
        public bool SearchTextEnabled
        {
            get => _searchTextEnabled;
            set => SetProperty(ref _searchTextEnabled, value);
        }

        private FontFamily _fontFamily;
        public FontFamily FontFamily
        {
            get => _fontFamily;
            set => SetProperty(ref _fontFamily, value);
        }

        private double _fontSize;
        public double FontSize
        {
            get => _fontSize;
            set => SetProperty(ref _fontSize, value);
        }

        private bool _useExtraContrastSuggestions;
        public bool UseExtraContrastSuggestions
        {
            get => _useExtraContrastSuggestions;
            set => SetProperty(ref _useExtraContrastSuggestions, value);
        }

        public ICollectionView ClassificationGridView { get; }

        public Color SelectedItemForeground
        {
            get
            {
                ThrowIfNotOnUIThread();
                if (SelectedClassification == null) return Colors.Transparent;
                return SelectedClassification.Foreground;
            }
            set
            {
                ThrowIfNotOnUIThread();
                if (SelectedClassification.Foreground != value) return;
                SelectedClassification.Foreground = value;
                NotifyPropertyChanged();
            }
        }

        public Color SelectedItemBackground
        {
            get
            {
                ThrowIfNotOnUIThread();
                if (SelectedClassification == null) return Colors.Transparent;
                return SelectedClassification.Background;
            }
            set
            {
                ThrowIfNotOnUIThread();
                if (SelectedClassification.Background == value) return;
                SelectedClassification.Background = value;
                NotifyPropertyChanged();
            }
        }

        public ICommand EditForegroundCommand { get; }
        public ICommand EditBackgroundCommand { get; }
        public ICommand ToggleIsBoldCommand { get; }
        public ICommand ResetToDefaultsCommand { get; }
        public ICommand UseForegroundSuggestionCommand { get; }
        public ICommand ResetAllToDefaultsCommand { get; }
        public ICommand UseAllForegroundSuggestionsCommand { get; }
        public ICommand ExportThemeCommand { get; }
        public ICommand ImportThemeCommand { get; }

        public void OnThemeChanged(ILookup<string, string> definitionNames)
        {
            ThrowIfNotOnUIThread();

            if (definitionNames.Contains("Plain Text"))
            {
                (FontFamily, FontSize) = FontsAndColorsHelper.GetEditorFontInfo();
            }
            
            ClassificationProvider.Refresh(definitionNames);

            UpdateContrastWarnings(definitionNames);
        }

        public void OnSelectedSpanChanged(IWpfTextView view, Span? span)
        {
            ThrowIfNotOnUIThread();

            if (span is null || view is null)
            {
                return;
            }

            if (!FollowCursorSelected)
            {
                return;
            }

            var classifications = ClassificationHelpers.GetClassificationsForSpan(view, span.Value);

            SearchText = string.Join("; ", classifications);
        }

        private void UpdateContrastWarnings(ILookup<string, string> definitionNames = null)
        {
            var minimumContrastRatio = UseExtraContrastSuggestions
                ? ContrastHelpers.AAA_Contrast
                : ContrastHelpers.AA_Contrast;

            foreach (var item in ClassificationGridItems)
            {
                if (definitionNames?.Contains(item.DefinitionName) == false)
                {
                    continue;
                }

                item.HasContrastWarning = item.IsForegroundEditable
                    && item.IsBackgroundEditable
                    && item.ContrastRatio < minimumContrastRatio;
            }
        }

        private bool FilterClassification(GridItem item)
        {
            if (string.IsNullOrEmpty(SearchText)) return true;

            if (item is null) return false;

            if (FollowCursorSelected)
            {
                var classifications = SearchText.Split(';')
                    .Select(name => name.Trim())
                    .Where(name => !string.IsNullOrEmpty(name))
                    .ToLookup(name => name);

                // If follow cursor is selected, we only want exact matches
                return classifications.Contains(item.Classification);
            }

            return item.Classification.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) != -1;
        }

        private void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case nameof(FollowCursorSelected):
                    OnFollowCursorChanged();
                    break;

                case nameof(SearchText):
                    OnSearchTextChanged();
                    break;

                case nameof(SelectedClassification):
                    NotifyPropertyChanged(nameof(SelectedItemBackground));
                    NotifyPropertyChanged(nameof(SelectedItemForeground));
                    break;

                case nameof(UseExtraContrastSuggestions):
                    ThreadHelper.JoinableTaskFactory.Run(UpdateUseExtraContrastOptionAsync);
                    break;
            }
        }

        private async System.Threading.Tasks.Task UpdateUseExtraContrastOptionAsync()
        {
            var ss = await OptionsHelper.GetWritableSettingsStoreAsync();
            ss.WriteBoolean(OptionsHelper.GeneralSettingsCollectionName, nameof(UseExtraContrastSuggestions), UseExtraContrastSuggestions);
            UpdateContrastWarnings();
        }

        private void OnSearchTextChanged()
        {
            ClassificationGridView.Refresh();
        }

        private void OnFollowCursorChanged()
        {
            SearchTextEnabled = !FollowCursorSelected;
            if (FollowCursorSelected) SearchText = string.Empty;
        }

        private void OnResetToDefaults(GridItem item)
        {
            ThrowIfNotOnUIThread();
            FontsAndColorsHelper.ResetGridItem(item);
        }

        private void OnUseForegroundSuggestion(GridItem item)
        {
            ThrowIfNotOnUIThread();

            var suggestions = UseExtraContrastSuggestions
                ? ContrastHelpers.FindSimilarAAAColor(item.Foreground, item.Background)
                : ContrastHelpers.FindSimilarAAColor(item.Foreground, item.Background);

            if (suggestions.Length == 0)
            {
                item.HasContrastWarning = false;
                return;
            }

            var topSuggestion = suggestions.OrderBy(suggestion => suggestion.Distance).First();

            item.Foreground = topSuggestion.Color;
        }

        private void OnToggleIsBold(GridItem item)
        {
            item.IsBold = !item.IsBold;
        }

        private void OnEditForeground(GridItem item)
        {
            ThrowIfNotOnUIThread();
            ShowColorPicker(item);
        }

        private void OnEditBackground(GridItem item)
        {
            ThrowIfNotOnUIThread();
            ShowColorPicker(item, true);
        }

        private void ShowColorPicker(GridItem item, bool editBackground = false)
        {
            ThrowIfNotOnUIThread();

            var window = new ColorPickerWindow(
                item.Foreground,
                item.Background,
                UseExtraContrastSuggestions,
                FontFamily,
                FontSize,
                editBackground: editBackground);

            if (window.ShowDialog() == true)
            {
                if (editBackground) item.Background = window.BackgroundColor;
                else item.Foreground = window.ForegroundColor;
            }
        }

        private void OnResetAllToDefaults()
        {
            ThrowIfNotOnUIThread();
            FontsAndColorsHelper.ResetAllGridItems();
            UpdateContrastWarnings();
        }

        private void OnExportTheme()
        {
            ThrowIfNotOnUIThread();

            var dialog = new SaveFileDialog
            {
                DefaultExt = "vssettings",
                Title = "Export Theme",
                Filter = "Settings Files (*.vssettings)|*.vssettings|All Files (*.*)|*.*",
                AddExtension = true,
                OverwritePrompt = true,
                CheckPathExists = true
            };

            if (dialog.ShowDialog() == true)
            {
                ThemeExporter.Export(dialog.FileName, ClassificationGridItems);
            }
        }

        private void OnImportTheme()
        {
            ThrowIfNotOnUIThread();

            var dialog = new OpenFileDialog
            {
                DefaultExt = "vssettings",
                Title = "Import Theme",
                Filter = "Settings Files (*.vssettings)|*.vssettings|All Files (*.*)|*.*",
                AddExtension = true,
                CheckPathExists = true,
                CheckFileExists = true,
                Multiselect = false, 
            };

            if (dialog.ShowDialog() == true)
            {
                var operationExecutor = VSServiceHelpers.GetMefExport<IUIThreadOperationExecutor>();
                operationExecutor.Execute(
                    "Carnation",
                    "Loading theme colors...",
                    allowCancellation: false,
                    showProgress: true,
                    (context) =>
                    {
                        ThemeImporter.Import(dialog.FileName, ClassificationGridItems);
                    });
            }
        }

        private void OnUseAllForegroundSuggestions()
        {
            var operationExecutor = VSServiceHelpers.GetMefExport<IUIThreadOperationExecutor>();
            operationExecutor.Execute(
                "Carnation",
                "Applying all foreground color suggestions...",
                allowCancellation: false,
                showProgress: true,
                (context) =>
                {
                    ThrowIfNotOnUIThread();
                    foreach (var item in ClassificationGridItems)
                    {
                        if (!item.HasContrastWarning) continue;
                        var suggestions = UseExtraContrastSuggestions
                            ? ContrastHelpers.FindSimilarAAAColor(item.Foreground, item.Background)
                            : ContrastHelpers.FindSimilarAAColor(item.Foreground, item.Background);
                        if (suggestions.Length == 0)
                        {
                            item.HasContrastWarning = false;
                            continue;
                        }

                        var topSuggestion = suggestions.OrderBy(suggestion => suggestion.Distance).First();
                        item.Foreground = topSuggestion.Color;
                    }
                });
        }
    }
}
