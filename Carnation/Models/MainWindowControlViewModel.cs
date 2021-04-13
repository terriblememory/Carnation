using System;
using System.Collections;
using System.Collections.Generic;
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

            EditForegroundCommand = new RelayCommand<IList>(OnEditForeground);
            EditBackgroundCommand = new RelayCommand<IList>(OnEditBackground);
            ToggleIsBoldCommand = new RelayCommand<IList>(OnToggleIsBold);
            ResetToDefaultsCommand = new RelayCommand<IList>(OnResetToDefaults);
            ResetAllToDefaultsCommand = new RelayCommand(OnResetAllToDefaults);

            foreach (var item in ClassificationProvider.GridItems) ClassificationGridItems.Add(item);

            ClassificationGridView = CollectionViewSource.GetDefaultView(ClassificationGridItems);
            ClassificationGridView.Filter = o => FilterClassification((GridItem)o);
            ClassificationGridView.SortDescriptions.Clear();
            ClassificationGridView.SortDescriptions.Add(new SortDescription(nameof(GridItem.DefinitionLocalizedName), ListSortDirection.Ascending));

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

        private bool _searchTextEnabled = true;
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

        private ILookup<string, string> SearchClassifications;

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
        public ICommand ResetAllToDefaultsCommand { get; }

        public void OnThemeChanged(ILookup<string, string> definitionNames)
        {
            ThrowIfNotOnUIThread();

            if (definitionNames.Contains("Plain Text"))
            {
                (FontFamily, FontSize) = FontsAndColorsHelper.GetEditorFontInfo();
            }
            
            ClassificationProvider.Refresh(definitionNames);
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

        private bool FilterClassification(GridItem item)
        {
            if (item is null) return false;

            if (FollowCursorSelected)
            {
                // If the classification list is empty just show the plain text colors.
                if (SearchClassifications.Count == 0 && item.Classification == "Plain Text") return true;

                // If follow cursor is selected, we only want exact matches.
                return SearchClassifications.Contains(item.Classification);
            }
            else
            {
                // If not following cursor and the search box is empty include everything.
                if (string.IsNullOrEmpty(SearchText)) return true;

                // But if there's anything in the search box filter on a simple string search.
                return item.DefinitionLocalizedName.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) != -1;
            }
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
            }
        }

        private void OnSearchTextChanged()
        {
            SearchClassifications = SearchText.Split(';')
                .Select(name => name.Trim())
                .Where(name => !string.IsNullOrEmpty(name))
                .ToLookup(name => name);

            ClassificationGridView.Refresh();
        }

        private void OnFollowCursorChanged()
        {
            SearchTextEnabled = !FollowCursorSelected;
            SearchText = string.Empty;
            ClassificationGridView.Refresh();
        }

        private void OnResetToDefaults(IList items)
        {
            ThrowIfNotOnUIThread();
            var list = items.Cast<GridItem>();
            foreach (var item in list)
            {
                FontsAndColorsHelper.ResetGridItem(item);
            }
        }

        private void OnToggleIsBold(IList items)
        {
            var list = items.Cast<GridItem>();
            var first = list.First();
            first.IsBold = !first.IsBold;
            foreach (var item in list)
            {
                if (item.IsBoldEditable)
                    item.IsBold = first.IsBold;
            }
        }

        private void OnEditForeground(IList items)
        {
            ThrowIfNotOnUIThread();

            var list = items.Cast<GridItem>();
            var first = list.First();

            var window = new ColorPickerWindow(
                first.Foreground,
                first.Background,
                UseExtraContrastSuggestions,
                FontFamily,
                FontSize,
                editBackground: false);

            if (window.ShowDialog() == true)
            {
                foreach (var item in list)
                {
                    if (item.IsForegroundEditable)
                        item.Foreground = window.ForegroundColor;
                }
            }
        }

        private void OnEditBackground(IList items)
        {
            ThrowIfNotOnUIThread();

            var list = items.Cast<GridItem>();
            var first = list.First();

            var window = new ColorPickerWindow(
                first.Foreground,
                first.Background,
                UseExtraContrastSuggestions,
                FontFamily,
                FontSize,
                editBackground: true);

            if (window.ShowDialog() == true)
            {
                foreach (var item in list)
                {
                    if (item.IsBackgroundEditable)
                        item.Background = window.BackgroundColor;
                }
            }
        }

        private void OnResetAllToDefaults()
        {
            ThrowIfNotOnUIThread();
            FontsAndColorsHelper.ResetAllGridItems();
        }
    }
}
