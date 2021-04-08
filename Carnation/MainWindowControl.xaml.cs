using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Controls;
using Microsoft.VisualStudio.Text.Classification;
using static Microsoft.VisualStudio.Shell.ThreadHelper;

namespace Carnation
{
    public partial class MainWindowControl : UserControl, IDisposable
    {
        private ActiveWindowTracker _activeWindowTracker;
        private readonly MainWindowControlViewModel _viewModel;

        public MainWindowControl()
        {
            ThrowIfNotOnUIThread();

            DataContext = _viewModel = new MainWindowControlViewModel();

            InitializeComponent();

            _activeWindowTracker = new ActiveWindowTracker();
            _activeWindowTracker.PropertyChanged += ActiveWindowPropertyChanged;

            var editorFormatMapService = VSServiceHelpers.GetMefExport<IEditorFormatMapService>();
            var editorFormatMap = editorFormatMapService.GetEditorFormatMap("text");

            editorFormatMap.FormatMappingChanged +=
                (object s, FormatItemsEventArgs e) => UpdateClassifications(e.ChangedItems);

            void UpdateClassifications(ReadOnlyCollection<string> definitionNames)
            {
                _viewModel.OnThemeChanged(definitionNames.ToLookup(name => name));
            }
        }

        private void ActiveWindowPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            ThrowIfNotOnUIThread();

            switch (e.PropertyName)
            {
                case nameof(ActiveWindowTracker.SelectedSpan):
                    _viewModel.OnSelectedSpanChanged(_activeWindowTracker.ActiveWpfTextView, _activeWindowTracker.SelectedSpan);
                    break;

                case nameof(ActiveWindowTracker.ActiveWpfTextView):
                    break;
            }
        }

        public void Dispose()
        {
            ThrowIfNotOnUIThread();
            _activeWindowTracker?.Dispose();
            _activeWindowTracker = null;
        }
    }
}
