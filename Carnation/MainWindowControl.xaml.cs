﻿using System;
using System.Windows.Controls;
using Microsoft.VisualStudio.Shell;

namespace Carnation
{
    /// <summary>
    /// Interaction logic for MainWindowControl.
    /// </summary>
    public partial class MainWindowControl : UserControl, IDisposable
    {
        private ActiveWindowTracker _activeWindowTracker;
        private readonly MainWindowControlViewModel _viewModel;

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindowControl"/> class.
        /// </summary>
        public MainWindowControl()
        {
            DataContext = _viewModel = new MainWindowControlViewModel();
            InitializeComponent();

            _activeWindowTracker = new ActiveWindowTracker();

            _activeWindowTracker.PropertyChanged += ActiveWindowPropertyChanged;
        }

        private void ActiveWindowPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case nameof(ActiveWindowTracker.SelectedSpan):
                    if (TrackCursor.IsChecked == true)
                    {
                        _viewModel.OnSelectedSpanChanged(_activeWindowTracker.ActiveWpfTextView, _activeWindowTracker.SelectedSpan);
                    }
                    break;

                case nameof(ActiveWindowTracker.ActiveWpfTextView):
                    break;
            }
        }

        public void Dispose()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _activeWindowTracker?.Dispose();
            _activeWindowTracker = null;
        }

        private void ToggleButton_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            _viewModel.SearchText = "";
        }
    }
}
