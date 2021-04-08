using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft;
using Task = System.Threading.Tasks.Task;
using static Microsoft.VisualStudio.Shell.ThreadHelper;

namespace Carnation
{
    internal sealed class MainWindowCommand
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("a432b46a-d8e0-4439-bd9e-58e40c02453c");
        private readonly AsyncPackage package;

        private MainWindowCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            Assumes.Present(package);
            Assumes.Present(commandService);
            this.package = package;
            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static MainWindowCommand Instance
        {
            get;
            private set;
        }

        public Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider => package;

        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in
            // MainWindowCommand's constructor requires the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new MainWindowCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
           ThrowIfNotOnUIThread();

            // Get the instance number 0 of this tool window. This window is
            // single instance so this instance is actually the only one. The
            // last flag is set to true so that if the tool window does not
            // exists it will be created.
            var window = package.FindToolWindow(typeof(MainWindow), 0, true);
            Assumes.NotNull(window);
            Assumes.NotNull(window.Frame);
            var windowFrame = (IVsWindowFrame)window.Frame;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
    }
}
