using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace Carnation
{
    [Guid("c1c74746-265a-4718-a5ab-63963d38e8bb")]
    public class MainWindow : ToolWindowPane
    {
        public MainWindow() : base(null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            VSServiceHelpers.GlobalServiceProvider =
                (Microsoft.VisualStudio.OLE.Interop.IServiceProvider)
                    Microsoft.VisualStudio.Shell.Package.GetGlobalService(
                        typeof(Microsoft.VisualStudio.OLE.Interop.IServiceProvider));

            Caption = "Carnation";

            // This is the user control hosted by the tool window; Note that,
            // even if this class implements IDisposable, we are not calling
            // Dispose on this object. This is because ToolWindowPane calls
            // Dispose on the object returned by the Content property.
            Content = new MainWindowControl();
        }
    }
}
