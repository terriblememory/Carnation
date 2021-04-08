﻿using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace Carnation
{
    // The minimum requirement for a class to be considered a valid package
    // for Visual Studio is to implement the IVsPackage interface and register
    // itself with the shell.
    //
    // This package uses the helper classes defined inside the Managed Package
    // Framework (MPF) to do it: it derives from the Package class that
    // provides the implementation of the IVsPackage interface and uses the
    // registration attributes defined in the framework to register itself and
    // its components with the shell. These attributes tell the pkgdef creation
    // utility what data to put into .pkgdef file.
    //
    // To get loaded into VS, the package must be referred by
    // <Asset Type="Microsoft.VisualStudio.VsPackage" ...> in .vsixmanifest file.

    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(CarnationPackage.PackageGuidString)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideToolWindow(typeof(MainWindow))]
    public sealed class CarnationPackage : AsyncPackage
    {
        /// CarnationPackage GUID string.
        public const string PackageGuidString = "2cc0a490-70eb-4dd6-a16e-a29b8f3d273c";

        // Initialization of the package; this method is called right after the
        // package is sited, so this is the place where you can put all the
        // initialization code that relies on services provided by VisualStudio.
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a
            // background thread at this point. Do any initialization that
            // requires the UI thread after switching to the UI thread.
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await MainWindowCommand.InitializeAsync(this);
        }
    }
}
