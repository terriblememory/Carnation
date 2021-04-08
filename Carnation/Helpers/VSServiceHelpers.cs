using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.ComponentModelHost;
using static Microsoft.VisualStudio.VSConstants;
using static Microsoft.VisualStudio.Shell.ThreadHelper;

namespace Carnation
{
    internal static class VSServiceHelpers
    {
        public static Microsoft.VisualStudio.OLE.Interop.IServiceProvider GlobalServiceProvider { get; set; }

        public static TServiceInterface GetMefExport<TServiceInterface>(Microsoft.VisualStudio.OLE.Interop.IServiceProvider serviceProvider = null) where TServiceInterface : class
        {
            serviceProvider = serviceProvider ?? GlobalServiceProvider;
            var componentModel = GetService<IComponentModel, SComponentModel>(serviceProvider);
            if (componentModel == null) return null;
            return componentModel.DefaultExportProvider.GetExportedValue<TServiceInterface>();
        }

        public static IEnumerable<TServiceInterface> GetMefExports<TServiceInterface>(Microsoft.VisualStudio.OLE.Interop.IServiceProvider serviceProvider = null) where TServiceInterface : class
        {
            serviceProvider = serviceProvider ?? GlobalServiceProvider;
            var componentModel = GetService<IComponentModel, SComponentModel>(serviceProvider);
            if (componentModel == null) return null;
            return componentModel.DefaultExportProvider.GetExportedValues<TServiceInterface>();
        }

        public static TServiceInterface GetService<TServiceInterface, TService>(
            Microsoft.VisualStudio.OLE.Interop.IServiceProvider serviceProvider = null)
            where TServiceInterface : class
            where TService : class
        {
            ThrowIfNotOnUIThread();
            serviceProvider = serviceProvider ?? GlobalServiceProvider;
            return (TServiceInterface)GetService(serviceProvider, typeof(TService).GUID, false);
        }

        public static object GetService(
            Microsoft.VisualStudio.OLE.Interop.IServiceProvider serviceProvider, Guid guidService, bool unique)
        {
            ThrowIfNotOnUIThread();

            var guidInterface = IID_IUnknown;
            if (serviceProvider.QueryService(ref guidService, ref guidInterface, out var ptr) != 0) return null;
            if (ptr == IntPtr.Zero) return null;

            try
            {
                if (unique) return Marshal.GetUniqueObjectForIUnknown(ptr);
                else return Marshal.GetObjectForIUnknown(ptr);
            }
            finally
            {
                Marshal.Release(ptr);
            }
        }
    }
}
