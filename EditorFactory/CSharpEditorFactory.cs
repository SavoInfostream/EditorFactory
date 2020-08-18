using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using IServiceProvider = Microsoft.VisualStudio.OLE.Interop.IServiceProvider;

namespace EditorFactory
{
    class CSharpEditorFactory : IVsEditorFactory
    {
        private Package _package;
        private ServiceProvider _serviceProvider;
        private readonly bool _promptEncodingOnLoad;

        public CSharpEditorFactory(Package package)
        {
            _package = package;
        }

        public int CreateEditorInstance(
            uint createEditorFlags,
            string documentMoniker,
            string physicalView,
            IVsHierarchy hierarchy,
            uint itemid,
            System.IntPtr docDataExisting,
            out System.IntPtr docView,
            out System.IntPtr docData,
            out string editorCaption,
            out Guid commandUIGuid,
            out int createDocumentWindowFlags)
        {
            // Initialize output parameters
            docView = IntPtr.Zero;
            docData = IntPtr.Zero;
            commandUIGuid = Guid.Empty;
            createDocumentWindowFlags = 0;
            editorCaption = null;

            // Validate inputs
            if ((createEditorFlags & (VSConstants.CEF_OPENFILE | VSConstants.CEF_SILENT)) == 0)
            {
                return VSConstants.E_INVALIDARG;
            }

            if (_promptEncodingOnLoad && docDataExisting != IntPtr.Zero)
            {
                return VSConstants.VS_E_INCOMPATIBLEDOCDATA;
            }

            IVsTextLines textLines = GetTextBuffer(documentMoniker); 

            docData = Marshal.GetIUnknownForObject(textLines);

            editorCaption = String.Empty;
            commandUIGuid = Guid.Empty;
            docView = CreateCodeView(documentMoniker, textLines, ref editorCaption, ref commandUIGuid);

            return VSConstants.S_OK;
        }

        private IVsTextLines GetTextBuffer(string documentMoniker)
        {

            var componentModel = (IComponentModel)_serviceProvider.GetService(typeof(SComponentModel));
            var AdapterService = componentModel.GetService<IVsEditorAdaptersFactoryService>();
            var contentTypeRegistry = componentModel.GetService<IContentTypeRegistryService>();

            var serviceProvider = (IServiceProvider) _serviceProvider.GetService(typeof(IServiceProvider));

            var fileContent = File.ReadAllText(documentMoniker);

            var textLines = AdapterService.CreateVsTextBufferAdapter(serviceProvider, contentTypeRegistry.GetContentType("csharp"));

            ErrorHandler.ThrowOnFailure(
                textLines.InitializeContent(fileContent.Substring(0, 200), 200)
            );

            return textLines as IVsTextLines;
        }

        protected virtual IntPtr CreateCodeView(string documentMoniker, IVsTextLines textLines, ref string editorCaption, ref Guid cmdUI)
        {

            Type codeWindowType = typeof(IVsCodeWindow);
            Guid riid = codeWindowType.GUID;
            Guid clsid = typeof(VsCodeWindowClass).GUID;
            IVsCodeWindow window = (IVsCodeWindow)_package.CreateInstance(ref clsid, ref riid, codeWindowType);
            ErrorHandler.ThrowOnFailure(window.SetBuffer(textLines));
            ErrorHandler.ThrowOnFailure(window.SetBaseEditorCaption(null));
            ErrorHandler.ThrowOnFailure(window.GetEditorCaption(READONLYSTATUS.ROSTATUS_Unknown, out editorCaption));

            IVsUserData userData = textLines as IVsUserData;
            if (userData != null)
            {
                if (_promptEncodingOnLoad)
                {
                    var guid = VSConstants.VsTextBufferUserDataGuid.VsBufferEncodingPromptOnLoad_guid;
                    userData.SetData(ref guid, (uint)1);
                }
            }

            cmdUI = VSConstants.GUID_TextEditorFactory;
            return Marshal.GetIUnknownForObject(window);
        }


        public int SetSite(IServiceProvider psp)
        {
            _serviceProvider = new ServiceProvider(psp);
            return VSConstants.S_OK;
        }

        public int Close()
        {
            return VSConstants.S_OK;
        }

        public int MapLogicalView(ref Guid logicalView, out string physicalView)
        {
            // initialize out parameter
            physicalView = null;

            bool isSupportedView = false;
            // Determine the physical view
            if (VSConstants.LOGVIEWID_Primary == logicalView ||
                VSConstants.LOGVIEWID_Debugging == logicalView ||
                VSConstants.LOGVIEWID_Code == logicalView ||
                VSConstants.LOGVIEWID_TextView == logicalView)
            {
                // primary view uses NULL as pbstrPhysicalView
                isSupportedView = true;
            }
            else if (VSConstants.LOGVIEWID_Designer == logicalView)
            {
                physicalView = "Design";
                isSupportedView = true;
            }

            if (isSupportedView)
                return VSConstants.S_OK;
            else
            {
                // E_NOTIMPL must be returned for any unrecognized rguidLogicalView values
                return VSConstants.E_NOTIMPL;
            }
        }
    }
}
