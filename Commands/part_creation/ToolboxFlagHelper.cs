using System;
using System.Diagnostics;
using SolidWorks.Interop.swdocumentmgr;

namespace SW2026RibbonAddin
{
    internal static class ToolboxFlagHelper
    {
        // Document Manager key – in your working build you left this as a placeholder,
        // and Document Manager still worked in your environment. Keep it the same.
        private const string DOC_MGR_KEY = "PUT-YOUR-DOC-MANAGER-KEY-HERE";

        // Cache the DM application instance
        private static SwDMApplication _dmApp;

        private static SwDMApplication GetDmApp()
        {
            if (_dmApp != null)
                return _dmApp;

            // Create the class factory COM object
            var factoryType = Type.GetTypeFromProgID("SwDocumentMgr.SwDMClassFactory");
            if (factoryType == null)
            {
                Debug.WriteLine("Document Manager SDK is not installed (SwDMClassFactory type missing).");
                return null;
            }

            var factory = Activator.CreateInstance(factoryType) as SwDMClassFactory;
            if (factory == null)
            {
                Debug.WriteLine("Failed to create SwDMClassFactory.");
                return null;
            }

            try
            {
                _dmApp = factory.GetApplication(DOC_MGR_KEY);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to get Document Manager application: " + ex);
                _dmApp = null;
            }

            return _dmApp;
        }

        /// <summary>
        /// Clears the Toolbox flag (IsToolboxPart) in the file header of the given part.
        /// After this runs successfully, the file behaves as a normal part on any machine.
        /// </summary>
        internal static void ClearToolboxFlagOnDisk(string partPath)
        {
            if (string.IsNullOrWhiteSpace(partPath))
                return;

            var app = GetDmApp();
            if (app == null)
                return;

            SwDmDocumentOpenError openErr;
            SwDMDocument dmDoc = null;

            try
            {
                dmDoc = app.GetDocument(
                    partPath,
                    SwDmDocumentType.swDmDocumentPart,
                    false,    // not read‑only
                    out openErr);

                if (dmDoc == null || openErr != SwDmDocumentOpenError.swDmDocumentOpenErrorNone)
                {
                    Debug.WriteLine($"DM: cannot open '{partPath}', error {openErr}.");
                    return;
                }

                // Already a normal part? Nothing to do.
                if (dmDoc.ToolboxPart == SwDmToolboxPartType.swDmNotAToolboxPart)
                    return;

                // Clear the toolbox flag
                dmDoc.ToolboxPart = SwDmToolboxPartType.swDmNotAToolboxPart;

                // Persist the change in the file header
                dmDoc.Save();   // SwDmDocumentSaveError can be ignored here
            }
            catch (Exception ex)
            {
                Debug.WriteLine("DM: failed to clear Toolbox flag: " + ex);
            }
            finally
            {
                try
                {
                    dmDoc?.CloseDoc();
                }
                catch
                {
                    // ignore close errors
                }
            }
        }
    }
}
