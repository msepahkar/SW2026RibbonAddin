using System;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin
{
    internal static class ToolboxFlagHelper
    {
        /// <summary>
        /// Marks the currently-open model as NOT a Toolbox part, then saves silently.
        /// This is enough for most workflows to remove the "bolt" behavior/icon.
        /// </summary>
        internal static bool TryMarkAsNotToolboxAndSave(IModelDoc2 model, out string reason)
        {
            reason = null;

            if (model == null)
            {
                reason = "No active model.";
                return false;
            }

            try
            {
                var ext = model.Extension as ModelDocExtension;
                if (ext == null)
                {
                    reason = "ModelDocExtension is not available.";
                    return false;
                }

                // Clear toolbox status in the active document
                ext.ToolboxPartType = (int)swToolBoxPartType_e.swNotAToolboxPart;

                // Save to persist it
                int saveErrors = 0;
                int saveWarnings = 0;

                bool ok = model.Save3(
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    ref saveErrors,
                    ref saveWarnings);

                if (!ok || saveErrors != 0)
                {
                    reason = $"Save failed. Error code: {saveErrors}.";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("TryMarkAsNotToolboxAndSave error: " + ex);
                reason = ex.Message;
                return false;
            }
        }
    }
}
