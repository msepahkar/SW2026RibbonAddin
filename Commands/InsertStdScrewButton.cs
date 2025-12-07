using System;
using System.Diagnostics;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace SW2026RibbonAddin.Commands
{
    /// <summary>
    /// Ribbon button that drives the STD screw insertion logic.
    /// </summary>
    internal sealed class InsertStdScrewButton : IMehdiRibbonButton
    {
        public string Id => "InsertStdScrew";

        public string DisplayName => "STD\nScrew";
        public string Tooltip => "Insert a standard screw from the STD folder into the selected hole stack(s).";
        public string Hint => "Insert screw into selected hole(s)";

        // Icon files to place under Resources/
        public string SmallIconFile => "std_screw_20.png";
        public string LargeIconFile => "std_screw_32.png";

        // Put it in the General section of the ribbon
        public RibbonSection Section => RibbonSection.General;
        public int SectionOrder => 5;

        // Treat as non‑free (paid) feature for future licensing
        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            try
            {
                StdScrewInserter.Run(context);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
                MessageBox.Show(
                    "Error while inserting STD screw:\r\n\r\n" + ex.Message,
                    "Insert STD Screw",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            try
            {
                IModelDoc2 model = context.ActiveModel;
                if (model == null)
                    return AddinContext.Disable;

                // Only meaningful in assemblies
                return model is IAssemblyDoc
                    ? AddinContext.Enable
                    : AddinContext.Disable;
            }
            catch
            {
                return AddinContext.Disable;
            }
        }
    }
}
