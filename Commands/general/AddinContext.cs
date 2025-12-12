using SolidWorks.Interop.sldworks;

namespace SW2026RibbonAddin
{
    /// <summary>
    /// Context passed to ribbon button implementations.
    /// Gives access to the add-in and the SolidWorks app.
    /// </summary>
    public sealed class AddinContext
    {
        // SolidWorks enable/disable values
        public const int Enable = 1;
        public const int Disable = 0;

        public AddinContext(Addin addin, SldWorks swApp)
        {
            Addin = addin;
            SwApp = swApp;
        }

        public Addin Addin { get; }
        public SldWorks SwApp { get; }

        public IModelDoc2 ActiveModel => SwApp?.IActiveDoc2 as IModelDoc2;
    }
}
