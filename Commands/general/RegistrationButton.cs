using System;
using System.Diagnostics;
using SW2026RibbonAddin.Licensing;

namespace SW2026RibbonAddin.Commands
{
    /// <summary>
    /// Opens registration UI:
    ///  - If activated: show a status dialog.
    ///  - If not activated: show the registration form.
    /// </summary>
    internal sealed class RegistrationButton : IMehdiRibbonButton
    {
        public string Id => "Registration";

        public string DisplayName => "Register";
        public string Tooltip => "View license status or activate this add-in";
        public string Hint => "Registration";

        public string SmallIconFile => "license_20.png";
        public string LargeIconFile => "license_32.png";

        public RibbonSection Section => RibbonSection.General;
        public int SectionOrder => 1;

        // Registration UI must always be available.
        public bool IsFreeFeature => true;

        public void Execute(AddinContext context)
        {
            try
            {
                LicensingUI.ShowRegistrationOrStatusDialog();
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        public int GetEnableState(AddinContext context) => AddinContext.Enable;
    }
}
