using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class HelloButton : IMehdiRibbonButton
    {
        public string Id => "Hello";

        public string DisplayName => "Hello";
        public string Tooltip => "Show a hello message";
        public string Hint => "Hello";

        public string SmallIconFile => "hello_20.png";
        public string LargeIconFile => "hello_32.png";

        public RibbonSection Section => RibbonSection.General;
        public int SectionOrder => 0;

        public bool IsFreeFeature => true;

        public void Execute(AddinContext context)
        {
            try
            {
                MessageBox.Show("Hello from Mehdi Tools ✨", "SW2026RibbonAddin");
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }

        public int GetEnableState(AddinContext context) => AddinContext.Enable;
    }
}
