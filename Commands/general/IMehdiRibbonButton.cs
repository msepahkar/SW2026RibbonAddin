namespace SW2026RibbonAddin.Commands
{
    /// <summary>
    /// Contract for all ribbon buttons in the add-in.
    /// </summary>
    internal interface IMehdiRibbonButton
    {
        // Internal identifier (must be unique)
        string Id { get; }

        // UI text
        string DisplayName { get; }
        string Tooltip { get; }
        string Hint { get; }

        // Icon filenames (from Resources/)
        string SmallIconFile { get; }
        string LargeIconFile { get; }

        // Layout
        RibbonSection Section { get; }
        int SectionOrder { get; }

        // Licensing
        bool IsFreeFeature { get; }

        // Behavior
        void Execute(AddinContext context);
        int GetEnableState(AddinContext context);
    }
}
