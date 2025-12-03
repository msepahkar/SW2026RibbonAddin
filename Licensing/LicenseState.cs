namespace SW2026RibbonAddin
{
    /// <summary>
    /// Global license state for the add-in.
    /// For now we map IsLicensed -> Licensed, otherwise -> TrialActive
    /// so existing commands stay fully functional.
    /// </summary>
    internal enum LicenseState
    {
        Unlicensed,
        TrialActive,
        TrialExpired,
        Licensed
    }
}
