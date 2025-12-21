namespace SW2026RibbonAddin.Commands
{
    internal enum NestingMode
    {
        FastRectangles = 0,
        ContourLevel1 = 1,
        ContourLevel2_NFP = 2,
    }

    internal readonly struct SheetPreset
    {
        public string Name { get; }
        public double WidthMm { get; }
        public double HeightMm { get; }

        public SheetPreset(string name, double wMm, double hMm)
        {
            Name = name ?? "";
            WidthMm = wMm;
            HeightMm = hMm;
        }

        public override string ToString() => $"{Name} ({WidthMm:0.###} x {HeightMm:0.###} mm)";
    }

    internal sealed class LaserCutRunSettings
    {
        public SheetPreset DefaultSheet { get; set; }

        // Material behavior (exact SolidWorks material string written into the block name by DWG export)
        public bool SeparateByMaterialExact { get; set; } = true;
        public bool OutputOneDwgPerMaterial { get; set; } = true;
        public bool KeepOnlyCurrentMaterialInSourcePreview { get; set; } = true;

        // Mode
        public NestingMode Mode { get; set; } = NestingMode.ContourLevel1;

        // Contour extraction tuning (mm)
        public double ContourChordMm { get; set; } = 0.8;
        public double ContourSnapMm { get; set; } = 0.05;

        // Performance guards
        public int MaxCandidatesPerTry { get; set; } = 7000;
        public int MaxNfpPartnersPerTry { get; set; } = 80;

        // Optional compatibility alias (if any of your other files used this older name)
        public bool SeparateByMaterial
        {
            get => SeparateByMaterialExact;
            set => SeparateByMaterialExact = value;
        }
    }
}
