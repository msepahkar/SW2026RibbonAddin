using System;
using System.IO;

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
        // Not heavily used now because every job has its own sheet dims,
        // but we keep it for logging and future features.
        public SheetPreset DefaultSheet { get; set; }

        // These are ALWAYS enabled now (your request).
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

        // Compatibility alias (if any older file uses this name)
        public bool SeparateByMaterial
        {
            get => SeparateByMaterialExact;
            set => SeparateByMaterialExact = value;
        }
    }

    internal sealed class LaserNestJob
    {
        public bool Enabled { get; set; } = true;

        public string ThicknessFilePath { get; set; }
        public double ThicknessMm { get; set; }

        // EXACT material string from SolidWorks (stored in block name)
        public string MaterialExact { get; set; }

        // Sheet for THIS job (material + thickness)
        public SheetPreset Sheet { get; set; }

        public string ThicknessFileName => Path.GetFileName(ThicknessFilePath) ?? "";
    }

    /// <summary>
    /// Progress sink used by the nesting engine.
    ///
    /// IMPORTANT: Nesting currently runs on the UI thread in this add-in.
    /// Progress implementations should keep the UI responsive (e.g., by
    /// pumping the message loop after updates).
    /// </summary>
    internal interface ILaserCutProgress
    {
        void BeginBatch(int totalTasks);

        void BeginTask(
            int taskIndex,
            int totalTasks,
            LaserNestJob job,
            int totalParts,
            NestingMode mode,
            double sheetWmm,
            double sheetHmm);

        void ReportPlaced(int placed, int total, int sheetsUsed);

        void EndTask(
            int doneTasks,
            int totalTasks,
            LaserNestJob job,
            bool success,
            string message);

        void SetStatus(string message);

        void ThrowIfCancelled();
    }
}
