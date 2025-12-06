using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using CSMath;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class LaserCutButton : IMehdiRibbonButton
    {
        public string Id => "LaserCut";

        public string DisplayName => "Laser\nCut";
        public string Tooltip => "Nest a combined DWG into laser sheets.";
        public string Hint => "Laser cut nesting";

        // Use your own icons here if you prefer.
        public string SmallIconFile => "laser_cut_20.png";
        public string LargeIconFile => "laser_cut_32.png";

        public RibbonSection Section => RibbonSection.DWG;
        public int SectionOrder => 3;

        public bool IsFreeFeature => false;

        public void Execute(AddinContext context)
        {
            string dwgPath = SelectCombinedDwg();
            if (string.IsNullOrEmpty(dwgPath))
                return;

            double sheetWidth;
            double sheetHeight;

            using (var dlg = new LaserCutOptionsForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK)
                    return;

                sheetWidth = dlg.SheetWidthMm;
                sheetHeight = dlg.SheetHeightMm;
            }

            try
            {
                DwgLaserNester.Nest(dwgPath, sheetWidth, sheetHeight);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Laser cut nesting failed:\r\n\r\n" + ex.Message,
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public int GetEnableState(AddinContext context)
        {
            // Independent of active SW document
            return AddinContext.Enable;
        }

        private static string SelectCombinedDwg()
        {
            using (var dlg = new OpenFileDialog())
            {
                dlg.Title = "Select combined thickness DWG";
                dlg.Filter = "DWG files (*.dwg)|*.dwg|All files (*.*)|*.*";
                dlg.CheckFileExists = true;
                dlg.Multiselect = false;

                if (dlg.ShowDialog() != DialogResult.OK)
                    return null;

                return dlg.FileName;
            }
        }
    }

    internal static class DwgLaserNester
    {
        private sealed class PartDefinition
        {
            public BlockRecord Block;
            public string BlockName;
            public double MinX;
            public double MinY;
            public double Width;
            public double Height;
            public int Quantity;
        }

        private sealed class SheetState
        {
            public int Index;
            public double OriginX;
            public double OriginY;
            public double CurrentX;
            public double CurrentY;
            public double RowHeight;
        }

        /// <summary>
        /// Reads a combined DWG (one produced by CombineDwg) and nests all plate blocks
        /// P_*_Q{qty} onto as many sheets as required.
        /// </summary>
        public static void Nest(string sourceDwgPath, double sheetWidth, double sheetHeight)
        {
            if (sheetWidth <= 0 || sheetHeight <= 0)
                throw new ArgumentException("Sheet width and height must be positive.");

            if (!File.Exists(sourceDwgPath))
                throw new FileNotFoundException("DWG file not found.", sourceDwgPath);

            CadDocument doc;
            using (var reader = new DwgReader(sourceDwgPath))
            {
                doc = reader.Read();
            }

            var parts = LoadPartDefinitions(doc).ToList();
            if (parts.Count == 0)
            {
                MessageBox.Show(
                    "No plate blocks were found in the selected DWG.\r\n" +
                    "Make sure it is one of the combined thickness DWGs.",
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            int totalInstances = parts.Sum(p => p.Quantity);
            if (totalInstances <= 0)
            {
                MessageBox.Show(
                    "All parts in the selected DWG have zero quantity.",
                    "Laser Cut",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            // Validate that every part fits inside a sheet with margins.
            const double sheetMargin = 5.0;
            double usableWidth = sheetWidth - 2 * sheetMargin;
            double usableHeight = sheetHeight - 2 * sheetMargin;

            foreach (var p in parts)
            {
                if (p.Width > usableWidth || p.Height > usableHeight)
                {
                    throw new InvalidOperationException(
                        $"Part '{p.BlockName}' ({p.Width:0.##} x {p.Height:0.##} mm) " +
                        $"does not fit inside sheet {sheetWidth:0.##} x {sheetHeight:0.##} mm.");
                }
            }

            // Build list of instances (copies) to place.
            var instances = new List<PartDefinition>();
            foreach (var def in parts)
            {
                for (int i = 0; i < def.Quantity; i++)
                    instances.Add(def);
            }

            // Largest parts first (by height, then width).
            instances.Sort((a, b) =>
            {
                int cmp = b.Height.CompareTo(a.Height);
                if (cmp != 0) return cmp;
                return b.Width.CompareTo(a.Width);
            });

            // Wipe existing model-space entities; we will rebuild layout for sheets only.
            doc.Entities.Clear();

            var modelSpace = doc.BlockRecords["*Model_Space"];

            const double partGap = 5.0;      // distance between plates
            const double sheetGap = 50.0;    // distance between sheet rectangles

            var sheets = new List<SheetState>();

            SheetState NewSheet()
            {
                var sheet = new SheetState
                {
                    Index = sheets.Count + 1,
                    OriginX = sheets.Count * (sheetWidth + sheetGap),
                    OriginY = 0.0,
                    CurrentX = sheetMargin,
                    CurrentY = sheetMargin,
                    RowHeight = 0.0
                };

                sheets.Add(sheet);

                // Draw sheet boundary as four lines.
                var bottom = new Line
                {
                    StartPoint = new XYZ(sheet.OriginX, sheet.OriginY, 0.0),
                    EndPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY, 0.0)
                };
                var right = new Line
                {
                    StartPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY, 0.0),
                    EndPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY + sheetHeight, 0.0)
                };
                var top = new Line
                {
                    StartPoint = new XYZ(sheet.OriginX + sheetWidth, sheet.OriginY + sheetHeight, 0.0),
                    EndPoint = new XYZ(sheet.OriginX, sheet.OriginY + sheetHeight, 0.0)
                };
                var left = new Line
                {
                    StartPoint = new XYZ(sheet.OriginX, sheet.OriginY + sheetHeight, 0.0),
                    EndPoint = new XYZ(sheet.OriginX, sheet.OriginY, 0.0)
                };

                modelSpace.Entities.Add(bottom);
                modelSpace.Entities.Add(right);
                modelSpace.Entities.Add(top);
                modelSpace.Entities.Add(left);

                // Add "SHEET n" text near upper-left corner.
                var txt = new MText
                {
                    Value = $"SHEET {sheet.Index}",
                    InsertPoint = new XYZ(sheet.OriginX + sheetMargin,
                                          sheet.OriginY + sheetHeight - 20.0, 0.0),
                    Height = 15.0
                };
                modelSpace.Entities.Add(txt);

                return sheet;
            }

            var progress = new LaserCutProgressForm(totalInstances)
            {
                Text = "Laser cut nesting"
            };

            int placed = 0;
            int sheetCount;
            try
            {
                progress.Show();
                Application.DoEvents();

                var sheet = NewSheet();

                foreach (var inst in instances)
                {
                    while (true)
                    {
                        // Is there room in current row?
                        if (sheet.CurrentX + inst.Width <= sheetWidth - sheetMargin)
                        {
                            // Place at current position.
                            double localX = sheet.CurrentX;
                            double localY = sheet.CurrentY;

                            double worldX = sheet.OriginX + localX;
                            double worldY = sheet.OriginY + localY;

                            // Adjust for block's local min corner so bounding box is inside the sheet.
                            double insertX = worldX - inst.MinX;
                            double insertY = worldY - inst.MinY;

                            var insert = new Insert(inst.Block)
                            {
                                InsertPoint = new XYZ(insertX, insertY, 0.0),
                                XScale = 1.0,
                                YScale = 1.0,
                                ZScale = 1.0
                            };

                            modelSpace.Entities.Add(insert);

                            sheet.CurrentX += inst.Width + partGap;
                            if (inst.Height > sheet.RowHeight)
                                sheet.RowHeight = inst.Height;

                            placed++;
                            progress.Step($"Placed {placed} of {totalInstances} plates...");
                            break;
                        }
                        else
                        {
                            // Move to next row.
                            sheet.CurrentX = sheetMargin;
                            sheet.CurrentY += sheet.RowHeight + partGap;
                            sheet.RowHeight = 0.0;

                            // Not enough vertical space either? Start new sheet.
                            if (sheet.CurrentY + inst.Height > sheetHeight - sheetMargin)
                            {
                                sheet = NewSheet();
                            }
                        }
                    }
                }

                sheetCount = sheets.Count;
            }
            finally
            {
                progress.Close();
            }

            // Write nested result to a new DWG next to the source.
            string dir = Path.GetDirectoryName(sourceDwgPath);
            string nameNoExt = Path.GetFileNameWithoutExtension(sourceDwgPath);
            string outPath = Path.Combine(dir ?? string.Empty, nameNoExt + "_nested.dwg");

            using (var writer = new DwgWriter(outPath, doc))
            {
                writer.Write();
            }

            MessageBox.Show(
                "Laser cut nesting finished.\r\n\r\n" +
                "Sheets used: " + sheetCount + Environment.NewLine +
                "Total parts: " + totalInstances + Environment.NewLine +
                "Output DWG: " + outPath,
                "Laser Cut",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private static IEnumerable<PartDefinition> LoadPartDefinitions(CadDocument doc)
        {
            var list = new List<PartDefinition>();

            foreach (var br in doc.BlockRecords)
            {
                if (string.IsNullOrEmpty(br.Name))
                    continue;

                // Skip model/paper space and other special blocks.
                if (br.Name.StartsWith("*", StringComparison.Ordinal))
                    continue;

                // We only care about plate blocks created by CombineDwg, which start with "P_".
                if (!br.Name.StartsWith("P_", StringComparison.OrdinalIgnoreCase))
                    continue;

                int qty = ParseQuantityFromBlockName(br.Name);
                if (qty <= 0)
                    qty = 1;

                double minX = double.MaxValue;
                double minY = double.MaxValue;
                double maxX = double.MinValue;
                double maxY = double.MinValue;

                foreach (var ent in br.Entities)
                {
                    try
                    {
                        var bb = ent.GetBoundingBox();
                        var bmin = bb.Min;
                        var bmax = bb.Max;

                        if (bmin.X < minX) minX = bmin.X;
                        if (bmin.Y < minY) minY = bmin.Y;
                        if (bmax.X > maxX) maxX = bmax.X;
                        if (bmax.Y > maxY) maxY = bmax.Y;
                    }
                    catch
                    {
                        // ignore entities without bbox
                    }
                }

                if (minX == double.MaxValue || maxX == double.MinValue ||
                    minY == double.MaxValue || maxY == double.MinValue)
                {
                    // no measurable geometry
                    continue;
                }

                double width = maxX - minX;
                double height = maxY - minY;

                if (width <= 0.0 || height <= 0.0)
                    continue;

                list.Add(new PartDefinition
                {
                    Block = br,
                    BlockName = br.Name,
                    MinX = minX,
                    MinY = minY,
                    Width = width,
                    Height = height,
                    Quantity = qty
                });
            }

            return list;
        }

        private static int ParseQuantityFromBlockName(string blockName)
        {
            // Expected pattern: "..._Q<number>"
            if (string.IsNullOrEmpty(blockName))
                return 1;

            int idx = blockName.LastIndexOf("_Q", StringComparison.OrdinalIgnoreCase);
            if (idx < 0 || idx + 2 >= blockName.Length)
                return 1;

            string s = blockName.Substring(idx + 2);
            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int qty) && qty > 0)
                return qty;

            return 1;
        }
    }

    internal sealed class LaserCutOptionsForm : Form
    {
        private readonly TextBox _txtWidth;
        private readonly TextBox _txtHeight;
        private readonly Button _btnOk;
        private readonly Button _btnCancel;

        public double SheetWidthMm { get; private set; }
        public double SheetHeightMm { get; private set; }

        public LaserCutOptionsForm()
        {
            Text = "Laser cut sheet size";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            AutoSize = false;
            ClientSize = new System.Drawing.Size(320, 140);

            var lblWidth = new Label
            {
                AutoSize = true,
                Text = "Sheet width (mm):",
                Location = new System.Drawing.Point(12, 20)
            };
            Controls.Add(lblWidth);

            _txtWidth = new TextBox
            {
                Location = new System.Drawing.Point(150, 16),
                Width = 140,
                Text = "3000"
            };
            Controls.Add(_txtWidth);

            var lblHeight = new Label
            {
                AutoSize = true,
                Text = "Sheet height (mm):",
                Location = new System.Drawing.Point(12, 55)
            };
            Controls.Add(lblHeight);

            _txtHeight = new TextBox
            {
                Location = new System.Drawing.Point(150, 51),
                Width = 140,
                Text = "1500"
            };
            Controls.Add(_txtHeight);

            _btnOk = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.None,
                Location = new System.Drawing.Point(134, 95),
                Width = 75
            };
            _btnOk.Click += Ok_Click;
            Controls.Add(_btnOk);

            _btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(215, 95),
                Width = 75
            };
            Controls.Add(_btnCancel);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;
        }

        private void Ok_Click(object sender, EventArgs e)
        {
            if (!double.TryParse(_txtWidth.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double w) ||
                w <= 0)
            {
                MessageBox.Show(this, "Please enter a valid positive sheet width (mm).", "Laser Cut",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtWidth.Focus();
                _txtWidth.SelectAll();
                return;
            }

            if (!double.TryParse(_txtHeight.Text.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double h) ||
                h <= 0)
            {
                MessageBox.Show(this, "Please enter a valid positive sheet height (mm).", "Laser Cut",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtHeight.Focus();
                _txtHeight.SelectAll();
                return;
            }

            SheetWidthMm = w;
            SheetHeightMm = h;

            DialogResult = DialogResult.OK;
            Close();
        }
    }

    internal sealed class LaserCutProgressForm : Form
    {
        private readonly ProgressBar _progressBar;
        private readonly Label _label;
        private readonly int _maximum;
        private int _current;

        public LaserCutProgressForm(int maximum)
        {
            if (maximum <= 0)
                maximum = 1;

            _maximum = maximum;

            Text = "Working...";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            AutoSize = false;
            ClientSize = new System.Drawing.Size(400, 90);

            _label = new Label
            {
                AutoSize = false,
                Text = "Preparing...",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                Location = new System.Drawing.Point(12, 9),
                Size = new System.Drawing.Size(376, 20)
            };
            Controls.Add(_label);

            _progressBar = new ProgressBar
            {
                Location = new System.Drawing.Point(12, 35),
                Size = new System.Drawing.Size(376, 20),
                Minimum = 0,
                Maximum = _maximum,
                Value = 0
            };
            Controls.Add(_progressBar);
        }

        public void Step(string statusText)
        {
            if (!IsHandleCreated)
                return;

            if (!string.IsNullOrEmpty(statusText))
                _label.Text = statusText;

            if (_current < _maximum)
            {
                _current++;
                _progressBar.Value = _current;
            }

            _progressBar.Refresh();
            _label.Refresh();
            Application.DoEvents();
        }
    }
}
