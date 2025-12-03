using ACadSharp;
using ACadSharp.Entities;
using ACadSharp.IO;
using ACadSharp.Tables;
using ACadSharp.Types;
using CSMath;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace SW2026RibbonAddin
{
    internal static class DwgPostProcessor
    {
        private static readonly Random _rng = new Random();

        /// <summary>
        /// Post-processes a DWG created by SolidWorks:
        /// 1) Moves all model-space entities into a new block.
        /// 2) Inserts one instance of that block at the origin with a random color.
        /// 3) Adds two text lines under the geometry:
        ///      PLATE: &lt;thicknessMm&gt;
        ///      QUANTITY: &lt;quantity&gt;
        /// </summary>
        public static void ConvertToBlockAndAnnotate(
            string dwgPath,
            double thicknessMm,
            int quantity)
        {
            if (string.IsNullOrWhiteSpace(dwgPath) || !File.Exists(dwgPath))
                return;

            CadDocument doc;
            try
            {
                doc = DwgReader.Read(dwgPath, OnNotification);
            }
            catch
            {
                return;
            }

            if (doc == null)
                return;

            // Get model space
            BlockRecord modelSpace;
            try
            {
                modelSpace = doc.BlockRecords["*Model_Space"];
            }
            catch
            {
                return;
            }

            // Copy entities out of model space
            List<Entity> originalEntities = modelSpace.Entities.ToList();
            if (originalEntities.Count == 0)
                return;

            // Bounding box of the original geometry
            ComputeBounds(originalEntities,
                out double minX,
                out double maxX,
                out double minY,
                out double maxY);

            // Create block and move entities into it
            string baseName = Path.GetFileNameWithoutExtension(dwgPath) ?? "PART";
            string blockName = $"FLAT_{baseName}";

            var block = new BlockRecord(blockName);
            doc.BlockRecords.Add(block);

            foreach (Entity e in originalEntities)
            {
                modelSpace.Entities.Remove(e);
                block.Entities.Add(e);
            }

            // Insert the block at origin
            var insert = new Insert(block)
            {
                InsertPoint = new XYZ(0, 0, 0)
            };
            ApplyRandomColor(insert);
            modelSpace.Entities.Add(insert);

            // Add text under the part, with some extra bottom margin
            AddPlateAndQuantityText(
                modelSpace,
                minX, maxX, minY, maxY,
                thicknessMm,
                quantity);

            // Save back
            using (var writer = new DwgWriter(dwgPath, doc))
            {
                writer.OnNotification += OnNotification;
                writer.Write();
            }
        }

        #region Bounding box + text placement

        /// <summary>
        /// Uses reflection on Entity.GetBoundingBox() so that ALL entity types
        /// (including polylines) contribute to the extents.
        /// </summary>
        private static void ComputeBounds(
            IEnumerable<Entity> entities,
            out double minX, out double maxX,
            out double minY, out double maxY)
        {
            bool initialized = false;
            double localMinX = 0, localMaxX = 0, localMinY = 0, localMaxY = 0;

            void Update(double x, double y)
            {
                if (!initialized)
                {
                    initialized = true;
                    localMinX = localMaxX = x;
                    localMinY = localMaxY = y;
                    return;
                }

                if (x < localMinX) localMinX = x;
                if (x > localMaxX) localMaxX = x;
                if (y < localMinY) localMinY = y;
                if (y > localMaxY) localMaxY = y;
            }

            void UpdateFromPoint(object point)
            {
                if (point == null) return;

                try
                {
                    var ptType = point.GetType();
                    var xProp = ptType.GetProperty("X");
                    var yProp = ptType.GetProperty("Y");
                    if (xProp == null || yProp == null) return;

                    object xObj = xProp.GetValue(point, null);
                    object yObj = yProp.GetValue(point, null);
                    if (xObj == null || yObj == null) return;

                    double x = Convert.ToDouble(xObj);
                    double y = Convert.ToDouble(yObj);
                    Update(x, y);
                }
                catch
                {
                    // ignore this point
                }
            }

            foreach (Entity ent in entities)
            {
                try
                {
                    var method = ent.GetType().GetMethod("GetBoundingBox", Type.EmptyTypes);
                    if (method == null)
                        continue;

                    var box = method.Invoke(ent, null);
                    if (box == null)
                        continue;

                    var boxType = box.GetType();
                    var minProp =
                        boxType.GetProperty("Min") ??
                        boxType.GetProperty("Minimum") ??
                        boxType.GetProperty("MinPoint");
                    var maxProp =
                        boxType.GetProperty("Max") ??
                        boxType.GetProperty("Maximum") ??
                        boxType.GetProperty("MaxPoint");

                    if (minProp == null || maxProp == null)
                        continue;

                    var minVal = minProp.GetValue(box, null);
                    var maxVal = maxProp.GetValue(box, null);

                    UpdateFromPoint(minVal);
                    UpdateFromPoint(maxVal);
                }
                catch
                {
                    // ignore this entity
                }
            }

            if (!initialized)
            {
                minX = maxX = minY = maxY = 0.0;
            }
            else
            {
                minX = localMinX;
                maxX = localMaxX;
                minY = localMinY;
                maxY = localMaxY;
            }
        }

        private static void AddPlateAndQuantityText(
            BlockRecord modelSpace,
            double minX, double maxX,
            double minY, double maxY,
            double thicknessMm,
            int quantity)
        {
            double width = maxX - minX;
            double height = maxY - minY;

            if (width <= 0 || height <= 0)
            {
                // Fallback: use some default extents
                width = height = 100.0;
                minX = 0;
                minY = 0;
            }

            // Reasonable default text size
            double textHeight = Math.Max(height * 0.05, 5.0);
            double lineGap = textHeight * 1.25;

            // Add extra margin under the part:
            // geometry remains where it is, we place text below everything.
            double bottomMargin = Math.Max(height * 0.15, 3.0 * textHeight);

            double xLeft = minX; // left aligned with part
            double yPlate = minY - bottomMargin;
            double yQty = yPlate - lineGap;

            string plateText = $"PLATE: {thicknessMm:0.###}";
            string qtyText = $"QUANTITY: {quantity}";

            var plateNote = new TextEntity
            {
                Value = plateText,
                InsertPoint = new XYZ(xLeft, yPlate, 0),
                Height = textHeight
            };

            var qtyNote = new TextEntity
            {
                Value = qtyText,
                InsertPoint = new XYZ(xLeft, yQty, 0),
                Height = textHeight
            };

            modelSpace.Entities.Add(plateNote);
            modelSpace.Entities.Add(qtyNote);
        }

        #endregion

        #region Random color

        /// <summary>
        /// Tries to set entity.Color.Index to a random ACI index (1..255) using reflection.
        /// Works even if we don't know the concrete color type at compile time.
        /// </summary>
        private static void ApplyRandomColor(Entity entity)
        {
            if (entity == null)
                return;

            try
            {
                var colorProp = entity.GetType().GetProperty(
                    "Color",
                    BindingFlags.Instance | BindingFlags.Public);

                if (colorProp == null || !colorProp.CanWrite)
                    return;

                object colorValue = colorProp.GetValue(entity, null);
                if (colorValue == null)
                    return;

                Type colorType = colorValue.GetType();
                var indexProp = colorType.GetProperty(
                    "Index",
                    BindingFlags.Instance | BindingFlags.Public);

                if (indexProp == null || !indexProp.CanWrite)
                    return;

                short index;
                lock (_rng)
                {
                    index = (short)_rng.Next(1, 255); // 0 is BYBLOCK
                }

                indexProp.SetValue(colorValue, index, null);
                colorProp.SetValue(entity, colorValue, null);
            }
            catch
            {
                // purely cosmetic; ignore any problems
            }
        }

        #endregion

        #region Notifications

        private static void OnNotification(object sender, NotificationEventArgs e)
        {
            // Optional: log DWG read/write notifications if you want to debug
            // System.Diagnostics.Debug.WriteLine("ACadSharp: " + e.Message);
        }

        #endregion
    }
}
