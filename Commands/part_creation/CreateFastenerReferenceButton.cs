using System;
using System.Diagnostics;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SW2026RibbonAddin.Forms;

namespace SW2026RibbonAddin.Commands
{
    internal sealed class CreateFastenerReferencesButton : IMehdiRibbonButton
    {
        public string Id => "CreateFastenerReferences";

        public string DisplayName => "Create fastener\nreferences...";
        public string Tooltip =>
            "Create Mate_Axis / Mate_Underside / Mate_Bottom / Mate_Top reference geometry on the active fastener part.";
        public string Hint => "Create standardized reference geometry for bolts, washers, and nuts";

        public string SmallIconFile => "reference_geometry_20.png";
        public string LargeIconFile => "reference_geometry_32.png";

        public RibbonSection Section => RibbonSection.PartCreation;
        public int SectionOrder => 2;

        public bool IsFreeFeature => true;

        public int GetEnableState(AddinContext context)
        {
            try
            {
                var model = context.ActiveModel;
                if (model == null)
                    return AddinContext.Disable;

                if (model.GetType() != (int)swDocumentTypes_e.swDocPART)
                    return AddinContext.Disable;

                // Only allow on standalone (non‑Toolbox) parts
                var ext = model.Extension as ModelDocExtension;
                if (ext != null &&
                    ext.ToolboxPartType != (int)swToolBoxPartType_e.swNotAToolboxPart)
                {
                    return AddinContext.Disable;
                }

                return AddinContext.Enable;
            }
            catch
            {
                return AddinContext.Disable;
            }
        }

        public void Execute(AddinContext context)
        {
            try
            {
                var model = context.ActiveModel;
                if (model == null || model.GetType() != (int)swDocumentTypes_e.swDocPART)
                {
                    MessageBox.Show("Open a fastener part (*.sldprt) before running this command.",
                        "Create fastener references");
                    return;
                }

                var part = model as IPartDoc;
                if (part == null)
                {
                    MessageBox.Show("The active document is not a part.", "Create fastener references");
                    return;
                }

                // Ensure this is not a Toolbox library part (we only want standalone clones)
                var ext = model.Extension as ModelDocExtension;
                if (ext != null &&
                    ext.ToolboxPartType != (int)swToolBoxPartType_e.swNotAToolboxPart)
                {
                    MessageBox.Show(
                        "This part is still controlled by Toolbox.\r\n\r\n" +
                        "Use the 'Clone Part' command first to create a standalone fastener,\r\n" +
                        "then run 'Create fastener references...'.",
                        "Create fastener references",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                var values = FastenerPropertyHelper.BuildInitialValues(model);
                string family = values.Family;

                if (!IsSupportedFamily(family))
                {
                    using (var typeForm = new FastenerReferenceTypeForm(family))
                    {
                        if (typeForm.ShowDialog() != DialogResult.OK)
                            return;

                        family = typeForm.SelectedFamily;
                    }

                    // Persist the chosen family so we do not have to ask again
                    values.Family = family;
                    FastenerPropertyHelper.WriteProperties(model, values);
                }

                var selMgr = model.SelectionManager as ISelectionMgr;
                if (selMgr == null)
                {
                    MessageBox.Show("Selection manager is not available.", "Create fastener references");
                    return;
                }

                var selCount = selMgr.GetSelectedObjectCount2(-1);
                if (selCount == 0)
                {
                    MessageBox.Show(
                        "Please preselect the required faces/planes before running this command:\n\n" +
                        "Bolt / Screw:\n" +
                        "  1) Cylindrical face of the shank.\n" +
                        "  2) Planar face on the underside of the head (or an existing plane).\n\n" +
                        "Washer / Nut:\n" +
                        "  1) Cylindrical face of the hole.\n" +
                        "  2) Planar face that should be Mate_Bottom (or an existing plane).\n" +
                        "  3) Planar face that should be Mate_Top (or an existing plane).\n\n" +
                        "Then run this command again.",
                        "Create fastener references");
                    return;
                }

                Feature axisFeat = null;
                Feature undersidePlane = null;
                Feature bottomPlane = null;
                Feature topPlane = null;

                var fam = (family ?? string.Empty).Trim();

                if (fam.Equals("Bolt", StringComparison.OrdinalIgnoreCase))
                {
                    CreateBoltReferences(model, part, selMgr, out axisFeat, out undersidePlane);
                }
                else if (fam.Equals("Washer", StringComparison.OrdinalIgnoreCase))
                {
                    CreateWasherReferences(model, part, selMgr, out axisFeat, out bottomPlane, out topPlane);
                }
                else if (fam.Equals("Nut", StringComparison.OrdinalIgnoreCase))
                {
                    CreateNutReferences(model, part, selMgr, out axisFeat, out bottomPlane, out topPlane);
                }
                else
                {
                    MessageBox.Show($"Fastener type '{family}' is not supported. Use Bolt, Washer or Nut.",
                        "Create fastener references");
                    return;
                }

                // Highlight what we created/updated
                model.ClearSelection2(true);
                axisFeat?.Select2(false, -1);
                undersidePlane?.Select2(true, -1);
                bottomPlane?.Select2(true, -1);
                topPlane?.Select2(true, -1);

                model.GraphicsRedraw2();

                MessageBox.Show(
                    "Reference geometry created/updated.\n" +
                    "Verify Mate_Axis and Mate_* planes visually before using automated insertion tools.",
                    "Create fastener references");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error while creating fastener references:\r\n\r\n" + ex.Message,
                    "Create fastener references");
                Debug.WriteLine("CreateFastenerReferencesButton.Execute error: " + ex);
            }
        }

        private static bool IsSupportedFamily(string family)
        {
            if (string.IsNullOrWhiteSpace(family))
                return false;

            var f = family.Trim();
            return f.Equals("Bolt", StringComparison.OrdinalIgnoreCase) ||
                   f.Equals("Washer", StringComparison.OrdinalIgnoreCase) ||
                   f.Equals("Nut", StringComparison.OrdinalIgnoreCase);
        }

        // ---------------- BOLT: Mate_Axis + Mate_Underside ----------------

        private static void CreateBoltReferences(
            IModelDoc2 model,
            IPartDoc part,
            ISelectionMgr selMgr,
            out Feature axisFeat,
            out Feature undersidePlane)
        {
            axisFeat = null;
            undersidePlane = null;

            Face2 cylFace = null;
            Face2 undersideFace = null;
            Feature undersidePlaneFeature = null;

            ClassifySelectionsForBolt(selMgr, ref cylFace, ref undersideFace, ref undersidePlaneFeature);

            if (cylFace == null)
            {
                throw new InvalidOperationException(
                    "No cylindrical face was found in the selection.\n" +
                    "Select the bolt shank cylindrical face and the underside planar face, then run the command again.");
            }

            axisFeat = EnsureAxisOnFace(model, part, cylFace, "Mate_Axis");

            if (undersidePlaneFeature != null)
            {
                undersidePlane = EnsurePlaneFromExisting(model, part, undersidePlaneFeature, "Mate_Underside");
            }
            else if (undersideFace != null)
            {
                undersidePlane = EnsurePlaneOnFace(model, part, undersideFace, "Mate_Underside");
            }
            else
            {
                MessageBox.Show(
                    "Mate_Axis was created/updated, but no planar face or plane was selected for Mate_Underside.",
                    "Create fastener references");
            }
        }

        private static void ClassifySelectionsForBolt(
            ISelectionMgr selMgr,
            ref Face2 cylFace,
            ref Face2 undersideFace,
            ref Feature undersidePlaneFeat)
        {
            if (selMgr == null) return;

            int count = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= count; i++)
            {
                int type = selMgr.GetSelectedObjectType3(i, -1);
                object obj = selMgr.GetSelectedObject6(i, -1);

                if (type == (int)swSelectType_e.swSelFACES && obj is Face2 face)
                {
                    var surf = face.GetSurface() as Surface;
                    if (surf != null && surf.IsCylinder())
                    {
                        if (cylFace == null)
                            cylFace = face;
                    }
                    else if (surf != null && surf.IsPlane())
                    {
                        if (undersideFace == null)
                            undersideFace = face;
                    }
                }
                else if (type == (int)swSelectType_e.swSelDATUMPLANES && obj is Feature feat)
                {
                    if (undersidePlaneFeat == null)
                        undersidePlaneFeat = feat;
                }
            }
        }

        // ----------- WASHER / NUT: Mate_Axis + Mate_Bottom + Mate_Top -----------

        private static void CreateWasherReferences(
            IModelDoc2 model,
            IPartDoc part,
            ISelectionMgr selMgr,
            out Feature axisFeat,
            out Feature bottomPlane,
            out Feature topPlane)
        {
            axisFeat = null;
            bottomPlane = null;
            topPlane = null;

            Face2 cylFace = null;
            var planarFaces = new System.Collections.Generic.List<Face2>();
            var planeFeatures = new System.Collections.Generic.List<Feature>();

            ClassifySelectionsForWasherOrNut(selMgr, ref cylFace, planarFaces, planeFeatures);

            if (cylFace == null)
            {
                throw new InvalidOperationException(
                    "No cylindrical face was found in the selection.\n" +
                    "Select the washer hole cylindrical face and the two planar faces, then run the command again.");
            }

            axisFeat = EnsureAxisOnFace(model, part, cylFace, "Mate_Axis");

            if (planeFeatures.Count >= 1)
                bottomPlane = EnsurePlaneFromExisting(model, part, planeFeatures[0], "Mate_Bottom");
            if (planeFeatures.Count >= 2)
                topPlane = EnsurePlaneFromExisting(model, part, planeFeatures[1], "Mate_Top");

            if (bottomPlane == null && planarFaces.Count >= 1)
                bottomPlane = EnsurePlaneOnFace(model, part, planarFaces[0], "Mate_Bottom");

            if (topPlane == null)
            {
                if (planarFaces.Count >= 2)
                    topPlane = EnsurePlaneOnFace(model, part, planarFaces[1], "Mate_Top");
                else if (planarFaces.Count == 1 && bottomPlane == null)
                    topPlane = EnsurePlaneOnFace(model, part, planarFaces[0], "Mate_Top");
            }
        }

        private static void CreateNutReferences(
            IModelDoc2 model,
            IPartDoc part,
            ISelectionMgr selMgr,
            out Feature axisFeat,
            out Feature bottomPlane,
            out Feature topPlane)
        {
            axisFeat = null;
            bottomPlane = null;
            topPlane = null;

            Face2 cylFace = null;
            var planarFaces = new System.Collections.Generic.List<Face2>();
            var planeFeatures = new System.Collections.Generic.List<Feature>();

            ClassifySelectionsForWasherOrNut(selMgr, ref cylFace, planarFaces, planeFeatures);

            if (cylFace == null)
            {
                throw new InvalidOperationException(
                    "No cylindrical face was found in the selection.\n" +
                    "Select the nut threaded hole cylindrical face and the two planar faces, then run the command again.");
            }

            axisFeat = EnsureAxisOnFace(model, part, cylFace, "Mate_Axis");

            if (planeFeatures.Count >= 1)
                bottomPlane = EnsurePlaneFromExisting(model, part, planeFeatures[0], "Mate_Bottom");
            if (planeFeatures.Count >= 2)
                topPlane = EnsurePlaneFromExisting(model, part, planeFeatures[1], "Mate_Top");

            if (bottomPlane == null && planarFaces.Count >= 1)
                bottomPlane = EnsurePlaneOnFace(model, part, planarFaces[0], "Mate_Bottom");

            if (topPlane == null)
            {
                if (planarFaces.Count >= 2)
                    topPlane = EnsurePlaneOnFace(model, part, planarFaces[1], "Mate_Top");
                else if (planarFaces.Count == 1 && bottomPlane == null)
                    topPlane = EnsurePlaneOnFace(model, part, planarFaces[0], "Mate_Top");
            }
        }

        private static void ClassifySelectionsForWasherOrNut(
            ISelectionMgr selMgr,
            ref Face2 cylFace,
            System.Collections.Generic.List<Face2> planarFaces,
            System.Collections.Generic.List<Feature> planeFeatures)
        {
            if (selMgr == null) return;

            int count = selMgr.GetSelectedObjectCount2(-1);
            for (int i = 1; i <= count; i++)
            {
                int type = selMgr.GetSelectedObjectType3(i, -1);
                object obj = selMgr.GetSelectedObject6(i, -1);

                if (type == (int)swSelectType_e.swSelFACES && obj is Face2 face)
                {
                    var surf = face.GetSurface() as Surface;
                    if (surf != null && surf.IsCylinder())
                    {
                        if (cylFace == null)
                            cylFace = face;
                    }
                    else if (surf != null && surf.IsPlane())
                    {
                        planarFaces.Add(face);
                    }
                }
                else if (type == (int)swSelectType_e.swSelDATUMPLANES && obj is Feature feat)
                {
                    planeFeatures.Add(feat);
                }
            }
        }

        // ---------------- AXIS / PLANE CREATION HELPERS ----------------

        private static Feature EnsureAxisOnFace(IModelDoc2 model, IPartDoc part, Face2 face, string featureName)
        {
            if (model == null || part == null || face == null)
                return null;

            var existing = part.FeatureByName(featureName) as Feature;
            if (existing != null)
            {
                var overwrite = MessageBox.Show(
                    $"A feature named '{featureName}' already exists.\nDo you want to recreate it?",
                    "Create fastener references",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (overwrite == DialogResult.Yes)
                {
                    existing.Select2(false, -1);
                    model.Extension.DeleteSelection2(0);
                }
                else
                {
                    return existing;
                }
            }

            model.ClearSelection2(true);

            var entity = (IEntity)face;
            entity.Select4(false, null);

            bool ok = model.InsertAxis2(true);
            if (!ok)
                throw new InvalidOperationException(
                    $"Failed to create {featureName} axis from the selected cylindrical face.");

            Feature lastAxis = null;
            Feature featIt = part.FirstFeature() as Feature;
            while (featIt != null)
            {
                if (string.Equals(featIt.GetTypeName2(), "RefAxis", StringComparison.OrdinalIgnoreCase))
                    lastAxis = featIt;

                featIt = featIt.GetNextFeature() as Feature;
            }

            if (lastAxis == null)
                throw new InvalidOperationException($"{featureName} axis feature was not found after creation.");

            lastAxis.Name = featureName;
            return lastAxis;
        }

        private static Feature EnsurePlaneOnFace(IModelDoc2 model, IPartDoc part, Face2 face, string featureName)
        {
            if (model == null || part == null || face == null)
                return null;

            var existing = part.FeatureByName(featureName) as Feature;
            if (existing != null)
            {
                var overwrite = MessageBox.Show(
                    $"A feature named '{featureName}' already exists.\nDo you want to recreate it?",
                    "Create fastener references",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (overwrite == DialogResult.Yes)
                {
                    existing.Select2(false, -1);
                    model.Extension.DeleteSelection2(0);
                }
                else
                {
                    return existing;
                }
            }

            model.ClearSelection2(true);

            var entity = (IEntity)face;
            entity.Select4(false, null);

            var featMgr = model.FeatureManager as IFeatureManager;
            if (featMgr == null)
                throw new InvalidOperationException("FeatureManager is not available.");

            featMgr.InsertRefPlane(
                (int)swRefPlaneReferenceConstraints_e.swRefPlaneReferenceConstraint_Coincident,
                0.0, 0,
                0.0, 0,
                0.0);

            // Find last RefPlane feature
            Feature lastPlane = null;
            Feature featIt = part.FirstFeature() as Feature;
            while (featIt != null)
            {
                if (string.Equals(featIt.GetTypeName2(), "RefPlane", StringComparison.OrdinalIgnoreCase))
                    lastPlane = featIt;

                featIt = featIt.GetNextFeature() as Feature;
            }

            if (lastPlane == null)
                throw new InvalidOperationException($"{featureName} plane feature was not found after creation.");

            lastPlane.Name = featureName;
            return lastPlane;
        }

        private static Feature EnsurePlaneFromExisting(IModelDoc2 model, IPartDoc part, Feature planeFeature, string featureName)
        {
            if (planeFeature == null)
                return null;

            var existing = part.FeatureByName(featureName) as Feature;
            if (existing != null && !ReferenceEquals(existing, planeFeature))
            {
                var overwrite = MessageBox.Show(
                    $"A feature named '{featureName}' already exists.\nDo you want to rename the selected plane and replace it?",
                    "Create fastener references",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (overwrite == DialogResult.Yes)
                {
                    existing.Select2(false, -1);
                    model.Extension.DeleteSelection2(0);
                }
                else
                {
                    return existing;
                }
            }

            planeFeature.Name = featureName;
            return planeFeature;
        }
    }
}
