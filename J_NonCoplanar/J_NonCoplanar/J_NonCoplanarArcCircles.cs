using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;

namespace MyAutoCADPlugin
{
    public class NonCoplanarArcsAndCirclesDetector
    {
        [CommandMethod("DETECTNONCOPLANARARCSANDCIRCLES")]
        public void DetectNonCoplanar()
        {
            // Get the current document and database
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;


            // Start a transaction
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open the ModelSpace block table record for read
                BlockTableRecord ms = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead) as BlockTableRecord;

                // Create a filter for circles and arcs
                TypedValue[] filterList = new TypedValue[]
                {
                    new TypedValue((int)DxfCode.Operator, "<OR"),
                    new TypedValue((int)DxfCode.Start, "CIRCLE"),
                    new TypedValue((int)DxfCode.Start, "ARC"),
                    new TypedValue((int)DxfCode.Operator, "OR>")
                };
                SelectionFilter filter = new SelectionFilter(filterList);

                // Select all circles and arcs in ModelSpace
                PromptSelectionResult selResult = doc.Editor.SelectAll(filter);
                if (selResult.Status != PromptStatus.OK)
                {
                    return;
                }

                // Get the selection set and iterate over the objects
                SelectionSet selSet = selResult.Value;
                List<ObjectId> nonCoplanarIds = new List<ObjectId>();

                foreach (SelectedObject selObj in selSet)
                {
                    // Open the object for read and check if it is a circle or an arc
                    Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                    if (ent is Circle || ent is Arc)
                    {
                        // Check if the entity is non-coplanar to the XY plane
                        if (!IsEntityCoplanar(ent))
                        {
                            // Accumulate the ObjectId in the collection
                            nonCoplanarIds.Add(ent.ObjectId);
                        }
                    }
                }

                // TODO: Do something with the nonCoplanarIds collection
                int numNonCoplanar = nonCoplanarIds.Count;

                ed.WriteMessage("\nFound {0} non-coplanar arcs and circles.", numNonCoplanar);

                // Ask the user if they want to flatten the non-coplanar curves
                PromptResult result = ed.GetKeywords("\nDo you want to fix/flatten those blocks? [Yes/No]", "Yes", "No");
                if (result.Status == PromptStatus.OK)
                {
                    if (result.StringResult == "Yes")
                    {
                        // TODO : Do project non-coplanar arcs and circles to the xy plane
                    }
                    else if (result.StringResult == "No")
                    {
                        // Do terminate the program if user selects "No"
                        ed.WriteMessage("\nNon-coplanar block replacement terminated.");
                        return;
                    }
                }

                // Commit the transaction
                tr.Commit();
            }
        }

        // Helper function to check whether the curve is non-coplanar
        private bool IsEntityCoplanar(Entity ent)
        {
            if (ent is Circle || ent is Arc)
            {
                var plane = new Plane(new Point3d(0, 0, 0), new Vector3d(0, 0, 1));
                var normal = ent.GetPlane().Normal;
                return normal.IsParallelTo(plane.Normal);
            }

            return true;
        }


        // Helper function to find the outer edge of all the objects in the nonCoplanarIds collection
        private Point3d[] FindBoundingBox(Database db, ObjectIdCollection objectIds)
        {
            // Create an object id array from the object id collection
            ObjectId[] ids = new ObjectId[objectIds.Count];
            objectIds.CopyTo(ids, 0);

            // Get the extents of the objects in the collection
            Extents3d ext = ids[0].GetObject(OpenMode.ForRead).GeometricExtents;
            for (int i = 1; i < ids.Length; i++)
            {
                Entity ent = ids[i].GetObject(OpenMode.ForRead) as Entity;
                if (ent == null) continue;
                ext.AddExtents(ent.GeometricExtents);
            }

            // Check if the extents are valid
            if (ext.IsNull)
            {
                return null;
            }

            // Return the bounding box vertices
            return new Point3d[]
            {
        new Point3d(ext.MinPoint.X, ext.MinPoint.Y, 0),
        new Point3d(ext.MinPoint.X, ext.MaxPoint.Y, 0),
        new Point3d(ext.MaxPoint.X, ext.MaxPoint.Y, 0),
        new Point3d(ext.MaxPoint.X, ext.MinPoint.Y, 0)
            };
        }





    }
}
