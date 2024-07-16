using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

using System;
using System.Collections.Generic;

namespace J_Tools
{
    public class J_NonCoplanarTools : BaseCommand
    {
        // Detect non-coplanar blocks in the drawing
        [CommandMethod("DETECTNONCOPLANARBLOCKS")]
        public void DetectNonCoplanarBlocks()
        {
            // Define the filter for selecting blocks and nested blocks
            TypedValue[] filterList = new TypedValue[]
            {
                new TypedValue((int)DxfCode.Start, "INSERT"),
                new TypedValue((int)DxfCode.BlockName, "*"),
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, "ACAD"),
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, "AutoCAD")
            };
            SelectionFilter filter = new SelectionFilter(filterList);

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Get the block table and block table record
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForWrite);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Create a list to store the names of non-coplanar blocks
                List<string> nonCoplanarBlocks = new List<string>();
                ObjectIdCollection nonCoplanarBlockIds = new ObjectIdCollection();

                // Iterate through the entities in the model space
                foreach (ObjectId objId in btr)
                {
                    Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                    if (ent == null) continue;

                    // Select only blocks and nested blocks
                    if (ent is BlockReference)
                    {
                        BlockReference blockRef = (BlockReference)ent;

                        // Select only non-coplanar blocks
                        Plane blockPlane = new Plane(blockRef.Position, blockRef.Normal);
                        if (!blockPlane.Normal.IsParallelTo(Vector3d.ZAxis))
                        {
                            nonCoplanarBlocks.Add(blockRef.Name ?? "");
                            nonCoplanarBlockIds.Add(blockRef.ObjectId);
                        }

                    }
                    else if (ent is BlockBegin)
                    {
                        BlockReference blockRef = (BlockReference)ent;

                        BlockTableRecord nestedBtr = (BlockTableRecord)tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForWrite);

                        // Iterate through the entities in the nested block
                        foreach (ObjectId nestedObjId in nestedBtr)
                        {
                            var nestedEnt = tr.GetObject(nestedObjId, OpenMode.ForWrite) as Entity;
                            if (nestedEnt == null) continue;

                            // Select only non-coplanar blocks
                            if (nestedEnt is BlockReference)
                            {
                                BlockReference nestedBlockRef = (BlockReference)nestedEnt;

                                // Select only non-coplanar blocks
                                Plane blockPlane = new Plane(nestedBlockRef.Position, nestedBlockRef.Normal);
                                if (!blockPlane.Normal.IsParallelTo(Vector3d.ZAxis))
                                {
                                    nonCoplanarBlocks.Add(nestedBlockRef.Name ?? "");
                                    nonCoplanarBlockIds.Add(nestedBlockRef.ObjectId);
                                }


                            }
                        }
                    }
                }

                // Print the names of non-coplanar blocks in the command window
                foreach (string blockName in nonCoplanarBlocks)
                {
                    ed.WriteMessage("\nNon-coplanar block found: " + blockName);
                }

                PromptResult result = ed.GetKeywords("\nDo you want to fix/flatten those blocks? [Yes/No]", "Yes", "No");
                if (result.Status == PromptStatus.OK)
                {
                    if (result.StringResult == "Yes")
                    {
                        // Do re-insert the block if user selects "Yes"
                        foreach (ObjectId nonCoplanarBlockId in nonCoplanarBlockIds)
                        {
                            BlockReference blockReference = (BlockReference)tr.GetObject(nonCoplanarBlockId, OpenMode.ForWrite);

                            BlockTableRecord blockDefinition = (BlockTableRecord)blockReference.BlockTableRecord.GetObject(OpenMode.ForWrite);

                            Point3d position = blockReference.Position;
                            Scale3d scaleFactors = blockReference.ScaleFactors;

                            blockReference.Erase();

                            BlockReference newBlockReference = new BlockReference(position, blockDefinition.ObjectId)
                            {
                                ScaleFactors = scaleFactors
                            };

                            blockReference.RecordGraphicsModified(true);

                            blockReference = newBlockReference;
                            blockReference.Position = position;

                            btr.AppendEntity(blockReference);
                            tr.AddNewlyCreatedDBObject(blockReference, true);
                        }

                    }
                    else if (result.StringResult == "No")
                    {
                        // Do terminate the program if user selects "No"
                        ed.WriteMessage("\nNon-coplanar block replacement terminated.");
                        return;
                    }
                }

                tr.Commit();
            }
        }

        /////////////////////////////////////////////////////////

        // Detect non-coplanar arcs and circles in the drawing
        [CommandMethod("DETECTNONCOPLANARARCSANDCIRCLES")]
        public void DetectNonCoplanarArcsAndCircles()
        {
            // Start a transaction
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open the ModelSpace block table record for read
                BlockTableRecord btr = tr.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForRead) as BlockTableRecord;

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

                // Print the number of non-coplanar entities found
                int numNonCoplanar = nonCoplanarIds.Count;
                ed.WriteMessage("\nFound {0} non-coplanar arcs and circles.", numNonCoplanar);

                // Ask the user if they want to flatten the non-coplanar curves
                PromptKeywordOptions pko = new PromptKeywordOptions("\nDo you want to fix/flatten those arcs and circles? [Yes/No]");
                pko.Keywords.Add("Yes");
                pko.Keywords.Add("No");
                pko.AllowNone = true;
                PromptResult result = ed.GetKeywords(pko);

                if (result.Status == PromptStatus.OK && result.StringResult == "Yes")
                {
                    // Flatten the non-coplanar arcs and circles to the XY plane
                    FlattenEntities(nonCoplanarIds, tr);
                }
                else
                {
                    // Terminate
                    ed.WriteMessage("\nNon-coplanar arc and circle flattening terminated.");
                }

                // Commit the transaction
                tr.Commit();
            }
        }

        // Helper function  - DETECTNONCOPLANARARCSANDCIRCLES
        // to check whether the curve is non-coplanar
        private bool IsEntityCoplanar(Entity ent)
        {
            if (ent is Circle circle)
            {
                //Check if the circle is coplanar with the XY plane
                return Math.Abs(circle.Normal.Z) < Tolerance.Global.EqualPoint;
            }
            else if (ent is Arc arc)
            {
                // Check if the arc is coplanar with the XY plane
                return Math.Abs(arc.Normal.Z) < Tolerance.Global.EqualPoint;
            }

            return false;
        }

        // Helper function - DETECTNONCOPLANARARCSANDCIRCLES
        // to flatten the non-coplanar entities to the XY plane
        private void FlattenEntities(List<ObjectId> nonCoplanarIds, Transaction tr)
        {
            foreach (ObjectId id in nonCoplanarIds)
            {
                Entity ent = tr.GetObject(id, OpenMode.ForWrite) as Entity;

                if (ent is Circle circle)
                {
                    // Flatten the circle to the XY plane
                    //circle.TransformBy(Matrix3d.WorldToPlane(new Plane(Point3d.Origin, Vector3d.ZAxis)));
                    circle.Normal = Vector3d.ZAxis;
                }
                else if (ent is Arc arc)
                {
                    // Flatten the arc to the XY plane
                    // arc.TransformBy(Matrix3d.WorldToPlane(new Plane(Point3d.Origin, Vector3d.ZAxis)));
                    arc.Normal = Vector3d.ZAxis;
                }
            }
        }

        /////////////////////////////////////////////////////////
    }
}

