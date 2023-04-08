using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System.Collections.Generic;

namespace MyAutoCADPlugin
{
    public class NonCoplanarBlockDetector
    {
        [CommandMethod("DETECTNONCOPLANARBLOCKS")]
        public void DetectNonCoplanarBlocks()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

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
    }
}
