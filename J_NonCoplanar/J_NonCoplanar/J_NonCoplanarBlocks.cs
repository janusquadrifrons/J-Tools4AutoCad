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
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // Create a list to store the names of non-coplanar blocks
                List<string> nonCoplanarBlocks = new List<string>();

                // Iterate through the entities in the model space
                foreach (ObjectId objId in btr)
                {
                    var ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
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
                        }

                    }
                    else if (ent is BlockBegin)
                    {
                        BlockReference blockRef = (BlockReference)ent;

                        BlockTableRecord nestedBtr = (BlockTableRecord)tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead);

                        // Iterate through the entities in the nested block
                        foreach (ObjectId nestedObjId in nestedBtr)
                        {
                            var nestedEnt = tr.GetObject(nestedObjId, OpenMode.ForRead) as Entity;
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

                PromptResult result = ed.GetKeywords("\nDo you want to fix/flatten those blocks? [Yes/No]", "Yes No");
                if (result.Status == PromptStatus.OK)
                {
                    if (result.StringResult == "Yes")
                    {
                        // Do re-insert the block if user selects "Yes"
                        foreach (var nonCoplanarBlock in nonCoplanarBlocks)
                        {
                            var objectId = new ObjectId(nonCoplanarBlock.ObjectId);
                            var blockReference = (BlockReference)objectId.GetObject(OpenMode.ForWrite);

                            var blockDefinition = (BlockTableRecord)blockReference.BlockTableRecord.GetObject(OpenMode.ForRead);

                            var position = blockReference.Position;
                            var scaleFactors = blockReference.ScaleFactors;

                            blockReference.Erase();

                            var newBlockReference = new BlockReference(position, blockDefinition.ObjectId)
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
                    else
                    {
                        // Do something if user selects "No"
                    }
                }

                tr.Commit();
            }
        }
    }
}
