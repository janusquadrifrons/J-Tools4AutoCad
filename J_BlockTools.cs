using System.IO; // --- File Management
using System.Text.RegularExpressions; // --- Regex 

/// Block related operators

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.Geometry;

namespace J_Tools
{
    public class J_BlockTools : BaseCommand
    {
        //  Get entity's layer properties, and assign to entity itself in order to shorten Layer Table.

        [CommandMethod("BLOCKSIMPLFY")]
        public void BlockSimplfy()
        {

            PromptEntityOptions peopt = new PromptEntityOptions("\nSelect block reference");
            peopt.SetRejectMessage("\nSelect only block reference");
            peopt.AddAllowedClass(typeof(BlockReference), false);
            peopt.AllowNone = false;

            PromptEntityResult peres = ed.GetEntity(peopt);
            if (peres != null)
            {
                try
                {
                    using (Transaction tx = db.TransactionManager.StartTransaction())
                    {
                        BlockReference bref = tx.GetObject(peres.ObjectId, OpenMode.ForRead) as BlockReference;
                        BlockTableRecord btrec = null;

                        LayerTable laytab = tx.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                        LayerTableRecord laytabrec = null;

                        if (bref.IsDynamicBlock) { btrec = tx.GetObject(bref.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord; }
                        else { btrec = tx.GetObject(bref.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord; }

                        if (btrec != null) { ed.WriteMessage("Block name is : " + btrec.Name + "\n"); }

                        foreach (ObjectId oid in btrec)
                        {
                            Entity en = tx.GetObject(oid, OpenMode.ForWrite) as Entity;
                            //DBObject dbObj = tx.GetObject(oid, OpenMode.ForRead) as DBObject;

                            ed.WriteMessage("\nType : " + en.GetType().Name);

                            if (en is Dimension) { ed.WriteMessage("\nDimension detected..." + en); continue; }
                            else
                            {
                                // Get entity properties
                                int en_clr = en.ColorIndex;
                                string en_lt = en.Linetype;
                                LineWeight en_lw = en.LineWeight;
                                ed.WriteMessage("\nEntity color/linetype/lineweight is : " + en_clr + "/" + en_lt + "/" + en_lw + "/n");

                                // Get entity's layer properties
                                laytabrec = tx.GetObject(laytab[en.Layer], OpenMode.ForRead) as LayerTableRecord;
                                Color laytabrec_clr = laytabrec.Color;
                                ObjectId laytabrec_lt_oid = laytabrec.LinetypeObjectId;
                                LinetypeTableRecord ltrec = tx.GetObject(laytabrec_lt_oid, OpenMode.ForRead) as LinetypeTableRecord;
                                string laytabrec_lt_name = ltrec.Name;
                                LineWeight laytabrec_lw = laytabrec.LineWeight;
                                ed.WriteMessage("\nEntity layers color/linetype/lineweight is : " + laytabrec_clr + "/" + laytabrec_lt_name + "/" + laytabrec_lw);

                                // Final setting
                                if (en_clr == 256) { en.Color = laytabrec_clr; }
                                if (en_lt == "ByLayer") { en.Linetype = laytabrec_lt_name; }
                                if (en_lw == LineWeight.ByLayer) { en.LineWeight = laytabrec_lw; }

                                en.Layer = "0";
                            }
                        }
                        ed.Regen();
                        tx.Commit();
                    }
                }

                catch (Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    Application.ShowAlertDialog("Exception occured : \n" + ex.Message);
                }
            }

        }

        /////////////////////////////////////////////////////////

        // Wblock all blocks for optimisation/inspection
        // ISSUE : Replace dll path with active document path

        [CommandMethod("SEPARATEBLOCKS")]
        public void SeparateBlocks()
        {

            string path = "blocks"; //--will be used to create sub folder
            RXClass blockclass = RXClass.GetClass(typeof(BlockReference)); //--will be used to determine enttiy type

            // Determine if the path exists > Create the directory
            if (Directory.Exists(path))
            {
                ed.WriteMessage("\nThat path exists already.");
            }
            else
            {
                DirectoryInfo diri = Directory.CreateDirectory(path);
                ed.WriteMessage("\nThe directory created successfully.");
            }

            // Iterating Blocks > Wblock to path
            using (Transaction tx = db.TransactionManager.StartTransaction())
            { 
                BlockTable bt = tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btrec = tx.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;

                foreach (ObjectId oid in btrec)
                {
                    Entity ent = tx.GetObject(oid, OpenMode.ForRead) as Entity;

                    // way 1 : just keeping for further use
                    if(oid.ObjectClass == blockclass)
                    {
                        ed.WriteMessage("\nIts a Block..."); 
                    }
                    // way 2 : in action
                    if (ent is BlockReference)
                    {
                        ed.WriteMessage("\nReference...");

                        // Object handling
                        BlockReference bref = tx.GetObject(oid, OpenMode.ForRead) as BlockReference;

                        // Casting obj name to string > make anonymous block/illegal names suitable > check on console 
                        string bref_name = bref.Name;
                         
                        Regex illegalInFileName = new Regex(@"[\\/:*?""<>|]");
                        bref_name = illegalInFileName.Replace(bref_name, "_");

                        ed.WriteMessage("\n" + bref_name);

                        // Put obj in local collection
                        ObjectIdCollection oidcoll = new ObjectIdCollection();
                        oidcoll.Add(oid);

                        // Clone database and wblock
                        using (Database newdb = new Database(true, false))
                        {
                            db.Wblock(newdb, oidcoll, Point3d.Origin, DuplicateRecordCloning.Ignore);
                            string filename = path + "\\" + bref_name + ".dwg";
                            newdb.SaveAs(filename, DwgVersion.Newest);
                        }
                    }
                }
                
                tx.Commit();
            }
        }

        /////////////////////////////////////////////////////////

        // Extracts nested object from its block

        [CommandMethod("EXTRACTNESTEDOBJECT")]
        public void ExtractNestedObject()
        {

            try
            {
                // Prompt the user to select a nested object which is in a block
                PromptNestedEntityResult result = ed.GetNestedEntity("\nSelect an object which is nested in a block: ");
                if (result.Status != PromptStatus.OK) 
                    return;
                
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Open the selected entity for read
                    Entity ent = tr.GetObject(result.ObjectId, OpenMode.ForRead) as Entity; 

                    if (ent == null) 
                    {
                        ed.WriteMessage("\nInvalid object selected.");
                        return;
                    }

                    // Find the container block reference
                    BlockReference parentBlockReference = null; 
                    
                    foreach(ObjectId containerId in result.GetContainers())
                    {
                        BlockReference containerBlockReference = tr.GetObject(containerId, OpenMode.ForRead) as BlockReference;

                        if(containerBlockReference != null)
                        {
                            parentBlockReference = containerBlockReference;
                            break;
                        }
                    }

                    if (parentBlockReference == null)
                    {
                        ed.WriteMessage("\nThe selected object is not nested in a block.");
                        return;
                    }

                    // Clone the entity to the model space
                    Entity entClone = ent.Clone() as Entity;
                    if (entClone == null)
                    {
                        ed.WriteMessage("\nFailed to clone the object.");
                        return;
                    }

                    // Calculate the "combined" transformation
                    Matrix3d transform = Matrix3d.Identity;
                    foreach (ObjectId containerId in result.GetContainers())
                    {
                        BlockReference containerBlockReference = tr.GetObject(containerId, OpenMode.ForRead) as BlockReference;
                        if (containerBlockReference != null)
                        {
                            transform = containerBlockReference.BlockTransform * transform;
                        }
                    }

                    // Apply the transformation to the entity
                    entClone.TransformBy(transform);

                    // Add the new entity to the current space
                    BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    btr.AppendEntity(entClone);
                    tr.AddNewlyCreatedDBObject(entClone, true);

                    // Remove the original entity from the block definition
                    BlockTableRecord blockDef = tr.GetObject(parentBlockReference.BlockTableRecord, OpenMode.ForWrite) as BlockTableRecord;
                    if(blockDef == null)
                    {
                        ed.WriteMessage("\nFailed to open the block definition");
                        return;
                    }

                    ent.UpgradeOpen(); // --- Upgrade the object to write mode
                    ent.Erase(true);
                    parentBlockReference.RecordGraphicsModified(true);
                    

                    tr.Commit();
                    ed.Regen();
                    ed.WriteMessage("\nObject extracted successfully.");
                }
            }

            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError: " + ex.Message);
            }
        }

        /////////////////////////////////////////////////////////
        
        // Extract a nested block from its parent blocks

        [CommandMethod("EXTRACTNESTEDBLOCK")]
        public void ExtractNestedBlock()
        {

            try
            {
                // Prompt the user to select a nested object which is in a block
                PromptNestedEntityOptions pneoptions = new PromptNestedEntityOptions("\nSelect an object or block to extract from its immediate parent: ");
                pneoptions.AllowNone = false;
                PromptNestedEntityResult result = ed.GetNestedEntity(pneoptions);

                if (result.Status != PromptStatus.OK)
                    return;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Open the selected entity for read
                    Entity ent = tr.GetObject(result.ObjectId, OpenMode.ForRead) as Entity;
                    BlockReference firstParentBlock = null;
                    BlockReference secondParentBlock = null;
                    Entity entityToExtract = null;

                    if (ent == null)
                    {
                        ed.WriteMessage("\nInvalid object selected.");
                        return;
                    }

                    // Count the block references that contain the selected entity & Dialogue
                    ObjectId[] containerIds = result.GetContainers();
                    ed.WriteMessage("\nThe selected object is nested in " + containerIds.Length + " block(s).");

                    // If the object is not nested
                    if (containerIds.Length == 0)
                    {
                        ed.WriteMessage("\nThe selected object is not nested in a block.");
                        return;
                    }
                    else if (containerIds.Length == 1)
                    {
                        firstParentBlock = tr.GetObject(containerIds[0], OpenMode.ForRead) as BlockReference;
                    }
                    else
                    {
                        firstParentBlock = tr.GetObject(containerIds[0], OpenMode.ForRead) as BlockReference;
                        secondParentBlock = tr.GetObject(containerIds[1], OpenMode.ForRead) as BlockReference;
                    }

                    // Extract the object or its immediate parent block
                    entityToExtract = (containerIds.Length == 1) ? ent : firstParentBlock;

                    // Clone the object
                    Entity clonedEntity = entityToExtract.Clone() as Entity;

                    // Calculate the "local - independent" transform and reverse it
                    Matrix3d localTransformUndo = Matrix3d.Identity;
                    localTransformUndo = firstParentBlock.BlockTransform.Inverse();
                    
                    // Calculate the "combined" transformation
                    Matrix3d transform = Matrix3d.Identity;
                    foreach (ObjectId containerId in containerIds)
                    {
                        BlockReference containerBlockReference = tr.GetObject(containerId, OpenMode.ForRead) as BlockReference;
                        if (containerBlockReference != null)
                        {
                            //transform = transform.PreMultiplyBy(containerBlockReference.BlockTransform);
                            transform = containerBlockReference.BlockTransform * transform;
                        }

                    }

                    // Apply the transformations
                    clonedEntity.TransformBy(localTransformUndo);
                    clonedEntity.TransformBy(transform);

                    // Add the cloned entity to model space
                    BlockTableRecord modelSpace = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    modelSpace.AppendEntity(clonedEntity);
                    tr.AddNewlyCreatedDBObject(clonedEntity, true);

                    // Remove the original entity from its parent
                    if (containerIds.Length == 1)
                    {
                        ent.UpgradeOpen();
                        ent.Erase(true);
                        firstParentBlock.RecordGraphicsModified(true);
                    }
                    else
                    {
                        firstParentBlock.UpgradeOpen();
                        firstParentBlock.Erase(true);
                        secondParentBlock.RecordGraphicsModified(true);
                    }

                    tr.Commit();
                    ed.Regen();
                    ed.WriteMessage("\nObject extracted successfully.");

                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError: " + ex.Message);
            }
           
        }
    }
}
