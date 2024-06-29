using System; // OpenJoint

using Autodesk.AutoCAD.Runtime; 
using Autodesk.AutoCAD.ApplicationServices; 
using Autodesk.AutoCAD.DatabaseServices; 
using Autodesk.AutoCAD.EditorInput; 
using Autodesk.AutoCAD.Geometry; 


namespace J_Tools
{
    public class J_ObjectTools
    {
        // Sub-func - for OPENJOINT
        public static double AngleBetweenLines(Point3d s1, Point3d s2, Point3d e1, Point3d e2)
        {
            double theta1 = Math.Atan2(s1.Y - e1.Y, s1.X - e1.X);
            double theta2 = Math.Atan2(s2.Y - e2.Y, s2.X - e2.X);
            double diff = Math.Abs(theta1 - theta2);
            double angle = Math.Min(diff, Math.Abs(180 - diff));
            return angle;
        }

        // Sub-func
        static public void QueryObjects()
        {
            // Get the current document and database
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // Start a transaction
            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btrec = tx.GetObject(bt[BlockTableRecord.ModelSpace],OpenMode.ForRead) as BlockTableRecord;
                Editor ed = doc.Editor;

                // Step through the Block table record
                foreach (ObjectId oid in btrec)
                {
                    ed.WriteMessage("\nDXF name: " + oid.ObjectClass.DxfName);
                    ed.WriteMessage("\nObjectID: " + oid.ToString());
                    ed.WriteMessage("\nHandle: " + oid.Handle.ToString());
                    ed.WriteMessage("\n");
                }

                // Dispose of the transaction
            }
        }

        // Sub-func - for MCOPY
        static public void IterateCopy(DBObjectCollection dbocoll, Vector3d vec)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btrec = tx.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                foreach (Entity en in dbocoll)
                {
                    Entity enclone = en.Clone() as Entity;

                    enclone.TransformBy(Matrix3d.Displacement(vec));

                    btrec.AppendEntity(enclone);
                    tx.AddNewlyCreatedDBObject(enclone, true);
                }
                tx.Commit();
            }
        }

        /////////////////////////////////////////////////////////

        /// Copy center of arc/circle to clipboard for further operations.
        // 200501 ISSUE : obj UCS yi centera atasak yeter. Bu komuta ihtiyaç yok

        [CommandMethod("GETCENTER")]
        static public void GetCenter()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptEntityOptions peopt = new PromptEntityOptions("\nSelect curve to get center : ");
            peopt.SetRejectMessage("\nSelect only a curve.");
            peopt.AddAllowedClass(typeof(Circle), false);
            peopt.AddAllowedClass(typeof(Arc), false);
            peopt.AllowNone = false;

            PromptEntityResult peres = ed.GetEntity(peopt);

            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                BlockTableRecord btrec = tx.GetObject(peres.ObjectId, OpenMode.ForRead) as BlockTableRecord;
                Entity en = tx.GetObject(btrec.ObjectId, OpenMode.ForRead) as Entity;

                if(en is Arc)
                { }

                if(en is Circle)
                { }
            }
        }

        /////////////////////////////////////////////////////////

        /// Half open corner joint 
        // 200504 ISSUE : not suitable for obtuse angles

        [CommandMethod("OPENJOINT")]
        static public void OpenJoint()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                // Asking for user inputs and var casting
                PromptDoubleOptions pdopt = new PromptDoubleOptions("Enter open joint size : \n");
                pdopt.AllowZero = false;
                pdopt.AllowNegative = false;
                pdopt.DefaultValue = 0.5;
                pdopt.AllowNone = false;
                PromptDoubleResult pdres = ed.GetDouble(pdopt);
                if(pdres.Status != PromptStatus.OK) { return; }
                double size = pdres.Value;

                PromptEntityOptions peopt_1 = new PromptEntityOptions("\nSelect first line : \n");
                peopt_1.SetRejectMessage("\nSelect only lines");
                peopt_1.AddAllowedClass(typeof(Line), false);
                peopt_1.AllowNone = false;
                PromptEntityResult peres_1 = ed.GetEntity(peopt_1);
                if(peres_1.Status != PromptStatus.OK) { return;  }
                ObjectId line1_id = peres_1.ObjectId;
                ed.WriteMessage(line1_id.ToString());

                PromptEntityOptions peopt_2 = new PromptEntityOptions("\nSelect second line : \n");
                peopt_2.SetRejectMessage("\nSelect only lines");
                peopt_2.AddAllowedClass(typeof(Line), false);
                peopt_2.AllowNone = false;
                PromptEntityResult peres_2 = ed.GetEntity(peopt_2);
                if(peres_2.Status != PromptStatus.OK) { return; }
                ObjectId line2_id = peres_2.ObjectId;
                ed.WriteMessage(line2_id.ToString());

                // Oid > Entity > Line
                Entity line1_ent = tx.GetObject(line1_id, OpenMode.ForWrite) as Entity;
                Entity line2_ent = tx.GetObject(line2_id, OpenMode.ForWrite) as Entity;
                Line line1 = line1_ent as Line;
                Line line2 = line2_ent as Line;

                // Get intersection point > Check parallelity > re-orient lines
                Point3dCollection intptcoll = new Point3dCollection();

                line1.IntersectWith(line2, Intersect.OnBothOperands, intptcoll, 0, 0);

                if (intptcoll.Count != 1) { Application.ShowAlertDialog("Lines are parallel or not intersecting");
                    ed.WriteMessage(intptcoll.Count.ToString()); 
                    return; }
                Point3d intpt = intptcoll[0];  //---APEX
                
                if(line1.StartPoint.Equals(intpt)!=true) 
                {
                    line1.EndPoint = line1.StartPoint;
                    line1.StartPoint = intpt;
                }
                if(line2.StartPoint.Equals(intpt)!=true)
                {
                    line2.EndPoint = line2.StartPoint;
                    line2.StartPoint = intpt;
                }

                double anglin = AngleBetweenLines(line1.StartPoint, line2.StartPoint, line1.EndPoint, line2.EndPoint);
                ed.WriteMessage("\nAngle between lines : " + anglin);

                // Calculate line angles
                LineAngularDimension2 angdim12 = new LineAngularDimension2();
                angdim12.XLine1Start = line1.StartPoint;
                angdim12.XLine2Start = line2.StartPoint;
                angdim12.XLine1End = line1.EndPoint;
                angdim12.XLine2End = line2.EndPoint;
                
                // Derive bisector geometry
                Point3d midpt1 = line1.GetPointAtDist(size);
                Point3d midpt2 = line2.GetPointAtDist(size);
                Point3d ovpt1 = line1.GetPointAtDist(2 * size);
                Point3d ovpt2 = line2.GetPointAtDist(2 * size);
                Point3d midptbis = new Point3d((midpt1.X + midpt2.X) / 2.0, (midpt1.Y + midpt2.Y) / 2.0, (midpt1.Z + midpt2.Z) / 2.0);
                Point3d ovptbis = new Point3d((ovpt1.X + ovpt2.X) / 2.0, (ovpt1.Y + ovpt2.Y) / 2.0, (ovpt1.Z + ovpt2.Z) / 2.0);
                Line linebis = new Line(intpt, ovptbis);
                double hypothenusedistance = (size / Math.Cos(anglin/2)); 
                if(hypothenusedistance<0) { hypothenusedistance = hypothenusedistance * (-1); }
                ed.WriteMessage("\nhypothenusedistance : " + hypothenusedistance);
                Point3d endptbis = linebis.GetPointAtDist(hypothenusedistance); //<------

                BlockTableRecord btrec = tx.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                line1.StartPoint = midpt1;
                line2.StartPoint = midpt2;

                using (Line line3 = new Line(endptbis, midpt1))
                {
                    btrec.AppendEntity(line3);
                    tx.AddNewlyCreatedDBObject(line3, true);
                }

                using (Line line4 = new Line(endptbis, midpt2))
                {
                    btrec.AppendEntity(line4);
                    tx.AddNewlyCreatedDBObject(line4, true);
                }

                tx.Commit();
            }
        }

        /////////////////////////////////////////////////////////

        /// Multiple copy - Sketchup way

        [CommandMethod("MCOPY")]
        static public void Mcopy()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Select the objects to copy
            PromptSelectionOptions psopt = new PromptSelectionOptions();
            psopt.MessageForAdding = "\nSelect objects : ";
            psopt.AllowDuplicates = false;
            psopt.AllowSubSelections = false;
            PromptSelectionResult psres = ed.GetSelection(psopt);
            if (psres.Status != PromptStatus.OK) { return; }

            SelectionSet selset = psres.Value;
            //ObjectIdCollection oidcoll = new ObjectIdCollection(psres.Value.GetObjectIds

            // Get copy distance from user by two points
            PromptPointOptions ppopt_source = new PromptPointOptions("\n Pick source point : ");
            ppopt_source.AllowArbitraryInput = true;
            ppopt_source.AllowNone = false;
            PromptPointResult ppres_source = ed.GetPoint(ppopt_source);
            //ed.WriteMessage("\nSource point = " + ppres_source);
            Point3d pts = ppres_source.Value;
            PromptPointOptions ppopt_dest = new PromptPointOptions("\n Pick destination point : ");
            ppopt_dest.AllowArbitraryInput = true;
            ppopt_dest.AllowNone = false;
            PromptPointResult ppres_dest = ed.GetPoint(ppopt_dest);
            //ed.WriteMessage("\nSource point = " + ppres_dest);
            Point3d ptd = ppres_dest.Value;

            // Convert two points to a vector
            Vector3d vec = (ptd - pts);
            double vlen = vec.Length;
            ed.WriteMessage("\nVector3d : " + vec + ">> Length : " + vlen);

            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                DBObjectCollection dbocoll = new DBObjectCollection();

                // Iterate in selection set & cast in db.obj.collection
                foreach (SelectedObject selobj in selset)
                {
                    if (selobj == null)
                    { ed.WriteMessage("\nNull obj detected > ID : " + selobj.ObjectId); }

                    Entity ent = tx.GetObject(selobj.ObjectId, OpenMode.ForRead) as Entity;

                    dbocoll.Add(ent);
                    //ed.WriteMessage("\nObjId original : " + selobj.ObjectId);
                }

                // Copy objects in obj collection
                IterateCopy(dbocoll, vec);
                tx.Commit();

                // Ask for repetition number
                PromptIntegerOptions piopt = new PromptIntegerOptions("\nMultiply Copy : ");
                        piopt.DefaultValue = 1;
                        piopt.AllowZero = false;
                        piopt.AllowNone = true;
                        piopt.AllowNegative = false;
                        PromptIntegerResult pires = ed.GetInteger(piopt);
                int i_pires = pires.Value;

                // Repeat copy
                if(i_pires > 1)
                {
                    for(int i = i_pires; i>=2; i--)
                    {
                        ed.WriteMessage("\n int i : " + i);

                        Vector3d vec_new = vec * i;
                        ed.WriteMessage("\nnew vector length = " + vec_new.Length);

                        IterateCopy(dbocoll, vec_new);
                    }
                }

            }
        }

        /////////////////////////////////////////////////////////

        /// Matchprop reverse

        [CommandMethod("MATCHPROPREVERSE", CommandFlags.UsePickSet)]
        static public void MatchPropReverse()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptSelectionResult psres = ed.SelectImplied();

            // if there is no selected obj prior this command
            if(psres.Status == PromptStatus.Error)
            {
                // Select objects
                PromptSelectionOptions psopt = new PromptSelectionOptions();
                psopt.MessageForAdding = "\nSelect objects : ";
                psopt.AllowDuplicates = false;
                psopt.AllowSubSelections = false;
                psres = ed.GetSelection(psopt);
            }
            // if there is an obj selected prior this command
            else
            {
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            // if the user has not cancelled this operation
            if(psres.Status == PromptStatus.OK)
            {
                using (Transaction tx = db.TransactionManager.StartTransaction())
                {
                    //BlockTable bt = tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    //BlockTableRecord btrec = tx.GetObject(bt[BlockTableRecord.ModelSpace],OpenMode.ForWrite) as BlockTableRecord;

                    LayerTable laytab = tx.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                    LayerTableRecord laytabrec = null;

                    DBObjectCollection dbocoll = new DBObjectCollection();
                    //ObjectId[] oids = psres.Value.GetObjectIds();
                    SelectionSet selset = psres.Value;

                    foreach (SelectedObject selobj in selset)
                    {
                        if (selobj == null)
                        { ed.WriteMessage("\nNull obj detected > ID : " + selobj.ObjectId); }

                        Entity ent = tx.GetObject(selobj.ObjectId, OpenMode.ForWrite) as Entity;

                        dbocoll.Add(ent);
                        //ed.WriteMessage("\nObjId original : " + selobj.ObjectId);
                    }

                    // Select donor obj v2
                    PromptEntityResult per_d = ed.GetEntity("\nSelect source object : ");
                    if (per_d.Status != PromptStatus.OK) { return; }
                    Entity en_d = tx.GetObject(per_d.ObjectId, OpenMode.ForRead) as Entity;


                    // Get donor obj properties
                    string en_lyr = en_d.Layer;
                    int en_clr = en_d.ColorIndex;
                    string en_lt = en_d.Linetype;
                    LineWeight en_lw = en_d.LineWeight;
                    ed.WriteMessage("\nDonor layer/color/linetype/lineweight is : " + en_lyr + en_clr + "/" + en_lt + "/" + en_lw + "/n");

                    // Casting properties to selection set
                    foreach (Entity en in dbocoll)
                    {
                        en.Layer = en_lyr;
                        en.ColorIndex = en_clr;
                        en.Linetype = en_lt;
                        en.LineWeight = en_lw;

                    }
                    ed.Regen();
                    tx.Commit();
                }
            }
        }

        /////////////////////////////////////////////////////////

        /// Temporary point depictor

        [CommandMethod("POINTDEPICTOR")]
        static public void PointDepictor()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Create a TypedValue array to define the filter criteria
            TypedValue[] tyval = new TypedValue[1];
            tyval.SetValue(new TypedValue((int)DxfCode.Start, "POINT"), 0);

            // Assign the filter criteria to a SelectionFilter object
            SelectionFilter selfil = new SelectionFilter(tyval);

            // Selecting 
            PromptSelectionResult psres = ed.SelectAll(selfil);
            if(psres.Status != PromptStatus.OK) { return; }
            SelectionSet selset = psres.Value;
            ObjectId[] oids = selset.GetObjectIds();
            
            // Cast dbobj to a point and draw a circle each
            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord btrec = tx.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                foreach (ObjectId oid in oids)
                {
                    DBPoint pt = (DBPoint)tx.GetObject(oid, OpenMode.ForRead);
                    Point3d pt3 = (Point3d)pt.Position;

                    using (Circle circ = new Circle())
                    {
                        circ.Center = pt3;
                        circ.Radius = 5;
                        circ.ColorIndex = 1;

                        btrec.AppendEntity(circ);
                        tx.AddNewlyCreatedDBObject(circ, true);
                    }
                }

                tx.Commit();
            }

            // Pause for a keystroke
            PromptStringOptions pstopt = new PromptStringOptions("Press any key to continue...");
                pstopt.AllowSpaces = true;
                PromptResult pstres = ed.GetString(pstopt);

            // Undo previous
            doc.SendStringToExecute("undo 1 ",true,false,false);
        }

        /////////////////////////////////////////////////////////

        /// Zoom selected objects one by one

        [CommandMethod("ZOOMEACH",CommandFlags.UsePickSet)]
        static public void ZommEach()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            PromptSelectionResult psres = ed.SelectImplied();

            // if there is no selected obj prior this command
            if (psres.Status == PromptStatus.Error)
            {
                // Select previous
                psres = ed.SelectPrevious();
            }
            // if there is an obj selected prior this command
            else
            {
                ed.SetImpliedSelection(new ObjectId[0]);
            }

            SelectionSet selset = psres.Value;
            int selset_n = selset.Count;
            ObjectId[] oids = selset.GetObjectIds();
            ed.WriteMessage("\nNumber of selected objects : " + selset_n);

            // Abort operation if the user cancels selection
            if(psres.Status != PromptStatus.OK) { return; }

            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                using (ViewTableRecord view = ed.GetCurrentView())
                {
                    Extents3d exts;
                    int loopcount = selset_n - 1;

                    foreach (SelectedObject selobj in selset)
                    {
                        // Get entity and position
                        Entity ent = tx.GetObject(selobj.ObjectId, OpenMode.ForRead) as Entity;

                        exts = ent.GeometricExtents;

                        double exts_width = exts.MaxPoint.X - exts.MinPoint.X;
                        double exts_height = exts.MaxPoint.Y - exts.MinPoint.Y;
                        Point2d exts_center = new Point2d((exts.MaxPoint.X + exts.MinPoint.X) / 2, (exts.MaxPoint.Y + exts.MinPoint.Y) / 2);

                        view.Height = exts_height;
                        view.Width = exts_width;
                        view.CenterPoint = exts_center;

                        // Zoom
                        ed.SetCurrentView(view);

                        // Pause for a keystroke
                        PromptStringOptions pstopt = new PromptStringOptions(loopcount + " objects remaining. Press any key to continue...");
                        pstopt.AllowSpaces = true;
                        PromptResult pstres = ed.GetString(pstopt);

                        loopcount--;
                    }
                }
            tx.Commit();
            }
        }

        /////////////////////////////////////////////////////////
        ///

        /// Find the centroid of a closed polyline ///
        /// 

        [CommandMethod("POLYCENTROID")]

        static public void PolyCentroid()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Select a polyline object
            PromptEntityOptions peopt = new PromptEntityOptions("\nSelect a polyline : ");  // Prompt options
            peopt.SetRejectMessage("\nSelect only a polyline.");                            // Reject message
            peopt.AddAllowedClass(typeof(Polyline), false);                                 // Allow only polylines
            peopt.AllowNone = false;                                                        // User must select an object

            PromptEntityResult peres = ed.GetEntity(peopt);                                 // Get the result

            if (peres.Status != PromptStatus.OK) { return; }                                // If the user cancels the operation

            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                // Get the polyline object
                Polyline pl = tx.GetObject(peres.ObjectId, OpenMode.ForRead) as Polyline;
                if (pl == null)
                {return;}

                // Get the number of vertices
                int vn = pl.NumberOfVertices;

                // Get the centroid
                Point3d centroid = new Point3d(0, 0, 0);
                double area = 0;

                for (int i = 0; i < vn; i++)
                {
                    Point2d pt1 = pl.GetPoint2dAt(i);
                    Point2d pt2 = pl.GetPoint2dAt((i + 1) % vn);
                    double a = pt1.X * pt2.Y - pt2.X * pt1.Y;
                    area += a;
                    centroid = new Point3d(centroid.X + (pt1.X + pt2.X) * a, centroid.Y + (pt1.Y + pt2.Y) * a, 0);
                }

                area /= 2;
                centroid = new Point3d(centroid.X / (6 * area), centroid.Y / (6 * area), 0);

                // Display the centroid
                ed.WriteMessage("\nCentroid of the polyline is : " + centroid.ToString());

                // Put a 2d point at the centroid
                using (DBPoint dbpt = new DBPoint(centroid))
                {
                    BlockTable bt = tx.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord btrec = tx.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    btrec.AppendEntity(dbpt);
                    tx.AddNewlyCreatedDBObject(dbpt, true);
                }

                // Commit the transaction
                tx.Commit();

            }
        }


    }
}
