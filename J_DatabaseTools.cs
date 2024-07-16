using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Colors;


namespace J_Tools
{
    public class J_DatabaseTools : BaseCommand
    {
        // Create a new active layer by command bar
         
        [CommandMethod("NEWLAYERACTIVE")]
        public void NewLayerActive()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Getting user input and casting to var
            PromptStringOptions psopt = new PromptStringOptions("\nEnter new layer name : ");
            psopt.AllowSpaces = true;
            PromptResult pkres = ed.GetString(psopt);

            string layname = pkres.StringResult;

            // Checking if exists and creating a new layer
            using (Transaction tx = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = tx.GetObject(db.LayerTableId, OpenMode.ForWrite) as LayerTable;

                if(lt.Has(layname) == false)
                {
                    using (LayerTableRecord ltrec = new LayerTableRecord())
                    {
                        ltrec.Name = layname;
                        ltrec.Color = Color.FromColorIndex(ColorMethod.ByAci, 256);

                        lt.Add(ltrec);
                        tx.AddNewlyCreatedDBObject(ltrec, true);
                    }
                }
                
                // Making it current
                db.Clayer = lt[layname];
                tx.Commit();
            }
        }
    }
}
