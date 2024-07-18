using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using Autodesk.AutoCAD.Runtime;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace J_Tools
{
    public class J_SelectionTools : BaseCommand
    {
        // 240717 - ISSUE : Does not work properly. Autocad Returned : "FATAL ERROR : Unhandled Access Violation Reading 0x0000 Exception at C0FE06BCh"

        [CommandMethod("SELECTSIMILARINVIEW")]
        public void SelectSimilarInView()
        {
            try
            {
                // Get the current value of the SELECTSIMILARMODE system variable
                int selectSimilarModeValue = Convert.ToInt32(Application.GetSystemVariable("SELECTSIMILARMODE"));

                // Find combinations of SELECTSIMILARMODE value
                List<List<int>> selectSimilarModeCombinations = SelectSimilarModeCombinationFinder.FindCombinations(selectSimilarModeValue);

                // Print the current value of the SELECTSIMILARMODE system variable (for debugging)
                ed.WriteMessage($"\nCurrent SELECTSIMILARMODE value: {selectSimilarModeValue}");

                // Print the combinations (for debugging)
                foreach (var combination in selectSimilarModeCombinations)
                {
                    ed.WriteMessage($"\nCombination: {string.Join(", ", combination)}");
                }

                // Start a transaction
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // Get the current view
                        ViewTableRecord view = ed.GetCurrentView();

                        if (view == null)
                        {
                            ed.WriteMessage("\nFailed to get the current view.");
                            return;
                        }

                        // Get the view boundaries
                        Point2d lowerLeft = view.CenterPoint - new Vector2d(view.Width / 2, view.Height / 2);
                        Point2d upperRight = view.CenterPoint + new Vector2d(view.Width / 2, view.Height / 2);

                        /// Debugging - Test function to get if the view boundaries are correct
                        /*
                        // Draw a rectangle as polyline segments
                        Autodesk.AutoCAD.DatabaseServices.Polyline rect = new Autodesk.AutoCAD.DatabaseServices.Polyline();
                        rect.AddVertexAt(0, lowerLeft, 0, 0, 0);
                        rect.AddVertexAt(1, new Point2d(lowerLeft.X, upperRight.Y), 0, 0, 0);
                        rect.AddVertexAt(2, upperRight, 0, 0, 0);
                        rect.AddVertexAt(3, new Point2d(upperRight.X, lowerLeft.Y), 0, 0, 0);
                        rect.Closed = true;

                        // Get the current block table record
                        BlockTableRecord btr = tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                        if (btr != null)
                        {
                            btr.AppendEntity(rect);
                            tr.AddNewlyCreatedDBObject(rect, true);
                        }
                        */

                        // Allow the user to select an object
                        PromptEntityOptions peo = new PromptEntityOptions("\nSelect an object: ");
                        PromptEntityResult per = ed.GetEntity(peo);

                        if (per.Status != PromptStatus.OK)
                        {
                            ed.WriteMessage("\nNo object selected.");
                            return;
                        }

                        Entity selectedEntity = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Entity;
                        if (selectedEntity == null)
                        {
                            ed.WriteMessage("\nSelected entity is null.");
                            return;
                        }
                        else // --- debugging
                        {
                            ed.WriteMessage($"\nSelected entity properties :");
                            ed.WriteMessage($"Type : {selectedEntity.GetType().Name}");
                            ed.WriteMessage($"Color : {selectedEntity.ColorIndex}");
                            ed.WriteMessage($"Layer : {selectedEntity.Layer}");
                        }

                        // Get the selection filterList based on the SELECTSIMILARMODE value
                        if (selectSimilarModeCombinations.Count > 0)
                        {
                            List<TypedValue> filterList = SelectSimilarModeFilter.GenerateSelectionFilter(selectSimilarModeCombinations[0], selectedEntity, ed);

                            // Create a selection filter
                            SelectionFilter filter = new SelectionFilter(filterList.ToArray());

                            // Debuging - Print the filterList
                            foreach (TypedValue tv in filterList)
                            {
                                ed.WriteMessage($"\nFilter: {tv.TypeCode} = {tv.Value}");
                            }

                            // Select objects within the view boundaries (lowerLeft, upperRight) and with the selection filter 
                            
                            PromptSelectionResult psr = ed.SelectWindow(
                                new Point3d(lowerLeft.X, lowerLeft.Y, 0),
                                new Point3d(upperRight.X, upperRight.Y, 0),
                                filter);
                            /*
                            // Debugging by neglecting filtering
                            PromptSelectionResult psr = ed.SelectWindow(
                                new Point3d(lowerLeft.X, lowerLeft.Y, 0),
                                new Point3d(upperRight.X, upperRight.Y, 0));
                            */

                            if (psr.Status == PromptStatus.OK)
                            {
                                ed.SetImpliedSelection(psr.Value.GetObjectIds());
                                ed.WriteMessage($"\nSelected {psr.Value.Count} similar object(s) in the current viewport.");
                            }
                            else
                            {
                                ed.WriteMessage("\nNo similar objects found in the current viewport.");
                            }
                        }
                        else
                        {
                            ed.WriteMessage("\nNo valid property combinations found for current SELECTSIMILARMODE value.");
                        }
                        
                        tr.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError during transaction: {ex.Message}");
                        ed.WriteMessage($"\nStack Trace: {ex.StackTrace}");
                        tr.Abort();
                    }
                }
            }

            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nUnexpected error: {ex.Message}");
                ed.WriteMessage($"\nStack Trace: {ex.StackTrace}");
            }
        }

        // Helper function - SELECTSIMILARINVIEWPORT
        // to find all combinations from the given set of numbers that add up to the selectSimilarModeValue

        public class SelectSimilarModeCombinationFinder
        {
            private static readonly int[] propertyvalues = { 1, 2, 4, 8, 16, 32, 64, 128 };
            
            // Get the SELECTSIMILARMODE value
            public static List<List<int>> FindCombinations(int selectSimilarModeValue)
            {
                List<List<int>> result = new List<List<int>>();
                FindCombinationRecursive(selectSimilarModeValue, 0, new List<int>(), result);
                return result;
            }

            // Recursive function to find all combinations of SELECTSIMILARMODE value
            private static void FindCombinationRecursive(int remainingSum, int startIndex, List<int> currentCombination, List<List<int>> result)
            {
                if (remainingSum == 0)
                {
                    result.Add(new List<int>(currentCombination));
                    return;
                }

                for (int i = startIndex; i < propertyvalues.Length; i++)
                {
                    if (propertyvalues[i] <= remainingSum)
                    {
                        currentCombination.Add(propertyvalues[i]);
                        FindCombinationRecursive(remainingSum - propertyvalues[i], i + 1, currentCombination, result);
                        currentCombination.RemoveAt(currentCombination.Count - 1);
                    }
                }
            }

        }

        // Helper function - SELECTSIMILARINVIEWPORT
        // to generate a selection filterList based on SELECTSIMILARMODE value

        public class SelectSimilarModeFilter
        {

            [Flags]
            public enum SelectSimilarMode
            {
                Color = 1,
                Layer = 2,
                Linetype = 4,
                LinetypeScale = 8,
                Lineweight = 16,
                PlotStyle = 32,
                ObjectStyle = 64,
                Name = 128
            }

            // Generate a selection filterList based on the combination and the selected entity
            public static List<TypedValue> GenerateSelectionFilter(List<int> combination, Entity selectedEntity, Editor ed)
            {
                List<TypedValue> filterList = new List<TypedValue>();

                try
                {


                    // Always add the object type to the filter
                    //filterList.Add(new TypedValue((int)DxfCode.Start, RXClass.GetClass(selectedEntity.GetType()).DxfName));
                    //filterList.Add(new TypedValue((int)DxfCode.Start, selectedEntity.GetType().Name));
                    filterList.Add(new TypedValue((int)DxfCode.Start, "*"));

                    // Loop through the solved combination of SELECTSIMILARMODE value and add the appropriate filter
                    foreach (int value in combination)
                    {
                        switch (value)
                        {
                            case 1:
                                filterList.Add(new TypedValue((int)DxfCode.Color, selectedEntity.ColorIndex));
                                break;
                            case 2:
                                filterList.Add(new TypedValue((int)DxfCode.LayerName, selectedEntity.Layer));
                                break;
                            case 4:
                                filterList.Add(new TypedValue((int)DxfCode.LinetypeName, selectedEntity.Linetype));
                                break;
                            case 8:
                                filterList.Add(new TypedValue((int)DxfCode.LinetypeScale, selectedEntity.LinetypeScale));
                                break;
                            case 16:
                                filterList.Add(new TypedValue((int)DxfCode.LineWeight, selectedEntity.LineWeight));
                                break;
                            case 32:
                                filterList.Add(new TypedValue((int)DxfCode.PlotStyleNameId, selectedEntity.PlotStyleName));
                                break;
                            case 64:
                                AddObjectStyleFilter(filterList, selectedEntity);
                                break;
                            case 128:
                                AddNameFilter(filterList, selectedEntity);
                                break;
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError in GenerateSelectionFilter: {ex.Message}");
                    ed.WriteMessage($"\nStack Trace: {ex.StackTrace}");
                }

                return filterList;
            }

            // Sub-Helper Function - Add the object style filter based on the selected entity type
            private static void AddObjectStyleFilter(List<TypedValue> filterList, Entity selectedEntity)
            {
                if (selectedEntity is DBText text)
                {
                    filterList.Add(new TypedValue((int)DxfCode.TextStyleName, text.TextStyleName));
                }
                else if (selectedEntity is MText mtext)
                {
                    filterList.Add(new TypedValue((int)DxfCode.TextStyleName, mtext.TextStyleName));
                }
                else if (selectedEntity is Dimension dim)
                {
                    filterList.Add(new TypedValue((int)DxfCode.DimStyleName, dim.DimensionStyleName));
                }
                else if (selectedEntity is Leader leader)
                {
                    filterList.Add(new TypedValue((int)DxfCode.DimStyleName, leader.DimensionStyleName));
                }
            }

            // Sub-Helper Function - Add the name filter based on the selected entity type
            private static void AddNameFilter(List<TypedValue> filterList, Entity selectedEntity)
            {
                if (selectedEntity is BlockReference blockReference)
                {
                    filterList.Add(new TypedValue((int)DxfCode.BlockName, blockReference.Name));
                }
                else if (selectedEntity is AttributeDefinition attributeDefinition)
                {
                    filterList.Add(new TypedValue((int)DxfCode.BlockName, attributeDefinition.Tag));
                }
                else if (selectedEntity is AttributeReference attributeReference)
                {
                    filterList.Add(new TypedValue((int)DxfCode.BlockName, attributeReference.Tag));
                }
            }
        }

    }
}
