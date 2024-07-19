using System.Linq;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;

namespace J_Tools
{
    public class J_FileTools : BaseCommand
    {
        // Close all documents open without saving and bothering the user with prompts
        [CommandMethod("CLOSEALLUNSAVED", CommandFlags.Session)]
        public void CloseAllUnsaved()
        {
            // Get the current document and application
            if (doc == null) return; // --- Exit if no document is open

            // Get all open documents and store in an array
            Document[] openDocs = Application.DocumentManager.Cast<Document>().ToArray();

            // Iterate through the array of open documents and close all unsaved documents
            foreach (Document d in openDocs)
            {
                // Check if it's not the current document
                if (!d.IsActive)
                {
                    try
                    {
                        ed.WriteMessage("\nAttempting to closing inactive document " + d.Name + "..."); // --- debugging

                        // Close the document
                        d.CloseAndDiscard();
                        
                        ed.WriteMessage("\nDocument " + d.Name + " closed without saving..."); // --- debugging
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage("\nError closing inactive document: " + ex.Message);
                    }
                }
                else
                {
                    try
                    {
                        ed.WriteMessage("\nActive document " + d.Name + " is closing..."); // --- debugging
                        d.CloseAndDiscard();
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage("\nError closing active document: " + ex.Message);
                    }
                }
            }
        }
    }
}
