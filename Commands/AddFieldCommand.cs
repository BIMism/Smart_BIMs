using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Smart_BIMs.UI;

namespace Smart_BIMs.Commands
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class AddFieldCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            if (!(doc.ActiveView is ViewSchedule schedule))
            {
                TaskDialog.Show("Error", "Please open a Revit Schedule view first.");
                return Result.Failed;
            }

            AddFieldWindow window = new AddFieldWindow(doc, schedule);
            bool? result = window.ShowDialog();

            if (result == true && window.SelectedFields.Count > 0)
            {
                try
                {
                    AdvancedSyncCommand syncCoordinator = new AdvancedSyncCommand();
                    
                    // 1. Silent Sync to prevent data loss in Excel (Safeguard)
                    syncCoordinator.ExecuteSilentlySync(commandData);

                    // 2. Add Fields to Revit Schedule natively
                    using (Transaction trans = new Transaction(doc, "Live Add Field"))
                    {
                        trans.Start();
                        ScheduleDefinition def = schedule.Definition;
                        foreach (var field in window.SelectedFields)
                        {
                            try { def.AddField(field.Schedulable); } catch { }
                        }
                        trans.Commit();
                    }

                    // 3. Silent Export to strictly repopulate Excel with the new column natively
                    syncCoordinator.ExecuteSilentlyExport(commandData);
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Live Push Failed", ex.Message);
                    return Result.Failed;
                }
                
                return Result.Succeeded;
            }

            return Result.Cancelled;
        }
    }
}
