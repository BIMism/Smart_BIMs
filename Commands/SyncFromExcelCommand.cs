using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using Microsoft.Win32;

namespace Smart_BIMs.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class SyncFromExcelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ViewSchedule schedule = doc.ActiveView as ViewSchedule;
            if (schedule == null)
            {
                TaskDialog.Show("Error", "Please make the target Schedule view active before syncing.");
                return Result.Failed;
            }

            try
            {
                OpenFileDialog ofd = new OpenFileDialog();
                ofd.Filter = "Excel Files (*.xlsx)|*.xlsx";
                if (ofd.ShowDialog() == true)
                {
                    ScheduleDefinition def = schedule.Definition;
                    int fieldCount = def.GetFieldCount();
                    List<ScheduleField> fields = new List<ScheduleField>();
                    for (int i = 0; i < fieldCount; i++)
                    {
                        fields.Add(def.GetField(i));
                    }

                    int updatedElements = 0;

                    using (Transaction trans = new Transaction(doc, "Sync from Excel"))
                    {
                        trans.Start();

                        using (XLWorkbook workbook = new XLWorkbook(ofd.FileName))
                        {
                            IXLWorksheet ws = workbook.Worksheet(1);
                            
                            // Map columns from Excel header
                            Dictionary<int, ScheduleField> colMap = new Dictionary<int, ScheduleField>();
                            int maxCol = ws.LastColumnUsed().ColumnNumber();
                            for (int c = 2; c <= maxCol; c++)
                            {
                                string header = ws.Cell(1, c).GetString();
                                foreach (ScheduleField sf in fields)
                                {
                                    if (sf.GetName() == header)
                                    {
                                        colMap[c] = sf;
                                        break;
                                    }
                                }
                            }

                            int maxRow = ws.LastRowUsed().RowNumber();
                            for (int r = 2; r <= maxRow; r++)
                            {
                                string idStr = ws.Cell(r, 1).GetString();
                                if (int.TryParse(idStr, out int elementIdInt))
                                {
                                    ElementId id = new ElementId(elementIdInt);
                                    Element el = doc.GetElement(id);
                                    if (el != null)
                                    {
                                        bool updated = false;
                                        foreach (var kvp in colMap)
                                        {
                                            int col = kvp.Key;
                                            ScheduleField field = kvp.Value;
                                            string val = ws.Cell(r, col).GetString();

                                            Parameter p = el.get_Parameter(field.ParameterId);
                                            if (p != null && !p.IsReadOnly)
                                            {
                                                try { 
                                                    p.Set(val); 
                                                    updated = true; 
                                                } 
                                                catch { /* Ignore format mismatch natively */ }
                                            }
                                        }
                                        if (updated) updatedElements++;
                                    }
                                }
                            }
                        }
                        
                        trans.Commit();
                    }

                    TaskDialog.Show("Success", $"Successfully synced {updatedElements} element(s) with Excel data.");
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
