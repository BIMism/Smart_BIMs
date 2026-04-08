using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Excel = Microsoft.Office.Interop.Excel;

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

            Excel.Application excelApp = null;
            try
            {
                excelApp = (Excel.Application)COMHelper.GetActiveObject("Excel.Application");
            }
            catch
            {
                TaskDialog.Show("Live Sync Error", "Microsoft Excel is not currently running.\nYou must have Excel OPEN to perform a live sync.");
                return Result.Failed;
            }

            Excel.Workbook wb = excelApp.ActiveWorkbook;
            if (wb == null)
            {
                TaskDialog.Show("Live Sync Error", "There is no active workbook open in Excel.\nPlease ensure your schedule is open in Excel.");
                return Result.Failed;
            }

            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;

            try
            {
                ScheduleDefinition def = schedule.Definition;
                int fieldCount = def.GetFieldCount();
                List<ScheduleField> fields = new List<ScheduleField>();
                for (int i = 0; i < fieldCount; i++)
                {
                    fields.Add(def.GetField(i));
                }

                int updatedElements = 0;

                Excel.Range usedRange = ws.UsedRange;
                object[,] values = usedRange.Value2 as object[,];

                if (values != null && values.GetLength(0) > 1)
                {
                    int rowCount = values.GetLength(0);
                    int colCount = values.GetLength(1);

                    // Map columns
                    Dictionary<int, ScheduleField> colMap = new Dictionary<int, ScheduleField>();
                    for (int c = 2; c <= colCount; c++)
                    {
                        string header = values[1, c]?.ToString();
                        if (!string.IsNullOrEmpty(header))
                        {
                            foreach (ScheduleField sf in fields)
                            {
                                if (sf.GetName() == header)
                                {
                                    colMap[c] = sf;
                                    break;
                                }
                            }
                        }
                    }

                    using (Transaction trans = new Transaction(doc, "Live Sync from Excel"))
                    {
                        trans.Start();

                        for (int r = 2; r <= rowCount; r++)
                        {
                            string idStr = values[r, 1]?.ToString();
                            if (!string.IsNullOrEmpty(idStr) && long.TryParse(idStr, out long elementIdLong))
                            {
                                ElementId id = new ElementId(elementIdLong);
                                Element el = doc.GetElement(id);
                                if (el != null)
                                {
                                    bool updated = false;
                                    foreach (var kvp in colMap)
                                    {
                                        int col = kvp.Key;
                                        ScheduleField field = kvp.Value;
                                        string val = values[r, col]?.ToString() ?? "";

                                        Parameter p = null;
                                        foreach (Parameter param in el.Parameters)
                                        {
                                            if (param.Id == field.ParameterId) { p = param; break; }
                                        }

                                        if (p != null && !p.IsReadOnly)
                                        {
                                            try { p.Set(val); updated = true; }
                                            catch { /* Ignore format mismatch natively */ }
                                        }
                                    }
                                    if (updated) updatedElements++;
                                }
                            }
                        }
                        
                        trans.Commit();
                    }
                }

                TaskDialog.Show("Live Sync Success", $"Successfully synced {updatedElements} element(s) with Excel LIVE!");

                // Release COM
                System.Runtime.InteropServices.Marshal.ReleaseComObject(usedRange);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

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
