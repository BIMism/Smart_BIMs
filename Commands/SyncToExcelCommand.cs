using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Smart_BIMs.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class SyncToExcelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ViewSchedule schedule = doc.ActiveView as ViewSchedule;
            if (schedule == null)
            {
                TaskDialog.Show("Error", "Please make a Schedule view active before syncing to Excel.");
                return Result.Failed;
            }

            try
            {
                ScheduleDefinition def = schedule.Definition;
                int fieldCount = def.GetFieldCount();
                List<ScheduleField> fields = new List<ScheduleField>();
                for (int i = 0; i < fieldCount; i++) fields.Add(def.GetField(i));

                var collectedElements = new FilteredElementCollector(doc, schedule.Id).ToElements();

                dynamic excelApp = null;
                bool isNew = false;
                try
                {
                    excelApp = COMHelper.GetActiveObject("Excel.Application");
                }
                catch
                {
                    Type t = Type.GetTypeFromProgID("Excel.Application");
                    excelApp = Activator.CreateInstance(t);
                    isNew = true;
                }

                excelApp.Visible = true;
                dynamic wb = null;
                if (isNew || excelApp.Workbooks.Count == 0)
                {
                    wb = excelApp.Workbooks.Add(System.Reflection.Missing.Value);
                    isNew = true;
                }
                else
                {
                    wb = excelApp.ActiveWorkbook;
                }

                dynamic ws = wb.ActiveSheet;

                if (isNew)
                {
                    ws.Name = "ScheduleLIVE";
                    // Initial Dump Export
                    ws.Cells[1, 1] = "ElementId";
                    for (int i = 0; i < fields.Count; i++) ws.Cells[1, i + 2] = fields[i].GetName();

                    dynamic headerRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, fields.Count + 1]];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.ColorIndex = 37;

                    int rows = collectedElements.Count;
                    int cols = fields.Count + 1;
                    if (rows > 0)
                    {
                        object[,] data = new object[rows, cols];
                        int r = 0;
                        foreach (Element el in collectedElements)
                        {
                            data[r, 0] = el.Id.Value.ToString();
                            for (int c = 0; c < fields.Count; c++)
                            {
                                Parameter p = null;
                                foreach (Parameter param in el.Parameters) { if (param.Id == fields[c].ParameterId) { p = param; break; } }
                                data[r, c + 1] = p != null ? (p.AsValueString() ?? p.AsString() ?? "") : "";
                            }
                            r++;
                        }
                        dynamic preciseRange = ws.Range[ws.Cells[2, 1], ws.Cells[rows + 1, cols]];
                        preciseRange.Value2 = data;
                    }
                    ws.Columns.AutoFit();
                }
                else
                {
                    // Sync to Existing Sheet (Update cells dynamically)
                    dynamic usedRange = ws.UsedRange;
                    object value2 = usedRange.Value2;
                    int existingRows = 1;
                    int existingCols = fields.Count + 1;
                    
                    Dictionary<string, int> colMap = new Dictionary<string, int>();

                    if (value2 is object[,] evalues)
                    {
                        existingRows = evalues.GetLength(0);
                        existingCols = Math.Max(existingCols, evalues.GetLength(1));
                        for (int c = 2; c <= evalues.GetLength(1); c++)
                        {
                            string header = evalues[1, c]?.ToString();
                            if (!string.IsNullOrEmpty(header)) colMap[header] = c;
                        }
                    }

                    Dictionary<string, int> rowMap = new Dictionary<string, int>();
                    if (value2 is object[,] v2)
                    {
                        for (int r = 2; r <= existingRows; r++)
                        {
                            string idStr = v2[r, 1]?.ToString();
                            if (!string.IsNullOrEmpty(idStr)) rowMap[idStr] = r;
                        }
                    }

                    int totalRowsNeeded = Math.Max(existingRows, rowMap.Count + collectedElements.Count + 1);
                    object[,] newData = new object[totalRowsNeeded, existingCols];

                    // Copy old Excel data
                    if (value2 is object[,] ov)
                    {
                        for (int r = 1; r <= existingRows; r++)
                            for (int c = 1; c <= ov.GetLength(1); c++)
                                newData[r - 1, c - 1] = ov[r, c];
                    }

                    int currentMaxRow = existingRows;

                    foreach (Element el in collectedElements)
                    {
                        string idStr = el.Id.Value.ToString();
                        int r = 0; // 0-indexed for array
                        if (rowMap.ContainsKey(idStr)) { r = rowMap[idStr] - 1; }
                        else { r = currentMaxRow; currentMaxRow++; newData[r, 0] = idStr; }

                        foreach (var f in fields)
                        {
                            if (colMap.ContainsKey(f.GetName()))
                            {
                                int c = colMap[f.GetName()] - 1;
                                Parameter p = null;
                                foreach (Parameter param in el.Parameters) { if (param.Id == f.ParameterId) { p = param; break; } }
                                newData[r, c] = p != null ? (p.AsValueString() ?? p.AsString() ?? "") : "";
                            }
                        }
                    }

                    dynamic updatedRange = ws.Range[ws.Cells[1, 1], ws.Cells[currentMaxRow, existingCols]];
                    updatedRange.Value2 = newData;
                    TaskDialog.Show("Live Sync Success", "Revit Schedule successfully pushed to the active Excel sheet!");
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                if (wb != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
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
