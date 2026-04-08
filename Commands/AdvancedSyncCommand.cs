using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Smart_BIMs.UI;

namespace Smart_BIMs.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class AdvancedSyncCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ViewSchedule schedule = doc.ActiveView as ViewSchedule;
            if (schedule == null)
            {
                TaskDialog.Show("Error", "Please make a Schedule view active before using the Advanced Excel Manager.");
                return Result.Failed;
            }

            try
            {
                ExcelManagerWindow ui = new ExcelManagerWindow(doc, schedule);
                ui.ShowDialog();

                if (!ui.DoExport && !ui.DoSync) return Result.Cancelled;

                // Validate selections
                bool hasSelections = false;
                foreach (var item in ui.AvailableFields) { if (item.IsSelected) hasSelections = true; }

                if (!hasSelections)
                {
                    TaskDialog.Show("Warning", "No fields were selected for syncing.");
                    return Result.Cancelled;
                }

                // Inject unassigned fields into Revit Schedule
                using (Transaction t = new Transaction(doc, "Add Fields to Schedule"))
                {
                    t.Start();
                    ScheduleDefinition def = schedule.Definition;
                    HashSet<ElementId> existingFieldIds = new HashSet<ElementId>();
                    for (int i = 0; i < def.GetFieldCount(); i++) existingFieldIds.Add(def.GetField(i).ParameterId);

                    foreach (var item in ui.AvailableFields)
                    {
                        if (item.IsSelected && !existingFieldIds.Contains(item.Schedulable.ParameterId))
                        {
                            try { def.AddField(item.Schedulable); } catch { }
                        }
                    }
                    t.Commit();
                }

                // Formulate final exportable fields
                List<ScheduleField> fields = new List<ScheduleField>();
                ScheduleDefinition updatedDef = schedule.Definition;
                for (int i = 0; i < updatedDef.GetFieldCount(); i++)
                {
                    ScheduleField sf = updatedDef.GetField(i);
                    var matchItem = ui.AvailableFields.FirstOrDefault(x => x.Schedulable.ParameterId == sf.ParameterId);
                    if (matchItem != null && matchItem.IsSelected)
                    {
                        fields.Add(sf);
                    }
                }

                dynamic excelApp = null;
                bool isNew = false;
                try { excelApp = COMHelper.GetActiveObject("Excel.Application"); }
                catch { Type t = Type.GetTypeFromProgID("Excel.Application"); excelApp = Activator.CreateInstance(t); isNew = true; }

                excelApp.Visible = true;
                dynamic wb = null;
                if (isNew || excelApp.Workbooks.Count == 0) { wb = excelApp.Workbooks.Add(System.Reflection.Missing.Value); isNew = true; }
                else { wb = excelApp.ActiveWorkbook; }

                dynamic ws = wb.ActiveSheet;

                if (ui.DoExport)
                {
                    ExportToExcelLive(doc, schedule, fields, ui, excelApp, wb, ws, isNew);
                }
                else if (ui.DoSync)
                {
                    SyncFromExcelLive(doc, schedule, fields, ws);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                    if (wb != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void ExportToExcelLive(Document doc, ViewSchedule schedule, List<ScheduleField> fields, ExcelManagerWindow ui, dynamic excelApp, dynamic wb, dynamic ws, bool isNew)
        {
            try { ws.Unprotect(); } catch { }
            var collectedElements = new FilteredElementCollector(doc, schedule.Id).ToElements();

            // Determine Read-Only status from first element
            bool[] isReadOnlyCol = new bool[fields.Count];
            if (collectedElements.Count > 0)
            {
                Element firstEl = collectedElements.First();
                for (int i = 0; i < fields.Count; i++)
                {
                    Parameter p = null;
                    foreach (Parameter param in firstEl.Parameters) { if (param.Id == fields[i].ParameterId) { p = param; break; } }
                    isReadOnlyCol[i] = p != null ? p.IsReadOnly : true;
                }
            }

            int existingRows = 1;
            int existingCols = fields.Count + 1;

            if (isNew)
            {
                ws.Name = "ScheduleAdvanced";
                ws.Cells[1, 1] = "ElementId";
                for (int i = 0; i < fields.Count; i++) ws.Cells[1, i + 2] = fields[i].GetName();

                int rows = collectedElements.Count;
                if (rows > 0)
                {
                    object[,] data = new object[rows, existingCols];
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
                    dynamic preciseRange = ws.Range[ws.Cells[2, 1], ws.Cells[rows + 1, existingCols]];
                    preciseRange.Value2 = data;
                }
                existingRows = rows + 1;
            }
            else
            {
                int maxExpectedRows = collectedElements.Count + 2000;
                int totalRows = ws.UsedRange.Rows.Count;
                int totalCols = ws.UsedRange.Columns.Count;

                if (totalRows > maxExpectedRows) totalRows = maxExpectedRows;
                if (totalCols > 256) totalCols = 256;

                dynamic safeRange = ws.Range[ws.Cells[1, 1], ws.Cells[totalRows, totalCols]];
                object value2 = safeRange.Value2;

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
                    int r = 0; 
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
                existingRows = currentMaxRow;
                dynamic updatedRange = ws.Range[ws.Cells[1, 1], ws.Cells[currentMaxRow, existingCols]];
                updatedRange.Value2 = newData;
            }

            // Apply Settings from UI
            ws.Columns.AutoFit();
            dynamic fullGrid = ws.Range[ws.Cells[1, 1], ws.Cells[existingRows, existingCols]];

            // 1. Gridlines / Borders
            if (ui.chkSyncBorders.IsChecked == true)
            {
                fullGrid.Borders.LineStyle = 1; // xlContinuous
            }
            else
            {
                fullGrid.Borders.LineStyle = -4142; // xlNone
            }

            // 2. Clear previous interior colors before applying shades/stripes
            fullGrid.Interior.ColorIndex = -4142;

            // 3. Shading (Gray for Read-Only) mapped safely
            if (ui.chkSyncShading.IsChecked == true)
            {
                ws.Cells.Locked = false;
                ws.Columns[1].Locked = true;
                ws.Columns[1].Interior.ColorIndex = 15;

                for (int i = 0; i < fields.Count; i++)
                {
                    if (isReadOnlyCol[i])
                    {
                        ws.Columns[i + 2].Locked = true;
                        ws.Columns[i + 2].Interior.ColorIndex = 15;
                    }
                }
            }

            // 4. Stripe Rows Pattern
            if (ui.chkStripeRows.IsChecked == true)
            {
                for (int r = 2; r <= existingRows; r += 2)
                {
                    dynamic rowRange = ws.Range[ws.Cells[r, 2], ws.Cells[r, existingCols]];
                    // Avoid overriding gray ReadOnly columns with white stripe! Just stripe unlocked cells.
                    for (int c = 2; c <= existingCols; c++)
                    {
                        dynamic cell = ws.Cells[r, c];
                        if (cell.Interior.ColorIndex == -4142) // if it has no background color yet
                        {
                            cell.Interior.ColorIndex = 24; // Light Blue/Gray Stripe
                        }
                    }
                }
            }

            // Export All Available Parameters to Hidden Sheet for Validation Dictionary
            IList<SchedulableField> allSchedulable = schedule.Definition.GetSchedulableFields();
            List<string> paramNames = new List<string>();
            foreach(var sf in allSchedulable)
            {
                string pName = sf.GetName(doc).Replace("'", "").Replace("=", "");
                if (!string.IsNullOrEmpty(pName) && !paramNames.Contains(pName)) paramNames.Add(pName);
            }

            dynamic dictWs = null;
            try { dictWs = wb.Worksheets["SmartBIM_Dictionary"]; }
            catch { dictWs = wb.Worksheets.Add(After: ws); dictWs.Name = "SmartBIM_Dictionary"; dictWs.Visible = 2; /*xlSheetVeryHidden*/ }
            
            if (paramNames.Count > 0)
            {
                object[,] dictData = new object[paramNames.Count, 1];
                for(int i=0; i<paramNames.Count; i++) dictData[i,0] = paramNames[i];
                dynamic dictRange = dictWs.Range[dictWs.Cells[1,1], dictWs.Cells[paramNames.Count, 1]];
                dictRange.Value2 = dictData;
            }

            ws.Activate(); // Ensure main sheet is active before validation
            
            // Allow Dropdown Additions
            dynamic valRange = ws.Range[ws.Cells[1, existingCols + 1], ws.Cells[1, existingCols + 10]];
            dynamic newCellDataRange = ws.Range[ws.Cells[2, existingCols + 1], ws.Cells[existingRows, existingCols + 10]];
            
            valRange.Locked = false;
            newCellDataRange.Locked = false;
            
            try
            {
                valRange.Validation.Delete();
                valRange.Validation.Add(Type: 3 /*xlValidateList*/, AlertStyle: 1, Operator: 1, Formula1: $"=SmartBIM_Dictionary!$A$1:$A${paramNames.Count}");
                valRange.Interior.ColorIndex = 36; // Light yellow background
                valRange.Value2 = "Add new...";
            }
            catch { /* Regional formula separator edge cases */ }

            // Header Stylings
            dynamic hdrRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, existingCols]];
            hdrRange.Font.Bold = true;
            hdrRange.Interior.ColorIndex = 37;

            // 5. Fonts (Standard uniform mapping)
            if (ui.chkSyncFonts.IsChecked == true)
            {
                fullGrid.Font.Name = "Arial"; // Safe default corresponding to basic Revit schedule font
            }

            // 6. Freeze Panes
            if (ui.chkFreezeHeader.IsChecked == true)
            {
                try
                {
                    ws.Activate();
                    excelApp.ActiveWindow.SplitRow = 1;
                    excelApp.ActiveWindow.FreezePanes = true;
                }
                catch { }
            }

            // Final sheet protection
            if (ui.chkSyncShading.IsChecked == true)
            {
                ws.Protect(AllowFormattingColumns: true, AllowFormattingRows: true, AllowSorting: true, AllowFiltering: true);
            }

            TaskDialog.Show("Advanced Export Success", "Revit Schedule successfully mapped and pushed to Excel!");

            System.Runtime.InteropServices.Marshal.ReleaseComObject(hdrRange);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(fullGrid);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            if (wb != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        private void SyncFromExcelLive(Document doc, ViewSchedule schedule, List<ScheduleField> fields, dynamic ws)
        {
            int scheduleElementCount = new FilteredElementCollector(doc, schedule.Id).ToElementIds().Count;
            int maxExpectedRows = scheduleElementCount + 2000;
            int totalRows = ws.UsedRange.Rows.Count;
            int totalCols = ws.UsedRange.Columns.Count;

            if (totalRows > maxExpectedRows) totalRows = maxExpectedRows;
            if (totalCols > 256) totalCols = 256;

            dynamic safeRange = ws.Range[ws.Cells[1, 1], ws.Cells[totalRows, totalCols]];
            object value2 = safeRange.Value2;

            if (value2 is object[,] values && values.GetLength(0) > 1)
            {
                int rowCount = values.GetLength(0);
                int colCount = values.GetLength(1);

                Dictionary<int, ScheduleField> colMap = new Dictionary<int, ScheduleField>();
                IList<SchedulableField> allSchedFields = schedule.Definition.GetSchedulableFields();

                int updatedElements = 0;
                using (Transaction trans = new Transaction(doc, "Advanced Live Sync"))
                {
                    trans.Start();

                    // Parse headers and dynamically inject missing fields
                    for (int c = 2; c <= colCount; c++)
                    {
                        string header = values[1, c]?.ToString();
                        if (!string.IsNullOrEmpty(header) && header != "Add new...")
                        {
                            ScheduleField matched = fields.FirstOrDefault(f => f.GetName() == header);
                            if (matched != null)
                            {
                                colMap[c] = matched;
                            }
                            else
                            {
                                SchedulableField sFieldToInject = null;
                                foreach (var sf in allSchedFields)
                                {
                                    if (sf.GetName(doc) == header) { sFieldToInject = sf; break; }
                                }

                                if (sFieldToInject != null)
                                {
                                    try
                                    {
                                        ScheduleField newlyAdded = schedule.Definition.AddField(sFieldToInject);
                                        colMap[c] = newlyAdded;
                                    }
                                    catch { }
                                }
                            }
                        }
                    }

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
                                    ScheduleField field = kvp.Value;
                                    string val = values[r, kvp.Key]?.ToString() ?? "";

                                    Parameter p = null;
                                    foreach (Parameter param in el.Parameters) { if (param.Id == field.ParameterId) { p = param; break; } }
                                    if (p != null && !p.IsReadOnly)
                                    {
                                        try { p.Set(val); updated = true; } catch { }
                                    }
                                }
                                if (updated) updatedElements++;
                            }
                        }
                    }
                    trans.Commit();
                }
                TaskDialog.Show("Advanced Sync Success", $"Successfully synced {updatedElements} element(s) exclusively for selected fields!");
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(safeRange);
        }
    }
}
