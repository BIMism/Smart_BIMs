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

            int existingRows = collectedElements.Count > 0 ? collectedElements.Count : 1;
            int existingCols = fields.Count + 1; // +1 for ElementId

            // 1. Title Row
            ws.Cells[1, 1].Value2 = schedule.Name;
            dynamic titleRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, existingCols + 3]];
            titleRange.Merge();
            titleRange.Font.Size = 16;
            titleRange.Font.Bold = true;
            titleRange.HorizontalAlignment = -4108; // xlCenter

            // 2. Headers (Moved to Row 2)
            object[,] headerData = new object[1, existingCols];
            headerData[0, 0] = "ElementId";
            for (int i = 0; i < fields.Count; i++) headerData[0, i + 1] = fields[i].ColumnHeading ?? fields[i].GetName();
            
            dynamic hdrRange = ws.Range[ws.Cells[2, 1], ws.Cells[2, existingCols]];
            hdrRange.Value2 = headerData;

            if (collectedElements.Count == 0) return; // Guard clause

            object[,] exportData = new object[existingRows, existingCols];

            for (int r = 0; r < collectedElements.Count; r++)
            {
                Element el = collectedElements[r];
                exportData[r, 0] = el.Id.IntegerValue.ToString();

                for (int c = 0; c < fields.Count; c++)
                {
                    ScheduleField field = fields[c];
                    Parameter param = null;
                    foreach (Parameter pI in el.Parameters) { if (pI.Id == field.ParameterId) { param = pI; break; } }
                    
                    if (param == null)
                    {
                        ElementId typeId = el.GetTypeId();
                        if (typeId != ElementId.InvalidElementId)
                        {
                            ElementType eType = doc.GetElement(typeId) as ElementType;
                            if (eType != null) {
                                foreach (Parameter pI in eType.Parameters) { if (pI.Id == field.ParameterId) { param = pI; break; } }
                            }
                        }
                    }

                    if (param != null)
                    {
                        if (r == 0) isReadOnlyCol[c] = param.IsReadOnly;
                        string val = param.AsValueString();
                        if (string.IsNullOrEmpty(val)) val = param.AsString();
                        if (string.IsNullOrEmpty(val)) {
                            if (param.StorageType == StorageType.Double) val = param.AsDouble().ToString("0.##");
                            else if (param.StorageType == StorageType.Integer) val = param.AsInteger().ToString();
                        }
                        exportData[r, c + 1] = val;
                    }
                }
            }
            
            // Write Data (Moved to Row 3)
            dynamic dataRange = ws.Range[ws.Cells[3, 1], ws.Cells[2 + existingRows, existingCols]];
            dataRange.Value2 = exportData;

            // Apply Settings from UI
            ws.Columns.AutoFit();
            dynamic fullGrid = ws.Range[ws.Cells[2, 1], ws.Cells[2 + existingRows, existingCols]];

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
                for (int r = 3; r <= 2 + existingRows; r += 2)
                {
                    dynamic rowRange = ws.Range[ws.Cells[r, 1], ws.Cells[r, existingCols]];
                    foreach (dynamic cell in rowRange.Cells)
                    {
                        if (cell.Interior.ColorIndex == -4142 || cell.Interior.ColorIndex == 2)
                        {
                            cell.Interior.ColorIndex = 24; // xlThemeColorAccent5 light tint
                        }
                    }
                }
            }

            // Export All Available Parameters to Hidden Sheet for Validation Dictionary
            HashSet<string> existingNames = new HashSet<string>();
            foreach (var f in fields) existingNames.Add(f.ColumnHeading);

            IList<SchedulableField> allSchedulable = schedule.Definition.GetSchedulableFields();
            List<string> paramNames = new List<string>();
            List<bool> paramReadOnly = new List<bool>();

            foreach(var sf in allSchedulable)
            {
                string pName = sf.GetName(doc).Replace("'", "").Replace("=", "");
                if (existingNames.Contains(pName)) continue;

                if (!string.IsNullOrEmpty(pName) && !paramNames.Contains(pName))
                {
                    paramNames.Add(pName);
                    bool isRO = true;
                    if (collectedElements.Count > 0)
                    {
                        Element firstEl = collectedElements.First();
                        Parameter p = null;
                        foreach (Parameter pI in firstEl.Parameters) { if (pI.Id == sf.ParameterId) { p = pI; break; } }
                        if (p != null) isRO = p.IsReadOnly;
                    }
                    paramReadOnly.Add(isRO);
                }
            }

            dynamic dictWs = null;
            try { dictWs = wb.Worksheets["SmartBIM_Dictionary"]; }
            catch { dictWs = wb.Worksheets.Add(After: ws); dictWs.Name = "SmartBIM_Dictionary"; dictWs.Visible = 2; /*xlSheetVeryHidden*/ }
            
            if (paramNames.Count > 0)
            {
                object[,] dictData = new object[paramNames.Count, 2];
                for(int i=0; i<paramNames.Count; i++) 
                {
                    dictData[i,0] = paramNames[i];
                    dictData[i,1] = paramReadOnly[i] ? "True" : "False";
                }
                dynamic dictRange = dictWs.Range[dictWs.Cells[1,1], dictWs.Cells[paramNames.Count, 2]];
                dictRange.Value2 = dictData;
            }

            ws.Activate(); // Ensure main sheet is active before validation
            
            // Allow Dropdown Additions
            int maxR = existingRows < 2 ? 100 : existingRows;
            dynamic valRange = ws.Range[ws.Cells[2, existingCols + 1], ws.Cells[2, existingCols + 3]];
            dynamic newCellDataRange = ws.Range[ws.Cells[3, existingCols + 1], ws.Cells[2 + maxR, existingCols + 3]];
            
            valRange.Locked = false;
            newCellDataRange.Locked = false;
            
            try
            {
                valRange.Validation.Delete();
                valRange.Validation.Add(Type: 3 /*xlValidateList*/, AlertStyle: 1, Operator: 1, Formula1: $"=SmartBIM_Dictionary!$A$1:$A${Math.Max(1, paramNames.Count)}");
                valRange.Value2 = "[ Add Parameter ]";
                valRange.Font.Italic = true;
                valRange.Font.ColorIndex = 16; // Dark grey font
                valRange.Interior.ColorIndex = 20; // Light Blue
                valRange.Borders.LineStyle = 1;
                
                int colRef = existingCols + 1;
                string colStr = "";
                while(colRef > 0) { int m = (colRef-1)%26; colStr = Convert.ToChar('A'+m) + colStr; colRef = (colRef-m)/26; }

                newCellDataRange.Validation.Delete();
                newCellDataRange.Validation.Add(Type: 7 /*xlValidateCustom*/, AlertStyle: 1, Operator: 1, 
                    Formula1: $"=IFERROR(VLOOKUP({colStr}$2, SmartBIM_Dictionary!$A$1:$B$1000, 2, FALSE), \"True\")<>\"True\"");
                newCellDataRange.Validation.ErrorTitle = "Read-Only Parameter";
                newCellDataRange.Validation.ErrorMessage = "This parameter is generated by Revit natively and is strictly Read-Only.";

                dynamic formatCond = newCellDataRange.FormatConditions.Add(Type: 2 /*xlExpression*/, 
                    Formula1: $"=IFERROR(VLOOKUP({colStr}$2, SmartBIM_Dictionary!$A$1:$B$1000, 2, FALSE), \"False\")=\"True\"");
                formatCond.Interior.ColorIndex = 15; // Gray to show it is locked
            }
            catch { /* Regional formula separator edge cases */ }

            // Header Stylings
            dynamic hdrRangeFormat = ws.Range[ws.Cells[2, 1], ws.Cells[2, existingCols]];
            hdrRangeFormat.Font.Bold = true;
            hdrRangeFormat.Borders.LineStyle = 1;
            hdrRangeFormat.Interior.ColorIndex = 37;

            // 5. Fonts (Standard uniform mapping)
            if (ui.chkSyncFonts.IsChecked == true)
            {
                dynamic overallGrid = ws.Range[ws.Cells[2, 1], ws.Cells[2 + existingRows, existingCols + 3]];
                overallGrid.Font.Name = "Arial"; // Safe default corresponding to basic Revit schedule font
            }

            // 6. Freeze Panes
            if (ui.chkFreezeHeader.IsChecked == true)
            {
                ws.Activate();
                ws.Application.ActiveWindow.SplitRow = 2; // Offset for Title
                ws.Application.ActiveWindow.FreezePanes = true;
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
            int totalRows = ws.UsedRange.Rows.Count;
            int totalCols = ws.UsedRange.Columns.Count;

            dynamic safeRange = ws.Range[ws.Cells[1, 1], ws.Cells[totalRows, totalCols]];
            object value2 = safeRange.Value2;

            if (value2 is object[,] values && values.GetLength(0) > 2)
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
                        string header = values[2, c]?.ToString();
                        if (!string.IsNullOrEmpty(header) && header.Trim() != "")
                        {
                            ScheduleField matched = fields.FirstOrDefault(f => f.ColumnHeading == header);
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

                    for (int r = 3; r <= rowCount; r++)
                    {
                        string idStr = values[r, 1]?.ToString();
                        if (!string.IsNullOrEmpty(idStr) && int.TryParse(idStr, out int elementIdInt))
                        {
                            ElementId eId = new ElementId(elementIdInt);
                            Element el = doc.GetElement(eId);
                            if (el != null)
                            {
                                bool updated = false;
                                for (int c = 2; c <= colCount; c++)
                                {
                                    if (colMap.ContainsKey(c))
                                    {
                                        ScheduleField matchedField = colMap[c];
                                        Parameter p = null;
                                        foreach (Parameter pI in el.Parameters) { if (pI.Id == matchedField.ParameterId) { p = pI; break; } }

                                        if (p == null)
                                        {
                                            ElementId typeId = el.GetTypeId();
                                            if (typeId != ElementId.InvalidElementId)
                                            {
                                                ElementType eType = doc.GetElement(typeId) as ElementType;
                                                if (eType != null) {
                                                    foreach (Parameter pI in eType.Parameters) { if (pI.Id == matchedField.ParameterId) { p = pI; break; } }
                                                }
                                            }
                                        }

                                        if (p != null && !p.IsReadOnly)
                                        {
                                            string sVal = values[r, c]?.ToString();
                                            try { p.Set(sVal); updated = true; } catch { }
                                        }
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
