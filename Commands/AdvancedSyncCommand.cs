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

        public void ExecuteSilentlySync(ExternalCommandData commandData)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ViewSchedule schedule = doc.ActiveView as ViewSchedule;
            if (schedule == null) return;

            dynamic excelApp = null;
            try { excelApp = COMHelper.GetActiveObject("Excel.Application"); } catch { return; } 
            
            if (excelApp.Workbooks.Count == 0) { System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp); return; }
            dynamic wb = excelApp.ActiveWorkbook;
            dynamic ws = wb.ActiveSheet;
            
            List<ScheduleField> fields = new List<ScheduleField>();
            for (int i = 0; i < schedule.Definition.GetFieldCount(); i++) fields.Add(schedule.Definition.GetField(i));

            SyncFromExcelLive(doc, schedule, fields, ws);
            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }

        public void ExecuteSilentlyExport(ExternalCommandData commandData)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ViewSchedule schedule = doc.ActiveView as ViewSchedule;
            if (schedule == null) return;

            List<ScheduleField> fields = new List<ScheduleField>();
            for (int i = 0; i < schedule.Definition.GetFieldCount(); i++) fields.Add(schedule.Definition.GetField(i));

            ExcelManagerWindow ui = new ExcelManagerWindow(doc, schedule);
            ui.chkSyncFonts.IsChecked = true;
            ui.chkSyncShading.IsChecked = true;
            ui.chkSyncBorders.IsChecked = true;
            ui.chkFreezeHeader.IsChecked = true;
            ui.chkStripeRows.IsChecked = false;

            dynamic excelApp = null;
            bool isNew = false;
            try { excelApp = COMHelper.GetActiveObject("Excel.Application"); }
            catch { Type t = Type.GetTypeFromProgID("Excel.Application"); excelApp = Activator.CreateInstance(t); isNew = true; }

            excelApp.Visible = true;
            dynamic wb = null;
            if (isNew || excelApp.Workbooks.Count == 0) { wb = excelApp.Workbooks.Add(System.Reflection.Missing.Value); isNew = true; }
            else { wb = excelApp.ActiveWorkbook; }

            dynamic ws = wb.ActiveSheet;
            ExportToExcelLive(doc, schedule, fields, ui, excelApp, wb, ws, isNew);
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
            dynamic titleRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, existingCols]];
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
            
            List<HashSet<string>> columnUniqueValues = new List<HashSet<string>>();
            for (int i = 0; i < fields.Count; i++) columnUniqueValues.Add(new HashSet<string>());
            Dictionary<ElementId, int> typeMasterRow = new Dictionary<ElementId, int>();

            for (int r = 0; r < collectedElements.Count; r++)
            {
                Element el = collectedElements[r];
                exportData[r, 0] = el.Id.IntegerValue.ToString();
                
                ElementId typeId = el.GetTypeId();
                bool isFirstOfType = false;
                if (typeId != ElementId.InvalidElementId && !typeMasterRow.ContainsKey(typeId)) {
                    typeMasterRow[typeId] = r + 3;
                    isFirstOfType = true;
                }

                for (int c = 0; c < fields.Count; c++)
                {
                    ScheduleField field = fields[c];
                    bool isTypeParam = false;
                    Parameter param = null;
                    foreach (Parameter pI in el.Parameters) { if (pI.Id == field.ParameterId) { param = pI; break; } }
                    
                    if (param == null)
                    {
                        if (typeId != ElementId.InvalidElementId)
                        {
                            ElementType eType = doc.GetElement(typeId) as ElementType;
                            if (eType != null) {
                                foreach (Parameter pI in eType.Parameters) { if (pI.Id == field.ParameterId) { param = pI; isTypeParam = true; break; } }
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
                        
                        if (!string.IsNullOrEmpty(val) && val.Trim() != "") columnUniqueValues[c].Add(val);

                        if (isTypeParam && !isFirstOfType) {
                            string colLetter = GetExcelColumnName(c + 2);
                            exportData[r, c + 1] = $"={colLetter}${typeMasterRow[typeId]}";
                        } else {
                            exportData[r, c + 1] = val;
                        }
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
            // Pre-populate global Enum values for Type parameters
            HashSet<ElementId> activeCategories = new HashSet<ElementId>();
            foreach (Element e in collectedElements) { if (e.Category != null) activeCategories.Add(e.Category.Id); }

            for (int c = 0; c < fields.Count; c++)
            {
                if (fields[c].ParameterId.IntegerValue == (int)BuiltInParameter.ELEM_TYPE_PARAM || fields[c].ColumnHeading == "Type")
                {
                    foreach (ElementId catId in activeCategories)
                    {
                        var types = new FilteredElementCollector(doc).OfCategoryId(catId).WhereElementIsElementType().ToElements();
                        foreach (var t in types) columnUniqueValues[c].Add(t.Name);
                    }
                }
            }
            
            if (paramNames.Count > 0)
            {
                int maxDictRows = paramNames.Count;
                foreach (var set in columnUniqueValues) if (set.Count > maxDictRows) maxDictRows = set.Count;

                object[,] dictData = new object[maxDictRows, 2 + fields.Count];
                for(int i=0; i<paramNames.Count; i++) 
                {
                    dictData[i,0] = paramNames[i];
                    dictData[i,1] = paramReadOnly[i] ? "True" : "False";
                }
                
                for (int c = 0; c < fields.Count; c++)
                {
                    List<string> listVals = columnUniqueValues[c].ToList();
                    for(int r = 0; r < listVals.Count; r++) {
                        dictData[r, 2 + c] = listVals[r];
                    }
                }

                dynamic dictRange = dictWs.Range[dictWs.Cells[1,1], dictWs.Cells[maxDictRows, 2 + fields.Count]];
                dictRange.Value2 = dictData;
            }

            // Dynamic Dropdowns per strict native data column
            int maxR = existingRows < 2 ? 100 : existingRows;
            
            try
            {
                // Add Dynamic Dropdowns natively to data columns based on Dictionary values
                for (int c = 0; c < fields.Count; c++)
                {
                    if (columnUniqueValues[c].Count > 0 && !isReadOnlyCol[c])
                    {
                        dynamic colDataRange = ws.Range[ws.Cells[3, c + 2], ws.Cells[2 + maxR, c + 2]];
                        string dictCol = GetExcelColumnName(c + 3);
                        colDataRange.Validation.Delete();
                        colDataRange.Validation.Add(Type: 3, AlertStyle: 1, Operator: 1, Formula1: $"=SmartBIM_Dictionary!${dictCol}$1:${dictCol}${columnUniqueValues[c].Count}");
                        colDataRange.Validation.ShowError = false; // Soft-warn, allow custom typing
                    }
                }
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
                dynamic overallGrid = ws.Range[ws.Cells[2, 1], ws.Cells[2 + existingRows, existingCols]];
                overallGrid.Font.Name = "Arial"; // Safe default corresponding to basic Revit schedule font
            }

            // Finalize Title Merge safely avoiding Column lock intersections
            try { titleRange.Merge(); } catch { }

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
                try { titleRange.Locked = true; } catch { }
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

                    HashSet<string> appliedTypeParams = new HashSet<string>();
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
                                        bool isTypeP = false;
                                        Parameter p = null;
                                        foreach (Parameter pI in el.Parameters) { if (pI.Id == matchedField.ParameterId) { p = pI; break; } }

                                        if (p == null)
                                        {
                                            ElementId typeId = el.GetTypeId();
                                            if (typeId != ElementId.InvalidElementId)
                                            {
                                                ElementType eType = doc.GetElement(typeId) as ElementType;
                                                if (eType != null) {
                                                    foreach (Parameter pI in eType.Parameters) { if (pI.Id == matchedField.ParameterId) { p = pI; isTypeP = true; break; } }
                                                }
                                            }
                                        }

                                        if (p != null && !p.IsReadOnly)
                                        {
                                            string sVal = values[r, c]?.ToString() ?? "";
                                            string curVal = p.AsValueString();
                                            if (string.IsNullOrEmpty(curVal)) curVal = p.AsString();
                                            if (string.IsNullOrEmpty(curVal)) {
                                                if (p.StorageType == StorageType.Double) curVal = p.AsDouble().ToString("0.##");
                                                else if (p.StorageType == StorageType.Integer) curVal = p.AsInteger().ToString();
                                            }
                                            
                                            if ((curVal ?? "") != sVal)
                                            {
                                                if (isTypeP)
                                                {
                                                    string sig = $"{el.GetTypeId().IntegerValue}_{p.Id.IntegerValue}";
                                                    if (appliedTypeParams.Contains(sig)) continue;
                                                    appliedTypeParams.Add(sig);
                                                }
                                                
                                                if (p.StorageType == StorageType.ElementId)
                                                {
                                                    if (p.Id.IntegerValue == (int)BuiltInParameter.ELEM_TYPE_PARAM || matchedField.ColumnHeading == "Type")
                                                    {
                                                        var types = new FilteredElementCollector(doc).OfCategoryId(el.Category.Id).WhereElementIsElementType().ToElements();
                                                        Element match = types.FirstOrDefault(t => t.Name == sVal);
                                                        if (match != null) { try { el.ChangeTypeId(match.Id); updated = true; } catch { } }
                                                    }
                                                }
                                                else
                                                {
                                                    bool setSuccess = false;
                                                    try { setSuccess = p.SetValueString(sVal); } catch { }
                                                    if (!setSuccess) { try { p.Set(sVal); } catch { } }
                                                    updated = true;
                                                }
                                            }
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
        private string GetExcelColumnName(int columnIndex)
        {
            string columnName = "";
            while (columnIndex > 0)
            {
                int modulo = (columnIndex - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                columnIndex = (columnIndex - modulo) / 26;
            }
            return columnName;
        }
    }
}
