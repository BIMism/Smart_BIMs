using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Excel = Microsoft.Office.Interop.Excel;

namespace Smart_BIMs.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ExportToExcelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ViewSchedule schedule = doc.ActiveView as ViewSchedule;
            if (schedule == null)
            {
                TaskDialog.Show("Error", "Please make a Schedule view active before exporting.");
                return Result.Failed;
            }

            try
            {
                ScheduleDefinition def = schedule.Definition;
                int fieldCount = def.GetFieldCount();
                List<ScheduleField> fields = new List<ScheduleField>();
                for (int i = 0; i < fieldCount; i++)
                {
                    fields.Add(def.GetField(i));
                }

                var collectedElements = new FilteredElementCollector(doc, schedule.Id).ToElements();

                Excel.Application excelApp = null;
                try
                {
                    excelApp = (Excel.Application)COMHelper.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Excel.Application();
                }

                excelApp.Visible = true;
                Excel.Workbook wb = excelApp.Workbooks.Add(Type.Missing);
                Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
                ws.Name = "ScheduleLIVE";

                // Add headers
                ws.Cells[1, 1] = "ElementId";
                for (int i = 0; i < fields.Count; i++)
                {
                    ws.Cells[1, i + 2] = fields[i].GetName();
                }

                Excel.Range headerRange = ws.Range[ws.Cells[1, 1], ws.Cells[1, fields.Count + 1]];
                headerRange.Font.Bold = true;
                headerRange.Interior.ColorIndex = 37; // Standard Light Blue COM Color

                // Add Data using 2D array for extreme speed
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
                            foreach(Parameter param in el.Parameters)
                            {
                                if (param.Id == fields[c].ParameterId) { p = param; break; }
                            }
                            data[r, c + 1] = p != null ? (p.AsValueString() ?? p.AsString() ?? "") : "";
                        }
                        r++;
                    }

                    Excel.Range dataRange = ws.Range[ws.Cells[2, 1], ws.Cells[rows + 1, cols]];
                    dataRange.Value2 = data;
                }

                ws.Columns.AutoFit();
                
                // Release COM gracefully
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
