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

                // Gather elements that appear in this schedule's rules
                var collectedElements = new FilteredElementCollector(doc, schedule.Id).ToElements();

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Excel Files (*.xlsx)|*.xlsx";
                sfd.FileName = schedule.Name + " - SyncData";
                
                if (sfd.ShowDialog() == true)
                {
                    using (XLWorkbook workbook = new XLWorkbook())
                    {
                        IXLWorksheet ws = workbook.Worksheets.Add("ScheduleData");

                        // Add headers
                        ws.Cell(1, 1).Value = "ElementId";
                        for (int i = 0; i < fields.Count; i++)
                        {
                            ws.Cell(1, i + 2).Value = fields[i].GetName();
                        }
                        
                        // Header Styling
                        var headerRow = ws.Row(1);
                        headerRow.Style.Font.Bold = true;
                        headerRow.Style.Fill.BackgroundColor = XLColor.AirForceBlue;
                        headerRow.Style.Font.FontColor = XLColor.White;

                        // Add Data
                        int row = 2;
                        foreach (Element el in collectedElements)
                        {
                            ws.Cell(row, 1).Value = (double)el.Id.Value;
                            for (int i = 0; i < fields.Count; i++)
                            {
                                Parameter p = null;
                                foreach(Parameter param in el.Parameters)
                                {
                                    if (param.Id == fields[i].ParameterId) { p = param; break; }
                                }
                                
                                if (p != null)
                                {
                                    ws.Cell(row, i + 2).Value = p.AsValueString() ?? p.AsString() ?? "";
                                }
                            }
                            row++;
                        }

                        // Auto-fit
                        ws.Columns().AdjustToContents();
                        workbook.SaveAs(sfd.FileName);
                    }
                    TaskDialog.Show("Success", "Schedule successfully exported to Excel!\nYou can now edit the cells and Sync them back using the Import tool.");
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
