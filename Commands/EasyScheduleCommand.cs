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
    public class EasyScheduleCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;

            try
            {
                // 1. Gather categorizes
                List<Category> scheduleableCategories = new List<Category>();
                Categories cats = doc.Settings.Categories;
                
                foreach (Category c in cats)
                {
                    if (c.CategoryType == CategoryType.Model || c.CategoryType == CategoryType.Internal)
                    {
                        if (c.HasMaterialQuantities || c.AllowsBoundParameters)
                        {
                            if (!string.IsNullOrEmpty(c.Name) && !c.Name.ToLower().Contains("dwg"))
                            {
                                scheduleableCategories.Add(c);
                            }
                        }
                    }
                }

                scheduleableCategories = scheduleableCategories.OrderBy(c => c.Name).ToList();

                // 2. Show UI WPF Window
                EasyScheduleWindow window = new EasyScheduleWindow(scheduleableCategories);
                window.ShowDialog();

                if (window.DialogResult == true)
                {
                    var selectedCategories = window.GetSelectedCategories();
                    if (selectedCategories.Count > 0)
                    {
                        int createdCount = 0;
                        using (Transaction trans = new Transaction(doc, "Create Easy Schedules"))
                        {
                            trans.Start();
                            foreach (Category cat in selectedCategories)
                            {
                                try
                                {
                                    ViewSchedule schedule = ViewSchedule.CreateSchedule(doc, cat.Id);
                                    schedule.Name = cat.Name + " - Schedule";
                                    createdCount++;
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine("Failed to create schedule for " + cat.Name + ": " + ex.Message);
                                }
                            }
                            trans.Commit();
                        }
                        
                        TaskDialog.Show("Success", $"Successfully created {createdCount} schedule(s).");
                    }
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
