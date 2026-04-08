using System;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using Smart_BIMs.Utils;

namespace Smart_BIMs
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                // 1. Create Ribbon Tab
                string tabName = "Smart BIMs";
                application.CreateRibbonTab(tabName);

                // 2. Create Panels
                RibbonPanel panelSchedules = application.CreateRibbonPanel(tabName, "Schedules");
                RibbonPanel panelAbout = application.CreateRibbonPanel(tabName, "About");

                // 3. Setup commands
                string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
                
                PushButtonData scheduleBtnData = new PushButtonData(
                    "cmdEasySchedule",
                    "Easy\nSchedule",
                    thisAssemblyPath,
                    "Smart_BIMs.Commands.EasyScheduleCommand"
                );
                scheduleBtnData.ToolTip = "Easily create schedules by selecting categories.";
                scheduleBtnData.LargeImage = GetImageFromResource("schedule_icon.png");

                PushButtonData aboutBtnData = new PushButtonData(
                    "cmdAbout",
                    "About",
                    thisAssemblyPath,
                    "Smart_BIMs.Commands.AboutCommand"
                );
                aboutBtnData.ToolTip = "Learn more about Smart BIMs at academyinnov.com";
                aboutBtnData.LargeImage = GetImageFromResource("about_icon.png");

                // Export to Excel Button
                PushButtonData exportBtnData = new PushButtonData(
                    "cmdExportExcel",
                    "Export to\nExcel",
                    thisAssemblyPath,
                    "Smart_BIMs.Commands.ExportToExcelCommand"
                );
                exportBtnData.ToolTip = "Export active schedule to Excel for bulk editing.";
                exportBtnData.LargeImage = GetImageFromResource("export_icon.png");

                // Sync from Excel Button
                PushButtonData importBtnData = new PushButtonData(
                    "cmdImportExcel",
                    "Sync from\nExcel",
                    thisAssemblyPath,
                    "Smart_BIMs.Commands.SyncFromExcelCommand"
                );
                importBtnData.ToolTip = "Read data from an Excel file to sync back to Revit.";
                importBtnData.LargeImage = GetImageFromResource("import_icon.png");

                // Add to panels
                panelSchedules.AddItem(scheduleBtnData);
                panelSchedules.AddItem(exportBtnData);
                panelSchedules.AddItem(importBtnData);
                panelAbout.AddItem(aboutBtnData);

                // Optional: Check for updates silently
                GithubUpdateChecker.CheckForUpdatesAsync("BIMism", "Smart_BIMs");

                return Result.Succeeded;
            }
            catch(Exception ex)
            {
                TaskDialog.Show("Error", ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private BitmapImage GetImageFromResource(string resourceName)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string[] names = assembly.GetManifestResourceNames();
                string match = null;
                foreach (string name in names)
                {
                    if (name.EndsWith(resourceName, StringComparison.OrdinalIgnoreCase))
                    {
                        match = name;
                        break;
                    }
                }
                
                if (match != null)
                {
                    using (Stream stream = assembly.GetManifestResourceStream(match))
                    {
                        if (stream != null)
                        {
                            BitmapImage image = new BitmapImage();
                            image.BeginInit();
                            image.StreamSource = stream;
                            image.CacheOption = BitmapCacheOption.OnLoad;
                            image.EndInit();
                            return image;
                        }
                    }
                }
                return null;
            }
            catch
            {
                return null;
            }
        }
    }
}
