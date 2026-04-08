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

                // 2. Create Panel
                RibbonPanel panel = application.CreateRibbonPanel(tabName, "Schedules");

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

                // Add to panel
                panel.AddItem(scheduleBtnData);
                panel.AddItem(aboutBtnData);

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
                using (Stream stream = assembly.GetManifestResourceStream("Smart_BIMs.Resources." + resourceName))
                {
                    if (stream == null) return null;
                    BitmapImage image = new BitmapImage();
                    image.BeginInit();
                    image.StreamSource = stream;
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.EndInit();
                    return image;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
