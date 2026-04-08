using System;
using System.Reflection;
using Autodesk.Revit.UI;
using Smart_BIMs.Utils;

namespace Smart_BIMs
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            // 1. Create Ribbon Tab
            string tabName = "Smart BIMs";
            application.CreateRibbonTab(tabName);

            // 2. Create Panel
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Schedules");

            // 3. Create Button
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData buttonData = new PushButtonData(
                "cmdEasySchedule",
                "Easy\nSchedule",
                thisAssemblyPath,
                "Smart_BIMs.Commands.EasyScheduleCommand"
            );

            buttonData.ToolTip = "Easily create schedules by selecting categories.";
            PushButton pushButton = panel.AddItem(buttonData) as PushButton;

            // Optional: Check for updates asynchronously (do not block UI)
            GithubUpdateChecker.CheckForUpdatesAsync("YOUR_GITHUB_USERNAME", "Smart_BIMs");

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
