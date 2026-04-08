using System;
using System.Reflection;
using Autodesk.Revit.UI;

namespace Smart_BIMs.Utils
{
    public static class GithubUpdateChecker
    {
        public static void CheckForUpdatesAsync(string user, string repo)
        {
            try
            {
                // In production, use HttpClient to read from 
                // https://api.github.com/repos/{user}/{repo}/releases/latest
                
                string currentVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
                
                // Placeholder test logic: assume 1.0.0.0 is the latest for now.
                string latestVersion = "1.0.0.0"; 

                if (currentVersion != latestVersion)
                {
                    TaskDialog.Show("Smart_BIMs Update", 
                        $"A new version ({latestVersion}) is available on GitHub! Please download the latest installer from releases.");
                }
            }
            catch
            {
                // Quietly fail if network is down
            }
        }
    }
}
