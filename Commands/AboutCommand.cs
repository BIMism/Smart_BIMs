using System;
using System.Diagnostics;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Smart_BIMs.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class AboutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("About Smart_BIMs", "Publisher: Asanka Dharmarathna\nWebsite: https://academyinnov.com/");
            
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = "https://academyinnov.com/",
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch(Exception e)
            {
                System.Diagnostics.Debug.WriteLine(e.Message);
            }

            return Result.Succeeded;
        }
    }
}
