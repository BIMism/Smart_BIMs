using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Smart_BIMs.UI;

namespace Smart_BIMs.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class AboutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            AboutWindow window = new AboutWindow();
            window.ShowDialog();
            return Result.Succeeded;
        }
    }
}
