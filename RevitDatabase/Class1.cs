using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;

namespace RevitDatabase
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]

    public class Class1 : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //GET APPLICATION AND DOCUMENT OBJECTS
            UIApplication uiApp = commandData.Application;

            MainForm myMainForm = new MainForm(uiApp); //CREATES A NEW MAINFORM AND PASSES THE REVIT APP TO ACCESS ELEMENTS

            myMainForm.ShowDialog();

            return Result.Succeeded;

        }
    }
}
