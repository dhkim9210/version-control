using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Windows.Forms;


namespace FirstProgram
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class1 : IExternalCommand
    {
        static Form1 a;
        
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData revit,
         ref string message, ElementSet elements)
        {
            TaskDialog.Show("Revit--", "프로그램이시작됩니다");
            a = new Form1();
            Application.Run(a);
            return Autodesk.Revit.UI.Result.Succeeded;
        }
    }
}
