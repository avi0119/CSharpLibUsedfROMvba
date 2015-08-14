using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection; 

namespace Excel13Addin_v2
{
    public partial class Ribbon2
    {
        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Excel._Application app = (Excel._Application)Globals.ThisAddIn.Application;
            Excel._Workbook wb2 = app.ActiveWorkbook;

            Type to = typeof(Excel.Application);

           //object f= to.InvokeMember("ThisWorkbook", BindingFlags.GetProperty, null, app, null);
            Excel.Workbook wb = (Excel.Workbook)app.ActiveWorkbook;
            string name = wb.Name;
           Excel.Workbook addedwb= Globals.ThisAddIn.Application.Workbooks.Add();
           addedwb.SaveAs( "newwb");
        }
    }
}
