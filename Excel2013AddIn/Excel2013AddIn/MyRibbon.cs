using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
//using Excel = Microsoft.Office.Interop.Excel;

namespace Excel2013AddIn
{
    public partial class MyRibbon
    {
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //Excel.Application app = Globals.ThisAddIn.Application;
            //string name = app.ThisWorkbook.Name;
            //Globals.ThisAddIn.Application.Workbooks.Add();
        }
    }
}
