using System;
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace VBAAccessedCodeNoLibrary
{

    public class ExcelDNA
    {
        
        [ExcelFunction(Description = "Multiplies two numbers", Category = "Useful functions")]
        public static double MultiplyThem(double x, double y)
        {
            return x * y;
        }

        [ExcelCommand(MenuName = "AddinHost", MenuText = "Display Message")]
        public static void PoppupMessage()
        {
            string message = "You did not enter a server name. Cancel this operation?";
            string caption = "Error Detected in Input";
            MessageBoxButtons buttons = MessageBoxButtons.YesNo;
            DialogResult result;

            // Displays the MessageBox.

            result = MessageBox.Show(message, caption, buttons);
        }
    }
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]

    public class CSharpLib
    {
        public double add(double x, double y)
        {
            return x + y;
        }
        public AreaCalculator AC()
        {
            return new AreaCalculator();
        }
    }


    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]

    public class CSharpLib2
    {
        public double add2(double x, double y)
        {
            return x + y;
        }
    }

    [ComVisible(false)]
    class ExcelAddin : IExcelAddIn
    {
        public void AutoOpen()
        {
            ComServer.DllRegisterServer();
        }
        public void AutoClose()
        {
            ComServer.DllUnregisterServer();
        }
    }

    [ComVisible(true)]
    public class Class1 : ExcelRibbon
    {

        public void ReformatSelection(IRibbonControl control)
        {
            MessageBox.Show("Hello");
        }
    }


}
