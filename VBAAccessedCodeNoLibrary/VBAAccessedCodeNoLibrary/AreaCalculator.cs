using System;
using ExcelDna.Integration;
using ExcelDna.ComInterop;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;

namespace VBAAccessedCodeNoLibrary
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [ProgId("MyAreaCaculator")]
    public class AreaCalculator
    {
        double prlength;
        double prwidth;
        public double Length { 
            get
            {
                return prlength;
            }
            set
            {
                prlength=value;
            } 
        }
        public double Width
        {
            get
            {
                return prwidth;
            }
            set
            {
                prwidth = value;
            }
        }
        public double Calculate()
        {
            return Length * Width;
        } 
    }
}
