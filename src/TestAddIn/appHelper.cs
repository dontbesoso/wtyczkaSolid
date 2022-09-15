using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestAddIn
{
    class appHelper
    {
        public static string getSheetSize(SolidEdgeDraft.SheetSetup objSheetSetup)
        {
            double sheetWidth = objSheetSetup.SheetWidth;
            double sheetHeight = objSheetSetup.SheetHeight;

            if ((sheetWidth == 0.210) && (sheetHeight == 0.297))
                return "A4";
            if ((sheetWidth == 0.297) && (sheetHeight == 0.210))
                return "A4R";
            if ((sheetWidth == 0.297) && (sheetHeight == 0.420))
                return "A3";
            if ((sheetWidth == 0.420) && (sheetHeight == 0.297))
                return "A3R";

            return null;
        }
    }
}
