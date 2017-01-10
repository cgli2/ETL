using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelTool
{
   public class CheckProcess
    {
       public static bool Existed(String proccessName)
        {
            return (System.Diagnostics.Process.GetProcessesByName(proccessName).ToList().Count > 0);
        }
    }
}
