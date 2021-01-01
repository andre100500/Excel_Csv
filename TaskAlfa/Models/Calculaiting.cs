using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskAlfa.Models
{
    public class Calculaiting
    {
        CsvData csd;
        public override bool Equals(object obj)
        {
            if (!(obj is ExcelData))
            {
                throw new Exception();
            }
            var data = obj as ExcelData;
            //var result = csd.Rest - data.Remainder;

            return true;
        }
    }
}
