using Excel;
using ExcelDataReader;
//using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;

namespace TaskAlfa.Models
{
    public class ExcelData : IReader
    {
        public DataModel Read(string path)
        {
            DataModel dm = new DataModel();

            foreach (var worksheet in Workbook.Worksheets(path))
               foreach (var row in worksheet.Rows)
                {
                    var item = new List<string>();
                    foreach(var cell in row.Cells)
                    {
                        item.Add(cell.Value);
                    }
                    dm.setValue(item[0], item[1], double.Parse(item[2]));
                }
            return dm;
        }

    }

}

