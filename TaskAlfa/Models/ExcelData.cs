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
        //public int AccountNumberExl { get; set; }
        //public string Сurrency { get; set; }
        //public double Remainder { get; set; }


        public DataModel Read(string path)
        {
            DataModel dm = new DataModel();

            //Range Rng;
            //Workbook xlWb;
            //Worksheet xlSht;
            //Application xlApp = new Application();
            //xlWb = xlApp.Workbooks.Open(path);
            //xlSht = (Worksheet)xlWb.ActiveвSheet;
            //Rng = xlSht.UsedRange;
            //Range last = xlSht.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);

            foreach (var worksheet in Workbook.Worksheets(path))
               foreach (var row in worksheet.Rows)
                {
                    var item = new List<string>();
                    foreach(var cell in row.Cells)
                    {
                        item.Add(cell.Value);
                    }
                    //exD.Add(new DataFile
                    //{
                    //    AccountNumber = long.Parse(item[0]),
                    //    Сurrency = "UAH",
                    //    Rest = int.Parse(item[2])
                    //});
                    dm.setValue(item[0], item[1], double.Parse(item[2]));


                }
            return dm;
                    

            //int lastUsedRow = last.Row;
            //Rng = xlSht.Range["C1", $"C{lastUsedRow - 1 }"];
            //sum = xlApp.WorksheetFunction.Sum(Rng);
            /*
            Returns a Sheets collection that represents all the worksheets in the specified workbook. Read-only Sheets object.

*/
            //            for (long row=1; row<= Rng.Rows.Count;row++)
            //{
            //        exD.Add(new DataFile
            //        {
            //            AccountNumber = long.Parse(Rng.Cells[$"A{row}", $"A1"]),
            //            Сurrency = Rng.Cells[$"B{row}",$"B1"],
            //            Rest = int.Parse(Rng.Cells[$"C{row}", $"C3"])
            //        });

            //}
            //xlWb.Save();
            //xlWb.Close();
            //xlApp.Quit();

            //excelReader.Read();
            //while (excelReader.Read())
            //{

            //    //ed.Add(new ExcelData
            //    //{
            //    //    AccountNumberExl = int.Parse(excelReader.GetValue(0).ToString()),
            //    //    Сurrency = excelReader.GetValue(1).ToString(),
            //    //    Remainder = double.Parse(excelReader.GetValue(2).ToString())
            //    //});

            //}


            // return exD ;
        }

    }

}

