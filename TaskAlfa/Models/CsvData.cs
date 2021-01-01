using CsvHelper;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskAlfa.Models
{
    public class CsvData : IReader , ISave
    {
        //public int Rest { get; set; }
        //public long AccountNumberCSV { get; set; }
        //public string СurrencyCSV { get; set; }

        public DataModel Read(string path)
        {
            // List<DataFile> cd = new List<DataFile>();
            //int sum = 0;
            DataModel dm = new DataModel();

            using (StreamReader reader = new StreamReader(path))
            {
                using (CsvReader csvReader = new CsvReader(reader, CultureInfo.InvariantCulture))
                {

                    csvReader.Configuration.Delimiter = ",";
                    //sum += csvReader.GetField<int>(0);

                    csvReader.Read();
                    while (csvReader.Read())
                    {
                        //sum += csvReader.GetField<int>(0);
                        //cd.Add(new DataFile
                        //{
                        //    Rest = int.Parse(csvReader.GetField<int>(0).ToString()),
                        //    AccountNumber = csvReader.GetField<long>(1),
                        //    Сurrency = csvReader.GetField(2)
                        //});
                        dm.setValue(csvReader.GetField(1), csvReader.GetField(2), csvReader.GetField<double>(0));
                    }
                }
                reader.Close();
            }
           
            // return cd;
            return dm;
        }

        public void Save(string path, List<string> result)
        {
            Application app = new Application();
            Workbook wb = app.Workbooks.Open(path);
            Worksheet ws = wb.ActiveSheet;
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);

            int lastUsedRow = last.Row;
            for(int i = 0; i < result.Count; i++)
            {
                ws.Range[$"A{i + 1}"].Value = result[i];
            }
            //for (int row=1; row<= lastUsedRow; row++ )
            //{
            //        ws.Range[$"A{row}"].Value = Rest;
            //        ws.Range[$"B{row}"].Value = AccountNumberCSV;
            //        ws.Range[$"C{row}"].Value = СurrencyCSV;
            //}

            wb.Save();
            wb.Close();
            
        }
    }
}
