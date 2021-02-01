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
        public DataModel Read(string path)
        {
            DataModel dm = new DataModel();

            using (StreamReader reader = new StreamReader(path))
            {
                using (CsvReader csvReader = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    csvReader.Configuration.Delimiter = ",";

                    csvReader.Read();
                    while (csvReader.Read())
                    {
                        dm.setValue(csvReader.GetField(1), csvReader.GetField(2), csvReader.GetField<double>(0));
                    }
                }
                reader.Close();
            }

            return dm;
        }

        public void Save(string path, List<string> result)
        {
            Application app = new Application();
            Workbook wb = app.Workbooks.Open(path);
            Worksheet ws = wb.ActiveSheet;
            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);

            for (int i = 0; i < result.Count; i++)
            {
              ws.Range[$"A{i + 1}"].Value = result[i];
            }

            wb.Save();
            wb.Close();
            
        }
    }
}
