using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskAlfa.Models
{
    public  class ReaderExcelData
    {
        public  class Factory
        {

            public static IReader CreatReader(string path)
            {
                Dictionary<string, IReader> keyValues = new Dictionary<string, IReader>()
                {
                    {".xls",new ExcelData() },
                    {".csv", new CsvData() },
                    {".xlsx", new ExcelData() }
                };
                var extention = Path.GetExtension(path);

                return keyValues[extention];
            }
        }
    }
}
