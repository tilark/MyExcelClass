using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUsingEPPlus
{
    public class WriteToExcel : IWriteToExcel
    {
        public string ExcelContentType
        {
            get
            {
               return ExcelExportHelper.ExcelContentType;
            }
        }

        public byte[] ExportExcel<T>(List<T> data, List<string> heading, bool isShowSlNo = false, params string[] ColumnsToTake)
        {
            return ExcelExportHelper.ExportExcel(data, heading, isShowSlNo, ColumnsToTake);
        }
    }
}
