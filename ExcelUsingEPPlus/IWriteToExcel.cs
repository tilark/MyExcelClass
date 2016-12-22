using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUsingEPPlus
{
    public interface IWriteToExcel
    {
        //利用EPPlus插件完成的文件下载操作
        string ExcelContentType { get; }
        byte[] ExportExcel<T>(List<T> data, List<string> heading, bool isShowSlNo = false, params string[] ColumnsToTake);
    }
}
