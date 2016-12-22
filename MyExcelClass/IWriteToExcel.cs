using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelClass
{
    public interface IWriteToExcel
    {
        //bool TransferSoureFileTokDesFile(string sourceFile, string desFile);
        bool TransferTemplateToWorkBook(string templateFile, string desFile);
        void WriteToExcelByRow<T>(string fileName, string cellName, List<T> valueList);
        void WriteToExcelByCell<T>(string fileName, string cellName, T appendValue);

        void WriteToExcelByCellReplace<T>(string fileName, string cellName, T appendValue);

        //利用EPPlus插件完成的文件下载操作
        string ExcelContentType { get;  }
        /// <summary>
        /// Exports the excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">The data.</param>
        /// <param name="heading">The heading.第一个元素为标题及工作簿名称</param>
        /// <param name="isShowSlNo">是否显示每行的序号.</param>
        /// <param name="ColumnsToTake">包含的行名称.</param>
        /// <returns>System.Byte[].</returns>
        byte[] ExportExcel<T>(List<T> data, List<string> heading, bool isShowSlNo = false, params string[] ColumnsToTake);
    }
}
