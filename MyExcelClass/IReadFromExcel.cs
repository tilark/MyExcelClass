using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelClass
{
    public interface IReadFromExcel
    {
        int GetRowCount(string fileName);
        Dictionary<string, string> ReadFromExcelFileByDOM(string fileName);

        Dictionary<string, string> ReadFromExcelByRow(string fileName, uint rowIndex);
        Dictionary<string, string> ReadFromExcelByCellRange(string fileName, string firstCellName, string lastCellName);
        string ReadFromExcelByCell(string fileName, string CellName);
    }
}
