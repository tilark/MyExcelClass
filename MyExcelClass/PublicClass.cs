using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MyExcelClass
{
    public class PublicClass
    {
        
        /// <summary>
        /// 获取单元格数据
        /// </summary>
        /// <param name="workbookPart"></param>
        /// <param name="cell">单元格名</param>
        /// <returns></returns>
        internal string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            string value = null;
            if (cell != null && cell.HasChildren)
            {
                value = cell.CellValue.Text;

                if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();
                            if (stringTable != null)
                            {
                                value = stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;
                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }
            return value;
        }
        // Given a cell name, parses the specified cell to get the row index.
        internal uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }
        // Given a cell name, parses the specified cell to get the column name.
        internal string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }
        // Given two columns, compares the columns.
        internal int CompareColumn(string column1, string column2)
        {
            if (column1.Length > column2.Length)
            {
                return 1;
            }
            else if (column1.Length < column2.Length)
            {
                return -1;
            }
            else
            {
                return string.Compare(column1, column2, true);
            }
        }

    }

}
