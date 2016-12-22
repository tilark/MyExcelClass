using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;

namespace MyExcelClass
{
    public class ExcelExportHelper
    {
        public static string ExcelContentType
        {
            get
            {
                return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            }
        }

        /// <summary>
        /// Lists 转成 data table.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">The data.</param>
        /// <returns>DataTable.</returns>
        /// <remarks></remarks>
        public static DataTable ListToDataTable<T>(List<T> data)
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));

            DataTable dataTable = new DataTable();
            for (int i = 0; i < properties.Count; i++)
            {
                PropertyDescriptor property = properties[i];

                dataTable.Columns.Add(property.Name, Nullable.GetUnderlyingType(property.PropertyType) ?? property.PropertyType);
            }

            object[] values = new object[properties.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = properties[i].GetValue(item);
                }

                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        /// <summary>
        /// 导出Excel.
        /// </summary>
        /// <param name="dataTable">数据源.</param>
        /// <param name="heading">第一行为工作簿Worksheet名称及标题（必填），第二行为其他备注消息（可选）.</param>
        /// <param name="showSrNo">是否显示行编号</c> [show sr no].</param>
        /// <param name="columnsToTake">要导出的列，如果为空，导出所有列.</param>
        /// <returns>System.Byte[].</returns>
        public static byte[] ExportExcel(DataTable dataTable, List<string> heading, bool showSrNo = false, params string[] columnsToTake)
        {
            byte[] results = null;
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheetName = heading.Count > 0 ? heading[0] : String.Empty;
                ExcelWorksheet workSheet = package.Workbook.Worksheets.Add(string.Format("{0} Data", worksheetName));

                //从哪一行开始填数据
                int startRowFrom = heading.Count + 1;

                //是否显示行编号
                if (showSrNo)
                {
                    DataColumn dataColumn = dataTable.Columns.Add("序号", typeof(int));
                    dataColumn.SetOrdinal(0);
                    int index = 1;
                    foreach (DataRow item in dataTable.Rows)
                    {
                        item[0] = index;
                        index++;
                    }
                }

                //Add Content Into the Excel File
                workSheet.Cells["A" + startRowFrom].LoadFromDataTable(dataTable, true);
                int columnIndex = 1;
                foreach (DataColumn item in dataTable.Columns)
                {
                    ExcelRange columnCells = workSheet.Cells[workSheet.Dimension.Start.Row, columnIndex, workSheet.Dimension.End.Row, columnIndex];
                    int maxLength = columnCells.Max(cell => cell.Value.ToString().Count());

                    if (maxLength < 150)
                    {
                        workSheet.Column(columnIndex).AutoFit();
                    }

                    columnIndex++;
                }

                //format header - bold, yellow on black
                using (ExcelRange r = workSheet.Cells[startRowFrom, 1, startRowFrom, dataTable.Columns.Count])
                {
                    r.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    r.Style.Font.Bold = true;
                    r.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    r.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#1fb5ad"));
                }

                //format cells - add borders
                using (ExcelRange r = workSheet.Cells[startRowFrom + 1, 1, startRowFrom + dataTable.Rows.Count, dataTable.Columns.Count])
                {
                    r.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    r.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    r.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    r.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    r.Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                    r.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                    r.Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                    r.Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);

                }

                //removed ignored columns
                if (columnsToTake.Count() > 0)
                {
                    for (int i = dataTable.Columns.Count - 1; i >= 0; i--)
                    {
                        if (i == 0 && showSrNo)
                        {
                            continue;
                        }

                        if (!columnsToTake.Contains(dataTable.Columns[i].ColumnName))
                        {
                            workSheet.DeleteColumn(i + 1);
                        }
                    }
                }
                //创建表标题
                if (heading.Count > 0)
                {
                    var titleColumn = "A";

                    for (int titleRow = 0; titleRow < heading.Count; titleRow++)
                    {
                        var titleCell = titleColumn + (titleRow + 1).ToString();
                        workSheet.Cells[titleCell].Value = heading[titleRow];

                    }
                    workSheet.Cells["A1"].Style.Font.Size = 20;
                    //此处需将标题能够跨行排列最好

                    //新插入一行一列
                    workSheet.InsertColumn(1, 1);
                    workSheet.InsertRow(1, 1);
                    workSheet.Column(1).Width = 2;
                    workSheet.Row(1).Height = 10;
                }

                results = package.GetAsByteArray();
            }

            return results;
        }

        /// <summary>
        /// 导出Excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">数据集合.</param>
        /// <param name="heading">The heading.</param>
        /// <param name="isShowSlNo">if set to <c>true</c> [is show sl no].</param>
        /// <param name="ColumnsToTake">The columns to take.</param>
        /// <returns>System.Byte[].</returns>
        public static byte[] ExportExcel<T>(List<T> data, List<string> heading, bool isShowSlNo = false, params string[] ColumnsToTake)
        {
            return ExportExcel(ListToDataTable<T>(data), heading, isShowSlNo, ColumnsToTake);
        }
    }
}
