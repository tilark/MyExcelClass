using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace MyExcelClass
{
    public class ReadFromExcel : IReadFromExcel
    {
        PublicClass publicClass = null;
        public ReadFromExcel()
        {
            this.publicClass = new PublicClass();
        }
        public int GetRowCount(string fileName)
        {
            int rowCount = 0;
            try
            {
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    rowCount = sheetData.Elements<Row>().Count();
                }
            }
            catch (Exception ex)
            {
                string errorMess = ex.Message;
            }
            return rowCount;
        }

        public  string ReadFromExcelByCell(string fileName, string CellName)
        {
            string result = null;
            try
            {
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    Cell theCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == CellName).FirstOrDefault();
                    if(theCell != null)
                    {
                        result = publicClass.GetCellValue(workbookPart, theCell);
                    }
                }
            }
            catch (Exception ex)
            {
                //如果文件不存在，会引发该异常，不处理，直接返回NULL值
                string errorMess = ex.Message;
            }
            return result;
        }
        public Dictionary<string, string> ReadFromExcelFileByDOM(string fileName)
        {
            Dictionary<string, string> ListData = new Dictionary<string, string>();
            try
            {
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    string text;
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            //根据DataType读取数据
                            text = publicClass.GetCellValue(workbookPart, cell);
                            ListData.Add(cell.CellReference,text);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //如果文件不存在，会引发该异常，不处理，直接返回NULL值
                string errorMess = ex.Message;

            }
            return ListData;
        }
        public Dictionary<string, string> ReadFromExcelByCellRange(string fileName, string firstCellName, string lastCellName)
        {
            Dictionary<string, string> ListData = new Dictionary<string, string>();
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    // Get the row number and column name for the first and last cells in the range.
                    uint firstrowIndex = publicClass.GetRowIndex(firstCellName);
                    uint lastrowIndex = publicClass.GetRowIndex(lastCellName);
                    string firstColumn = publicClass.GetColumnName(firstCellName);
                    string lastColumn = publicClass.GetColumnName(lastCellName);
                    string text = null;
                    //加.Skip<Row>(1)可过滤第一行标题。
                    foreach (Row row in worksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value >= firstrowIndex && r.RowIndex.Value <= lastrowIndex))
                    {
                        foreach (Cell cell in row)
                        {
                            //某些Cell中并没有数据
                            string columnName = publicClass.GetColumnName(cell.CellReference.Value);
                            if (publicClass.CompareColumn(columnName, firstColumn) >= 0 && publicClass.CompareColumn(columnName, lastColumn) <= 0)
                            {
                                //在选定的范围内，可以读取数据
                                text = publicClass.GetCellValue(workbookPart, cell);
                                ListData.Add(cell.CellReference, text);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //如果文件不存在，会引发该异常，不处理，直接返回NULL值
                string errorMess = ex.Message;
            }
            return ListData;
        }

        public Dictionary<string, string> ReadFromExcelByRow(string fileName, uint rowIndex)
        {
            Dictionary<string, string> ListData = new Dictionary<string, string>();
            try
            {
                using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(fileName, false))
                {
                    WorkbookPart workbookPart = spreadsheetDoc.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    string text;
                    Row row = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
                    if (row == null)
                    {
                        return ListData;
                    }
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        //根据DataType读取数据
                        text = publicClass.GetCellValue(workbookPart, cell);
                        ListData.Add(cell.CellReference, text);
                    }
                }
            }
            catch (Exception ex)
            {

                //如果文件不存在，会引发该异常，不处理，直接返回NULL值
                string errorMess = ex.Message;

            }
            return ListData;
        }      

    }
}
