using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace MyExcelClass
{
    public class WriteToExcel : IWriteToExcel
    {
        PublicClass publicClass = null;
       

        public WriteToExcel()
        {
            this.publicClass = new PublicClass();
        }
        public string ExcelContentType
        {
            get
            {
               return ExcelExportHelper.ExcelContentType;
            }

        }
        /// <summary>
        /// Exports the excel.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data">The data.</param>
        /// <param name="heading">The heading.</param>
        /// <param name="isShowSlNo">if set to <c>true</c> [is show sl no].</param>
        /// <param name="ColumnsToTake">The columns to take.</param>
        /// <returns>System.Byte[].</returns>
        /// <exception cref="NotImplementedException"></exception>
       public  byte[] ExportExcel<T>(List<T> data, List<string> heading, bool isShowSlNo = false, params string[] ColumnsToTake)
        {
            return ExcelExportHelper.ExportExcel(data, heading, isShowSlNo, ColumnsToTake);
        }
        /// <summary>
        /// 将Excel模版文件xltx转成xlsx.
        /// </summary>
        /// <param name="template">The template.</param>
        /// <param name="desFile">The DES file.</param>
        /// <returns><c>true</c> if XXXX, <c>false</c> otherwise.</returns>
        public virtual bool TransferTemplateToWorkBook(string templateFile, string desFile)
        {
            bool result = true;
            try
            {
                byte[] byteArray = File.ReadAllBytes(templateFile);
                using (MemoryStream stream = new MemoryStream())
                {
                    stream.Write(byteArray, 0, (int)byteArray.Length);
                    using (SpreadsheetDocument spreadsheetDoc = SpreadsheetDocument.Open(stream, true))
                    {
                        // Change from template type to workbook type
                        spreadsheetDoc.ChangeDocumentType(SpreadsheetDocumentType.Workbook);
                    }
                    File.WriteAllBytes(desFile, stream.ToArray());
                }
            }
            catch (Exception)
            {
                result = false;
                throw new FileNotFoundException("TransferTemplateToWorkBook ：创建新文件失败！");
            }
            
            return result;
        }
        /// <summary>
        /// 从模版文件读取内容，转存到目标文件后再操作.
        /// </summary>
        /// <param name="sourceFile">The source file.</param>
        /// <param name="desFile">The DES file.</param>
        /// <exception cref="System.NotImplementedException"></exception>
        //public virtual bool TransferSoureFileToDesFile(string sourceFile, string desFile)
        //{
        //    bool result = true;
        //    FileInfo source = new FileInfo(sourceFile);
        //    FileInfo des = new FileInfo(desFile);
        //    try
        //    {
        //        // open the file and clean up handles.
        //        FileStream fs = source.Open(FileMode.Open);
        //        try
        //        {
        //        }
        //        finally
        //        {
        //            fs.Dispose();
        //        }
        //        //Ensure that the target does not exist.
        //        des.Delete();
        //        //Copy the file.
        //        source.CopyTo(desFile);
        //    }
        //    catch (Exception ex)
        //    {
        //        //如果文件不存在，会引发该异常，不处理，直接返回NULL值
        //        string errorMess = ex.Message;
        //        result = false;
        //    }
        //    return result;
        //}
        /// <summary>
        /// 将数据集合按行写入到指定的EXCEL文件中.
        /// </summary>
        /// <param name="fileName">Excel文件名.</param>
        /// <param name="cellName">指定单元格开始.</param>
        /// <param name="valueList">数据集合.</param>
        public virtual void WriteToExcelByRow<T>(string fileName, string cellName, List<T> valueList)
        {
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
                {
                    WorksheetPart worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.First();
                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                    var shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault(); ;

                    //将数据列写入到Excel的rowIndex中
                    WriteToExcelColumnWithString(worksheetPart, shareStringPart, cellName, valueList);
                }
            }
            catch (Exception)
            {

                throw new FileNotFoundException("WriteToExcelByRow : 无法打到目标文件！");
            }
            
        }
        /// <summary>
        /// 修改指定单元格的内容，若单元格内有内容，将新值附加在原内容后面.
        /// </summary>
        /// <param name="fileName">Excel文件名.</param>
        /// <param name="CellName">指定单元格.</param>
        /// <param name="valueData">新值.</param>
        /// <exception cref="System.NotImplementedException"></exception>
        /// <remarks>主要用于修改标题内容，如果指定单元格内无数据，新建单元格，再写入值；若有数据，将值附加在原内容中</remarks>
        public virtual void WriteToExcelByCell<T>(string fileName, string cellName, T appendValue)
        {
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
                {
                    //根据cellName找到单元格是否存在
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    Cell theCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellName).FirstOrDefault();
                    var newCellValue = appendValue.ToString();
                    if (theCell != null)
                    {
                        //如果存在，将新值附加其后
                        //先删除Cell中的原值，再添加新值。                    
                        var oldCellValue = publicClass.GetCellValue(workbookPart, theCell);
                        newCellValue = oldCellValue + appendValue.ToString();
                        try
                        {
                            UpdateCellValue(spreadsheetDocument, worksheetPart, theCell, oldCellValue, newCellValue);
                        }
                        catch (ArgumentException)
                        {


                        }
                    }
                    else
                    {
                        //如果不存在，先创建一个Cell
                        var rowIndex = publicClass.GetRowIndex(cellName);
                        Row row = worksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
                        if (row == null)
                        {
                            row = new Row() { RowIndex = rowIndex };
                            sheetData.Append(row);
                        }
                        theCell = new Cell() { CellReference = cellName };
                        row.AppendChild(theCell);
                        worksheetPart.Worksheet.Save();
                        //将新值写入单元格           
                        CreateCellValue(spreadsheetDocument, worksheetPart, theCell, newCellValue);
                    }
                }
            }
            catch (Exception)
            {

                throw new FileNotFoundException("WriteToExcelByCell : 无法打到目标文件");
            }
            
        }
        /// <summary>
        /// 修改指定单元格的内容，若单元格内有内容，删除原有内容，再将新值放入单元格.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="cellName">Name of the cell.</param>
        /// <param name="appendValue">The append value.</param>
        /// <exception cref="NotImplementedException"></exception>
        public void WriteToExcelByCellReplace<T>(string fileName, string cellName, T appendValue)
        {
            try
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, true))
                {
                    //根据cellName找到单元格是否存在
                    WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    Cell theCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellName).FirstOrDefault();
                    var newCellValue = appendValue.ToString();
                    if (theCell != null)
                    {
                        //如果存在
                        //先删除Cell中的原值，再添加新值。                    
                        var oldCellValue = publicClass.GetCellValue(workbookPart, theCell);
                        try
                        {
                            UpdateCellValue(spreadsheetDocument, worksheetPart, theCell, oldCellValue, newCellValue);
                        }
                        catch (ArgumentException)
                        {


                        }
                    }
                    else
                    {
                        //如果不存在，先创建一个Cell
                        var rowIndex = publicClass.GetRowIndex(cellName);
                        Row row = worksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
                        if (row == null)
                        {
                            row = new Row() { RowIndex = rowIndex };
                            sheetData.Append(row);
                        }
                        theCell = new Cell() { CellReference = cellName };
                        row.AppendChild(theCell);
                        worksheetPart.Worksheet.Save();
                        //将新值写入单元格           
                        CreateCellValue(spreadsheetDocument, worksheetPart, theCell, newCellValue);
                    }
                }
            }
            catch (Exception)
            {

                throw new FileNotFoundException("WriteToExcelByCell : 无法打到目标文件");
            }
        }

        #region 内部操作
        /// <summary>
        /// 将数值和字符串类型按行方式写入到Excel中.如果整个列表项是string，则数值也会当作字符串处理。
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="ListText">The list text.</param>
        private void WriteToExcelColumnWithString<T>(WorksheetPart worksheetPart, SharedStringTablePart shareStringPart, string cellName, List<T> ListText)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            //先判断cellName是否有值，如果有值，直接覆盖
            
            uint rowIndex = publicClass.GetRowIndex(cellName);
            string beginColumnName = publicClass.GetColumnName(cellName);
            char beginColumn = beginColumnName[0];
            //检查rowIndex是否存在
            Row row = worksheetPart.Worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
            if(row == null)
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }            
            char[] columnIndex = new char[] { ' ', ' ', beginColumn };
            //columnIndex[2] = beginColumn;
            string cellReference = null;
            foreach (var textItem in ListText)
            {
                //将数据写入该列名中
                //判断该数据是不是Text，如果是数值，可直接加入
                string columnName = String.Empty;
                for (int i = 0; i < 3; i++)
                {
                    if(columnIndex[i] != ' ')
                    {
                        columnName += columnIndex[i].ToString();
                    }                    
                }
                //已经是行数据了，可将值插入到Row中
                //cellReference = columnName + columnIndex.ToString() + rowIndex;
                cellReference = columnName + rowIndex;
                Cell newCell = worksheetPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == cellReference).FirstOrDefault();
                if(newCell == null)
                {
                    newCell = new Cell() { CellReference = cellReference };
                    row.AppendChild(newCell);
                    worksheet.Save();
                }
                
                //需对传入的Text进行判断，如果是数值型的直接填入，如果是字符串，再填入到SharedStringItem
                var value = textItem.ToString();
                Decimal itemValue;
                //对数值进行Decimal判断，如果解析成功，则直接存入，未成功，则为string类型
                if (!Decimal.TryParse(value, out itemValue))
                {
                    var index = InsertSharedStringItem(value, shareStringPart);
                    newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                    newCell.CellValue = new CellValue(index.ToString());

                }
                else
                {
                    newCell.CellValue = new CellValue(itemValue.ToString());

                }
                worksheetPart.Worksheet.Save();
                columnIndex[2]++;
                if (columnIndex[2] > 'Z')
                {
                    //A 之后为AA ,AZ然后是BA,ZZ后面是AAA
                    columnIndex[2] = beginColumn;
                    columnIndex[1] = (columnIndex[1] == ' ') ? beginColumn : columnIndex[1]++;
                    if (columnIndex[1] > 'Z')
                    {
                        columnIndex[1] = beginColumn;

                        columnIndex[0] = (columnIndex[0] == ' ') ? beginColumn : columnIndex[0]++;

                        if (columnIndex[0] > 'Z')
                        {
                            //如果出现这种情况，说明文本过大，从头开始写
                            columnIndex[0] = ' ';
                            columnIndex[1] = ' ';
                            columnIndex[2] = beginColumn;
                        }
                    }
                }
            }
            rowIndex++;
        }

        private void DeleteCellValue(SpreadsheetDocument spreadsheetDocument, Cell cell, string deleteValue)
        {
            //先判断该Cell中的内容是否为文本
            Decimal itemValue;
            if (!Decimal.TryParse(deleteValue, out itemValue))
            {
                //是的话在共享文本表删除共享文本，
                int sharedStringId;
                if (int.TryParse(cell.CellValue.Text, out sharedStringId))
                {

                    RemoveSharedStringItem(sharedStringId, spreadsheetDocument);
                }
            }

        }
        private void UpdateCellValue<T>(SpreadsheetDocument spreadsheetDocument, WorksheetPart worksheetPart, Cell cell, T oldValue, T newValue)
        {
            //判断oldValue是否为文本，如果是，保存旧单元格的sharedStringID，创建新值后，再删除共享文本表对应内容
            Decimal itemValue;
            if (!Decimal.TryParse(oldValue.ToString(), out itemValue))
            {
                //是的话在共享文本表删除共享文本，
                int sharedStringId;
                if (int.TryParse(cell.CellValue.Text, out sharedStringId))
                {
                    CreateCellValue(spreadsheetDocument, worksheetPart, cell, newValue);
                    RemoveSharedStringItem(sharedStringId, spreadsheetDocument);
                }
                else
                {
                    throw new ArgumentException("单元格的SharedStringId无效!");
                }
            }
            else
            {
                //如果不是，直接创建
                CreateCellValue(spreadsheetDocument, worksheetPart, cell, newValue);
            }
        }
        private void CreateCellValue<T>(SpreadsheetDocument spreadsheetDocument, WorksheetPart worksheetPart, Cell cell, T newValue)
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();
            //需对传入的Text进行判断，如果是数值型的直接填入，如果是字符串，再填入到SharedStringItem
            var value = newValue.ToString();
            Decimal itemValue;
            //对数值进行Decimal判断，如果解析成功，则直接存入，未成功，则为string类型
            if (Decimal.TryParse(value, out itemValue))
            {
                cell.CellValue = new CellValue(value);
            }
            else
            {
                if (sharedStringTablePart == null)
                {
                    sharedStringTablePart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                var index = InsertSharedStringItem(value, sharedStringTablePart);
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.CellValue = new CellValue(index.ToString());
            }
            worksheetPart.Worksheet.Save();
        }

        private WorksheetPart CreateWorksheetPart(SpreadsheetDocument spreadsheetDocument)
        {
            //WorksheetPart worksheetPart = null;
            #region ini spreadsheetDocument
            WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.
                AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = spreadsheetDocument.WorkbookPart.
                GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Report"
            };
            sheets.Append(sheet);

            workbookPart.Workbook.Save();
            #endregion
            return worksheetPart;
        }

        private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }


        // Given a shared string ID and a SpreadsheetDocument, verifies that other cells in the document no longer 
        // reference the specified SharedStringItem and removes the item.
        private void RemoveSharedStringItem(int shareStringId, SpreadsheetDocument document)
        {
            bool remove = true;

            foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
            {
                Worksheet worksheet = part.Worksheet;
                foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                {
                    // Verify if other cells in the document reference the item.
                    if (cell.DataType != null &&
                        cell.DataType.Value == CellValues.SharedString &&
                        cell.CellValue.Text == shareStringId.ToString())
                    {
                        // Other cells in the document still reference the item. Do not remove the item.
                        remove = false;
                        break;
                    }
                }

                if (!remove)
                {
                    break;
                }
            }

            // Other cells in the document do not reference the item. Remove the item.
            if (remove)
            {
                SharedStringTablePart shareStringTablePart = document.WorkbookPart.SharedStringTablePart;
                if (shareStringTablePart == null)
                {
                    return;
                }

                SharedStringItem item = shareStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(shareStringId);
                if (item != null)
                {
                    item.Remove();

                    // Refresh all the shared string references.
                    foreach (var part in document.WorkbookPart.GetPartsOfType<WorksheetPart>())
                    {
                        Worksheet worksheet = part.Worksheet;
                        foreach (var cell in worksheet.GetFirstChild<SheetData>().Descendants<Cell>())
                        {
                            if (cell.DataType != null &&
                                cell.DataType.Value == CellValues.SharedString)
                            {
                                int itemIndex = int.Parse(cell.CellValue.Text);
                                if (itemIndex > shareStringId)
                                {
                                    cell.CellValue.Text = (itemIndex - 1).ToString();
                                }
                            }
                        }
                        worksheet.Save();
                    }

                    document.WorkbookPart.SharedStringTablePart.SharedStringTable.Save();
                }
            }
        }






        #endregion

    }
}
