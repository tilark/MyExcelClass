using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MyExcelClass;
using TestReadExcel.TestModels;

namespace TestReadExcel
{
    class Program
    {

        static void Main(string[] args)
        {
            TestExcelExportHelper();
            Console.ReadLine();
        }

        public static void TestExcelExportHelper()
        {
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IWriteToExcel writeToExcel = factory.CreateWriteToExcel();

            List<Student> lstStudent = StaticDataOfStudent.ListStudent;
            var titleHeading = new List<string>();
            titleHeading.Add("Test");
            titleHeading.Add(DateTime.Now.ToString());
            string[] columns = { "ID", "Name", "Age" };
            byte[] filecontent = writeToExcel.ExportExcel(lstStudent, titleHeading, true, columns);
            foreach(var item in filecontent)
            {
                Console.Write(item);
            }
        }
        public static void TestWriteToExcelByCellReplace()
        {
            string fileName = @"E:\net_dev\TestData\test2.xltx";
            string desFile = @"E:\net_dev\TestData\test2Des5.xlsx";
            TestTransferTemplateToWorkBook(fileName, desFile);
            //测试能不能将sourceFile中的内容写入到desFile中
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IWriteToExcel writeToExcel = factory.CreateWriteToExcel();
            //writeToExcel.WriteToExcelByCell(fileName, "A7", 9);
            writeToExcel.WriteToExcelByCellReplace(desFile, "A1", "测试替换标题项");
        }
        public static void TestWriteToExcelByRow()
        {
            string fileName = @"E:\net_dev\TestData\test2.xltx";
            string desFile = @"E:\net_dev\TestData\test2Des3.xlsx";
            TestTransferTemplateToWorkBook(fileName, desFile);
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IWriteToExcel writeToExcel = factory.CreateWriteToExcel();
            List<int> listData = new List<int>();
            for(int i = 0; i < 3; i++)
            {
                listData.Add(i);
            }

            writeToExcel.WriteToExcelByRow(desFile, "B3", listData);
        }
        public static void TestWriteToExcelByCell()
        {
            string fileName = @"E:\net_dev\TestData\test2.xltx";
            string desFile = @"E:\net_dev\TestData\test2Des2.xlsx";
            TestTransferTemplateToWorkBook(fileName, desFile);
            //测试能不能将sourceFile中的内容写入到desFile中
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IWriteToExcel writeToExcel = factory.CreateWriteToExcel();
            //writeToExcel.WriteToExcelByCell(fileName, "A7", 9);
            writeToExcel.WriteToExcelByCell(desFile, "A1", "测试添加标题附加项");

        }
        public static void TestTransferTemplateToWorkBook(string fileName, string desFile)
        {
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IWriteToExcel writeToExcel = factory.CreateWriteToExcel();

            var result = writeToExcel.TransferTemplateToWorkBook(fileName, desFile);
            if (!result)
            {
                Console.WriteLine("File is not exist");
            }
        }
        public static void TestTransferSoureFileToDesFile(string fileName, string desFile)
        {
            //测试能不能将sourceFile中的内容写入到desFile中
            
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IWriteToExcel writeToExcel = factory.CreateWriteToExcel();

            //var result = writeToExcel.TransferSoureFileToDesFile(fileName, desFile);
            //if (!result)
            //{
            //    Console.WriteLine("File is not exist");
            //}

        }
        //public static void TestTransferSoureFileToDesFile()
        //{
        //    //测试能不能将sourceFile中的内容写入到desFile中
        //    string fileName = @"E:\net_dev\TestData\test1.xltx";
        //    string desFile = @"E:\net_dev\TestData\test1Des2.xlsx";
        //    MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
        //    IWriteToExcel writeToExcel = factory.CreateWriteToExcel();

        //    //var result = writeToExcel.TransferSoureFileToDesFile(fileName, desFile);
        //    //if(!result)
        //    //{
        //    //    Console.WriteLine("File is not exist");
        //    //}

        //}
        public static void TestReadFromExcelByCell()
        {
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IReadFromExcel readFromExcel = factory.CreateReadFromExcel();
            string fileName = @"E:\net_dev\TestData\test1.xlsx";
            var result = readFromExcel.ReadFromExcelByCell(fileName, "B10");
            Console.WriteLine("The cell value is {0}", result);
        }
        public static void TestReadFromExcelByCellRange()
        {
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IReadFromExcel readFromExcel = factory.CreateReadFromExcel();
            string fileName = @"E:\net_dev\TestData\test1.xlsx";
            var result = readFromExcel.ReadFromExcelByCellRange(fileName, "A1", "E3");
            foreach (var item in result)
            {
                Console.WriteLine("Key : {0}  , Value: {1}", item.Key, item.Value);
            }
        }
        public static void TestReadFromExcelByDOM()
        {
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IReadFromExcel readFromExcel = factory.CreateReadFromExcel();
            string fileName = @"E:\net_dev\TestData\设备信息.xlsx";
            var result = readFromExcel.ReadFromExcelFileByDOM(fileName);
            foreach( var item in result)
            {
                Console.WriteLine("Key : {0}  , Value: {1}", item.Key, item.Value);
            }

        }
        public static void TestGetRowCount()
        {
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IReadFromExcel readFromExcel = factory.CreateReadFromExcel();
            string fileName = @"E:\net_dev\TestData\设备信息.xlsx";
            var result = readFromExcel.GetRowCount(fileName);
            Console.WriteLine("The row count is {0}", result);
        }

        public static void TestReadFromExcelByRow()
        {
            MyExcelClassFactory factory = MyExcelClassFactory.GetInstance();
            IReadFromExcel readFromExcel = factory.CreateReadFromExcel();
            string fileName = @"E:\net_dev\TestData\test1.xlsx";
            var result = readFromExcel.ReadFromExcelByRow(fileName, 1);
            foreach(var item in result)
            {
                Console.WriteLine("Key is:{0}, Value is{1}", item.Key, item.Value);
            }
        }

    }
}
