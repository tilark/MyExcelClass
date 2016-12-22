using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelClass
{
    public class MyExcelClassFactory
    {
        #region single pattern
        private MyExcelClassFactory()
        {

        }

        // A private static instance of the same class
        private static readonly MyExcelClassFactory instance = null;

        static MyExcelClassFactory()
        {
            // create the instance only if the instance is null
            instance = new MyExcelClassFactory();
        }

        public static MyExcelClassFactory GetInstance()
        {
            // return the already existing instance
            return instance;
        }
        #endregion

        public IReadFromExcel CreateReadFromExcel()
        {
            IReadFromExcel result = null;
            result = new ReadFromExcel();
            return result;
        }

        public IWriteToExcel CreateWriteToExcel()
        {
            IWriteToExcel result = null;
            result = new WriteToExcel();
            return result;
        }
    }
}
