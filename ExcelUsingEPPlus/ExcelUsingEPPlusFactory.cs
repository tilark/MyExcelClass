using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUsingEPPlus
{
    public class ExcelUsingEPPlusFactory
    {
        #region single pattern
        private ExcelUsingEPPlusFactory()
        {

        }

        // A private static instance of the same class
        private static readonly ExcelUsingEPPlusFactory instance = null;

        static ExcelUsingEPPlusFactory()
        {
            // create the instance only if the instance is null
            instance = new ExcelUsingEPPlusFactory();
        }

        public static ExcelUsingEPPlusFactory GetInstance()
        {
            // return the already existing instance
            return instance;
        }
        #endregion


        public IWriteToExcel CreateWriteToExcel()
        {
            
            return new WriteToExcel();
        }
    }
}
