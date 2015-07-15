using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord
{
    abstract class Import
    {
        protected string fileToExcel;

        public Import(string fileExcel)
        {
            this.fileToExcel = fileExcel;
        }
    }
}
