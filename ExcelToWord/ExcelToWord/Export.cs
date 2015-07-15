using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToWord
{
    abstract class Export
    {
        
        protected string fileToWord;

        public Export(string fileToWord)
        {
            this.fileToWord = fileToWord;
        }
    }
}
