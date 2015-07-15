using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToWord
{
    class ExcelApplication : Import
    {
        public Dictionary<int, string[]> valuesDictionary = new Dictionary<int, string[]>();

        protected Excel.Application excelApp;  //Переменная с приложением excel
        protected Excel.Workbook excelBook;    //Переменная с текущем листом
        protected Excel.Worksheet excelSheet;  //Переменная с ячейками
        protected Excel.Range excelRange;      //Переменная для работы с диапозоном ячеек

        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false; 


        /// <summary>
        /// Конструктор для инициалицазии переменных класса
        /// </summary>
        /// <param name="fileExcel">Путь к файлу для открытия</param>
        public ExcelApplication(string fileExcel) : base(fileExcel)
        {
            try
            {
            excelApp = new Excel.Application();
            excelBook = excelApp.Workbooks.Open(fileToExcel);
            excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(1);
            excelRange = excelSheet.UsedRange;
            }
            catch 
            {
                MessageBox.Show("Заданный файл Excel не доступен!");
                //excelBook.Close(falseObj,missingObj,missingObj);
                excelApp.Quit();
                excelBook = null;
                excelApp = null;
                Application.Exit();
            }

        }

        /// <summary>
        /// Метод для получения значений из диапозона
        /// </summary>
        public void GetValueToDictionary(int startTofield)
        {
            try
            {

                //Создаем по начальной строке массив заголовков
                string[] arrayLabel = new string[excelRange.Columns.Count];
                int index = 0;
                for (int row = startTofield; row < startTofield + 1; row++)
                {
                    for (int col = 1; col <= excelRange.Columns.Count; col++)
                    {
                        arrayLabel[index++] = (string)(excelRange.Cells[row, col] as Excel.Range).Value2;
                    }
                }

                //Заносим массив заголовков в словарь
                valuesDictionary.Add(startTofield, arrayLabel);

                //Получаем следующие значия строк за заголовками
                for (int row = startTofield+1; row <= excelRange.Rows.Count; row++)
                {
                    int index2 = 0;
                    string[] arrayValue = new string[excelRange.Columns.Count];

                    for (int col = 1; col <= excelRange.Columns.Count; col++)
                    {
                        arrayValue[index2++] = (excelRange.Cells[row, col] as Excel.Range).Value2.ToString();
                    }

                    //Заносим массив значений в словарь если строка полная
                    if (arrayValue.Length == excelRange.Columns.Count)
                    {

                        valuesDictionary.Add(++startTofield, arrayValue);
                    }


                    arrayValue = null;//После каждой итерации необходимо убрать ссылку на массив
                }

            }
            catch
            {
                MessageBox.Show("Возникла ошибка с диапозоном значений, проверте чтобы они представляли прямоугольник!");
                return;
            }
          
        }

        /// <summary>
        /// Метод закрытия приложения Excel
        /// </summary>
        public void CloseExcelDoc()
        {
            excelBook.Close(falseObj, missingObj, missingObj);
            excelApp.Quit();
            excelBook = null;
            excelApp = null;
        }

    }
}
