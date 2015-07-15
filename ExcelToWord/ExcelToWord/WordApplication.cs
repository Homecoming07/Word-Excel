using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

namespace ExcelToWord
{
    class WordApplication: Export
    {

        Word.Application application;
        Word.Document document;
        Word.Range wordRange;// диапазон документа Word 

        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;


        public WordApplication(string fileToWord):base(fileToWord)
        {
            //создаем обьект приложения word  
            application = new Word.Application();

            // создаем путь к файлу шаблона   
            Object templatePathObj = fileToWord;

            // если вылетим не этом этапе, приложение останется открытым  
            try
            {
                document = application.Documents.Add(ref  templatePathObj, ref missingObj, ref missingObj, ref missingObj);
            }

            catch (Exception error)
            {
                CloseWordApp();
                throw error;
            }
        }

        public void FindAndReplace(string findText, string replaceText)
        {
            // обьектные строки для Word  
            object strToFindObj ="{"+findText+"}";
            object replaceStrObj = replaceText;

            //тип поиска и замены  
            object replaceTypeObj;
            replaceTypeObj = Word.WdReplace.wdReplaceAll;

            // обходим все разделы документа  
            for (int i = 1; i <= document.Sections.Count; i++)
            {

                // берем всю секцию диапазоном 
                wordRange = document.Sections[i].Range;

                Word.Find wordFindObj = wordRange.Find;

                object[] wordFindParameters = new object[15] { strToFindObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, missingObj, replaceStrObj, replaceTypeObj, missingObj, missingObj, missingObj, missingObj };

                wordFindObj.GetType().InvokeMember("Execute", BindingFlags.InvokeMethod, null, wordFindObj, wordFindParameters);

            }
        }

        public void SaveDocument(string pathToSave, Word.WdSaveFormat saveFormat)
        {
            Object pathToSaveObj = pathToSave;
            document.SaveAs(ref pathToSaveObj, saveFormat, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj);
        }

        public void CloseWordApp()
        {
            document.Close(ref falseObj, ref  missingObj, ref missingObj);
            application.Quit(ref missingObj, ref  missingObj, ref missingObj);
            document = null;
            application = null;
        }

        public Word.WdSaveFormat SafeFileFormat(string format)
        {
            switch (format)
            {
                case ".docx": return Word.WdSaveFormat.wdFormatDocument;
                case ".pdf": return Word.WdSaveFormat.wdFormatPDF;
                default: return Word.WdSaveFormat.wdFormatDocumentDefault;
            }
        }

    }

}

