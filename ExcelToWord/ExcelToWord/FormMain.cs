using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;



namespace ExcelToWord
{
    public partial class FormMain : Form
    {
        ExcelApplication excelDoc;
        WordApplication wordDoc;

        string fileToExcel;
        string fileToWord;
        string fileToSavePath;

        public FormMain()
        {
            InitializeComponent();

        }

        private void buttonToOpenFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "Excel 2010-16 |*.xlsx|Excel 2003|*.xls|Все файлы|*.*";
            if (openfile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileToExcel = openfile.FileName;
                tBtoExcel.Text = fileToExcel;
            }

            EnableButtonExport();

        }

        private void buttonToSaveFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.Filter = "Word 2010-16 |*.docx|Word 2003|*.doc|Все файлы|*.*";
            if (openfile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileToWord = openfile.FileName;
                tBtoWord.Text = fileToWord;
            }

            EnableButtonExport();
        }

        private void buttonExport_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();

            if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileToSavePath = folderBrowser.SelectedPath + "\\";
            }

           
            excelDoc = new ExcelApplication(fileToExcel);
            excelDoc.GetValueToDictionary(1);

            wordDoc = new WordApplication(fileToWord);

            string[] labels = excelDoc.valuesDictionary[1];
            foreach (KeyValuePair<int, string[]> item in excelDoc.valuesDictionary)
            {

                if (item.Key == 1)
                {
                    labels = item.Value;
                }
                else
                {
                    for (int i = 0; i < labels.Length; i++)
                    {
                        wordDoc.FindAndReplace(labels[i], item.Value[i]);
                    }
                    wordDoc.SaveDocument(fileToSavePath + (item.Key - 1).ToString(), wordDoc.SafeFileFormat(cBformatExp.SelectedItem.ToString()));
                    wordDoc.CloseWordApp();
                    wordDoc = new WordApplication(fileToWord);
                }
            }

          
        }

        private void EnableButtonExport()
        {
            if (tBtoExcel.Text != "" && tBtoWord.Text != "")
            buttonExport.Enabled = true;
        }

        private void FormMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (excelDoc != null)
            {
                excelDoc.CloseExcelDoc();
                wordDoc.CloseWordApp(); 
            }

        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            ToolTip toolTip1 = new ToolTip();

            toolTip1.ToolTipIcon = ToolTipIcon.Warning;

            toolTip1.ToolTipTitle = "Внимание";

            toolTip1.IsBalloon = true;

            toolTip1.BackColor = Color.Yellow;

            toolTip1.ForeColor = Color.Red;
            toolTip1.SetToolTip(buttonToSaveFile, "Укажите файл шаблона Word.");
            toolTip1.SetToolTip(buttonToOpenFile, "Укажите где находится Excel файл.");

            cBformatExp.SelectedIndex = 0;
        }


    }
}
