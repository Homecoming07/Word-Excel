namespace ExcelToWord
{
    partial class FormMain
    {
        /// <summary>
        /// Требуется переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Обязательный метод для поддержки конструктора - не изменяйте
        /// содержимое данного метода при помощи редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonExport = new System.Windows.Forms.Button();
            this.tBtoExcel = new System.Windows.Forms.TextBox();
            this.labelPathExcel = new System.Windows.Forms.Label();
            this.labelPathExport = new System.Windows.Forms.Label();
            this.tBtoWord = new System.Windows.Forms.TextBox();
            this.cBformatExp = new System.Windows.Forms.ComboBox();
            this.labelFormat = new System.Windows.Forms.Label();
            this.buttonToOpenFile = new System.Windows.Forms.Button();
            this.buttonToSaveFile = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonExport
            // 
            this.buttonExport.Enabled = false;
            this.buttonExport.Location = new System.Drawing.Point(12, 274);
            this.buttonExport.Margin = new System.Windows.Forms.Padding(1, 4, 1, 4);
            this.buttonExport.Name = "buttonExport";
            this.buttonExport.Size = new System.Drawing.Size(160, 40);
            this.buttonExport.TabIndex = 0;
            this.buttonExport.Text = "Экспорт";
            this.buttonExport.UseVisualStyleBackColor = true;
            this.buttonExport.Click += new System.EventHandler(this.buttonExport_Click);
            // 
            // tBtoExcel
            // 
            this.tBtoExcel.Location = new System.Drawing.Point(154, 58);
            this.tBtoExcel.Margin = new System.Windows.Forms.Padding(1, 4, 1, 4);
            this.tBtoExcel.Name = "tBtoExcel";
            this.tBtoExcel.ReadOnly = true;
            this.tBtoExcel.Size = new System.Drawing.Size(388, 22);
            this.tBtoExcel.TabIndex = 1;
            // 
            // labelPathExcel
            // 
            this.labelPathExcel.AutoSize = true;
            this.labelPathExcel.Location = new System.Drawing.Point(1, 61);
            this.labelPathExcel.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.labelPathExcel.Name = "labelPathExcel";
            this.labelPathExcel.Size = new System.Drawing.Size(132, 16);
            this.labelPathExcel.TabIndex = 2;
            this.labelPathExcel.Text = "Путь к файлу Excel :";
            // 
            // labelPathExport
            // 
            this.labelPathExport.AutoSize = true;
            this.labelPathExport.Location = new System.Drawing.Point(1, 153);
            this.labelPathExport.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.labelPathExport.Name = "labelPathExport";
            this.labelPathExport.Size = new System.Drawing.Size(151, 16);
            this.labelPathExport.TabIndex = 4;
            this.labelPathExport.Text = "Путь к файлу Шаблона :";
            // 
            // tBtoWord
            // 
            this.tBtoWord.Location = new System.Drawing.Point(154, 150);
            this.tBtoWord.Margin = new System.Windows.Forms.Padding(1, 4, 1, 4);
            this.tBtoWord.Name = "tBtoWord";
            this.tBtoWord.ReadOnly = true;
            this.tBtoWord.Size = new System.Drawing.Size(388, 22);
            this.tBtoWord.TabIndex = 3;
            // 
            // cBformatExp
            // 
            this.cBformatExp.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cBformatExp.FormattingEnabled = true;
            this.cBformatExp.Items.AddRange(new object[] {
            ".docx",
            ".pdf"});
            this.cBformatExp.Location = new System.Drawing.Point(127, 24);
            this.cBformatExp.Margin = new System.Windows.Forms.Padding(1, 4, 1, 4);
            this.cBformatExp.Name = "cBformatExp";
            this.cBformatExp.Size = new System.Drawing.Size(62, 24);
            this.cBformatExp.TabIndex = 5;
            // 
            // labelFormat
            // 
            this.labelFormat.AutoSize = true;
            this.labelFormat.Location = new System.Drawing.Point(10, 27);
            this.labelFormat.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.labelFormat.Name = "labelFormat";
            this.labelFormat.Size = new System.Drawing.Size(115, 16);
            this.labelFormat.TabIndex = 6;
            this.labelFormat.Text = "Формат экспорта:";
            // 
            // buttonToOpenFile
            // 
            this.buttonToOpenFile.Location = new System.Drawing.Point(545, 56);
            this.buttonToOpenFile.Margin = new System.Windows.Forms.Padding(1, 4, 1, 4);
            this.buttonToOpenFile.Name = "buttonToOpenFile";
            this.buttonToOpenFile.Size = new System.Drawing.Size(50, 24);
            this.buttonToOpenFile.TabIndex = 7;
            this.buttonToOpenFile.Text = "...";
            this.buttonToOpenFile.UseVisualStyleBackColor = true;
            this.buttonToOpenFile.Click += new System.EventHandler(this.buttonToOpenFile_Click);
            // 
            // buttonToSaveFile
            // 
            this.buttonToSaveFile.Location = new System.Drawing.Point(545, 148);
            this.buttonToSaveFile.Margin = new System.Windows.Forms.Padding(1, 4, 1, 4);
            this.buttonToSaveFile.Name = "buttonToSaveFile";
            this.buttonToSaveFile.Size = new System.Drawing.Size(50, 26);
            this.buttonToSaveFile.TabIndex = 8;
            this.buttonToSaveFile.Text = "...";
            this.buttonToSaveFile.UseVisualStyleBackColor = true;
            this.buttonToSaveFile.Click += new System.EventHandler(this.buttonToSaveFile_Click);
            // 
            // label1
            // 
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(154, 12);
            this.label1.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(391, 42);
            this.label1.TabIndex = 9;
            this.label1.Text = "Таблица Excel должна быть заполнена полностью в виде прямоугольного пространства." +
    "\r\n";
            // 
            // label2
            // 
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(154, 104);
            this.label2.Margin = new System.Windows.Forms.Padding(1, 0, 1, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(391, 42);
            this.label2.TabIndex = 10;
            this.label2.Text = "Файл шаблона должен содержать имена заголовков из Excel файла в фигурных скобках " +
    "{}.";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.labelPathExport);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.tBtoExcel);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.labelPathExcel);
            this.panel1.Controls.Add(this.buttonToSaveFile);
            this.panel1.Controls.Add(this.tBtoWord);
            this.panel1.Controls.Add(this.buttonToOpenFile);
            this.panel1.Location = new System.Drawing.Point(12, 65);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(604, 202);
            this.panel1.TabIndex = 11;
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(628, 322);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.labelFormat);
            this.Controls.Add(this.cBformatExp);
            this.Controls.Add(this.buttonExport);
            this.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(1, 4, 1, 4);
            this.MaximizeBox = false;
            this.Name = "FormMain";
            this.Text = "Программа beta";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormMain_FormClosing);
            this.Load += new System.EventHandler(this.FormMain_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button buttonExport;
        private System.Windows.Forms.TextBox tBtoExcel;
        private System.Windows.Forms.Label labelPathExcel;
        private System.Windows.Forms.Label labelPathExport;
        private System.Windows.Forms.TextBox tBtoWord;
        private System.Windows.Forms.ComboBox cBformatExp;
        private System.Windows.Forms.Label labelFormat;
        private System.Windows.Forms.Button buttonToOpenFile;
        private System.Windows.Forms.Button buttonToSaveFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel1;
    }
}

