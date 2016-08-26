namespace Export_to_Excel
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.No1 = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.cmb1 = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cmb2 = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.No2 = new System.Windows.Forms.NumericUpDown();
            this.txtEnter = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnExportWordToExcel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnCloseExcel = new System.Windows.Forms.Button();
            this.btnCreateChart = new System.Windows.Forms.Button();
            this.btnOpen = new System.Windows.Forms.Button();
            this.btnSaveAs = new System.Windows.Forms.Button();
            this.btnCreateExcel = new System.Windows.Forms.Button();
            this.btnSaveText = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnFormulas_Dialog = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.fontDialog1 = new System.Windows.Forms.FontDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnExportTxtToExcel = new System.Windows.Forms.Button();
            this.btnMacros = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.No1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.No2)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // No1
            // 
            this.No1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.No1.Location = new System.Drawing.Point(648, 26);
            this.No1.Name = "No1";
            this.No1.Size = new System.Drawing.Size(105, 23);
            this.No1.TabIndex = 0;
            this.No1.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(562, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 17);
            this.label1.TabIndex = 1;
            this.label1.Text = "First Row:";
            // 
            // cmb1
            // 
            this.cmb1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmb1.FormattingEnabled = true;
            this.cmb1.Items.AddRange(new object[] {
            "A",
            "B",
            "C",
            "D",
            "E",
            "F",
            "G",
            "H",
            "I",
            "J",
            "K",
            "L",
            "M",
            "N",
            "O",
            "P",
            "Q",
            "R",
            "S",
            "T",
            "U",
            "V",
            "W",
            "X",
            "Y",
            "Z"});
            this.cmb1.Location = new System.Drawing.Point(421, 24);
            this.cmb1.Name = "cmb1";
            this.cmb1.Size = new System.Drawing.Size(80, 24);
            this.cmb1.TabIndex = 2;
            this.cmb1.Text = "A";
            this.cmb1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(312, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(103, 17);
            this.label2.TabIndex = 3;
            this.label2.Text = "First Column:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(290, 79);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(125, 17);
            this.label3.TabIndex = 7;
            this.label3.Text = "Second Column:";
            // 
            // cmb2
            // 
            this.cmb2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmb2.FormattingEnabled = true;
            this.cmb2.Items.AddRange(new object[] {
            "A",
            "B",
            "C",
            "D",
            "E",
            "F",
            "G",
            "H",
            "I",
            "J",
            "K",
            "L",
            "M",
            "N",
            "O",
            "P",
            "Q",
            "R",
            "S",
            "T",
            "U",
            "V",
            "W",
            "X",
            "Y",
            "Z"});
            this.cmb2.Location = new System.Drawing.Point(421, 71);
            this.cmb2.Name = "cmb2";
            this.cmb2.Size = new System.Drawing.Size(80, 24);
            this.cmb2.TabIndex = 6;
            this.cmb2.Text = "H";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(540, 78);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(102, 17);
            this.label4.TabIndex = 5;
            this.label4.Text = "Second Row:";
            // 
            // No2
            // 
            this.No2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.No2.Location = new System.Drawing.Point(648, 73);
            this.No2.Name = "No2";
            this.No2.Size = new System.Drawing.Size(105, 23);
            this.No2.TabIndex = 4;
            this.No2.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // txtEnter
            // 
            this.txtEnter.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtEnter.Location = new System.Drawing.Point(421, 118);
            this.txtEnter.Name = "txtEnter";
            this.txtEnter.Size = new System.Drawing.Size(332, 23);
            this.txtEnter.TabIndex = 8;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(254, 121);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(161, 17);
            this.label5.TabIndex = 9;
            this.label5.Text = "Enter values to cells:";
            // 
            // btnExportWordToExcel
            // 
            this.btnExportWordToExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportWordToExcel.Location = new System.Drawing.Point(59, 389);
            this.btnExportWordToExcel.Name = "btnExportWordToExcel";
            this.btnExportWordToExcel.Size = new System.Drawing.Size(332, 34);
            this.btnExportWordToExcel.TabIndex = 10;
            this.btnExportWordToExcel.Text = "Export Word Document To Excel";
            this.btnExportWordToExcel.UseVisualStyleBackColor = true;
            this.btnExportWordToExcel.Click += new System.EventHandler(this.btnExportWordToExcel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnCloseExcel);
            this.groupBox1.Controls.Add(this.btnCreateChart);
            this.groupBox1.Controls.Add(this.btnOpen);
            this.groupBox1.Controls.Add(this.btnSaveAs);
            this.groupBox1.Controls.Add(this.btnCreateExcel);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(40, 24);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(177, 290);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Excel File";
            // 
            // btnCloseExcel
            // 
            this.btnCloseExcel.Enabled = false;
            this.btnCloseExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCloseExcel.Location = new System.Drawing.Point(19, 242);
            this.btnCloseExcel.Name = "btnCloseExcel";
            this.btnCloseExcel.Size = new System.Drawing.Size(137, 34);
            this.btnCloseExcel.TabIndex = 15;
            this.btnCloseExcel.Text = "Close";
            this.btnCloseExcel.UseVisualStyleBackColor = true;
            this.btnCloseExcel.Click += new System.EventHandler(this.btnCloseExcel_Click);
            // 
            // btnCreateChart
            // 
            this.btnCreateChart.Enabled = false;
            this.btnCreateChart.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateChart.Location = new System.Drawing.Point(19, 190);
            this.btnCreateChart.Name = "btnCreateChart";
            this.btnCreateChart.Size = new System.Drawing.Size(137, 34);
            this.btnCreateChart.TabIndex = 14;
            this.btnCreateChart.Text = "Create Chart";
            this.btnCreateChart.UseVisualStyleBackColor = true;
            this.btnCreateChart.Click += new System.EventHandler(this.btnCreateChart_Click);
            // 
            // btnOpen
            // 
            this.btnOpen.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpen.Location = new System.Drawing.Point(19, 138);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(137, 34);
            this.btnOpen.TabIndex = 13;
            this.btnOpen.Text = "Open";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // btnSaveAs
            // 
            this.btnSaveAs.Enabled = false;
            this.btnSaveAs.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveAs.Location = new System.Drawing.Point(19, 88);
            this.btnSaveAs.Name = "btnSaveAs";
            this.btnSaveAs.Size = new System.Drawing.Size(137, 34);
            this.btnSaveAs.TabIndex = 12;
            this.btnSaveAs.Text = "Save As";
            this.btnSaveAs.UseVisualStyleBackColor = true;
            this.btnSaveAs.Click += new System.EventHandler(this.btnSaveAs_Click);
            // 
            // btnCreateExcel
            // 
            this.btnCreateExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateExcel.Location = new System.Drawing.Point(19, 37);
            this.btnCreateExcel.Name = "btnCreateExcel";
            this.btnCreateExcel.Size = new System.Drawing.Size(137, 34);
            this.btnCreateExcel.TabIndex = 11;
            this.btnCreateExcel.Text = "Create New";
            this.btnCreateExcel.UseVisualStyleBackColor = true;
            this.btnCreateExcel.Click += new System.EventHandler(this.button1_Click);
            // 
            // btnSaveText
            // 
            this.btnSaveText.Enabled = false;
            this.btnSaveText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveText.Location = new System.Drawing.Point(421, 162);
            this.btnSaveText.Name = "btnSaveText";
            this.btnSaveText.Size = new System.Drawing.Size(332, 34);
            this.btnSaveText.TabIndex = 12;
            this.btnSaveText.Text = "Save Text";
            this.btnSaveText.UseVisualStyleBackColor = true;
            this.btnSaveText.Click += new System.EventHandler(this.btnSaveText_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnFormulas_Dialog);
            this.groupBox2.Controls.Add(this.button5);
            this.groupBox2.Controls.Add(this.button4);
            this.groupBox2.Controls.Add(this.button3);
            this.groupBox2.Controls.Add(this.button2);
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(421, 202);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(332, 244);
            this.groupBox2.TabIndex = 13;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Format Excel";
            // 
            // btnFormulas_Dialog
            // 
            this.btnFormulas_Dialog.Enabled = false;
            this.btnFormulas_Dialog.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFormulas_Dialog.Location = new System.Drawing.Point(22, 183);
            this.btnFormulas_Dialog.Name = "btnFormulas_Dialog";
            this.btnFormulas_Dialog.Size = new System.Drawing.Size(289, 34);
            this.btnFormulas_Dialog.TabIndex = 16;
            this.btnFormulas_Dialog.Text = "Formulas Dialog";
            this.btnFormulas_Dialog.UseVisualStyleBackColor = true;
            this.btnFormulas_Dialog.Click += new System.EventHandler(this.btnFormulas_Dialog_Click);
            // 
            // button5
            // 
            this.button5.Enabled = false;
            this.button5.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button5.Location = new System.Drawing.Point(22, 132);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(289, 34);
            this.button5.TabIndex = 16;
            this.button5.Text = "Back Ground Color";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            this.button4.Enabled = false;
            this.button4.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button4.Location = new System.Drawing.Point(174, 81);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(137, 34);
            this.button4.TabIndex = 15;
            this.button4.Text = "Font Color";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Enabled = false;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Location = new System.Drawing.Point(22, 81);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(137, 34);
            this.button3.TabIndex = 14;
            this.button3.Text = "Border";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(174, 28);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(137, 34);
            this.button2.TabIndex = 13;
            this.button2.Text = "Font";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(22, 28);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(137, 34);
            this.button1.TabIndex = 12;
            this.button1.Text = "Merge Cells";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "\"Excel files (*.xls)| *.xlsx\"";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnExportTxtToExcel
            // 
            this.btnExportTxtToExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExportTxtToExcel.Location = new System.Drawing.Point(59, 440);
            this.btnExportTxtToExcel.Name = "btnExportTxtToExcel";
            this.btnExportTxtToExcel.Size = new System.Drawing.Size(332, 34);
            this.btnExportTxtToExcel.TabIndex = 14;
            this.btnExportTxtToExcel.Text = "Export Text File To Excel";
            this.btnExportTxtToExcel.UseVisualStyleBackColor = true;
            this.btnExportTxtToExcel.Click += new System.EventHandler(this.btnExportTxtToExcel_Click);
            // 
            // btnMacros
            // 
            this.btnMacros.Enabled = false;
            this.btnMacros.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMacros.Location = new System.Drawing.Point(59, 334);
            this.btnMacros.Name = "btnMacros";
            this.btnMacros.Size = new System.Drawing.Size(332, 34);
            this.btnMacros.TabIndex = 15;
            this.btnMacros.Text = "Macros ...";
            this.btnMacros.UseVisualStyleBackColor = true;
            this.btnMacros.Click += new System.EventHandler(this.btnMacros_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(819, 529);
            this.Controls.Add(this.btnMacros);
            this.Controls.Add(this.btnExportTxtToExcel);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnSaveText);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnExportWordToExcel);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtEnter);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmb2);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.No2);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmb1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.No1);
            this.Name = "Form1";
            this.Text = " ";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.No1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.No2)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.NumericUpDown No1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmb1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmb2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.NumericUpDown No2;
        private System.Windows.Forms.TextBox txtEnter;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnExportWordToExcel;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnCloseExcel;
        private System.Windows.Forms.Button btnCreateChart;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Button btnSaveAs;
        private System.Windows.Forms.Button btnCreateExcel;
        private System.Windows.Forms.Button btnSaveText;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ColorDialog colorDialog1;
        private System.Windows.Forms.FontDialog fontDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnExportTxtToExcel;
        private System.Windows.Forms.Button btnFormulas_Dialog;
        private System.Windows.Forms.Button btnMacros;
    }
}

