namespace Export_to_Excel
{
    partial class frmMacros
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
            this.txtVBACode = new System.Windows.Forms.TextBox();
            this.btnCreateMacro = new System.Windows.Forms.Button();
            this.btnRunMacro = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMacroName = new System.Windows.Forms.TextBox();
            this.cmbMacrosNames = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.SuspendLayout();
            // 
            // txtVBACode
            // 
            this.txtVBACode.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtVBACode.Location = new System.Drawing.Point(20, 97);
            this.txtVBACode.Multiline = true;
            this.txtVBACode.Name = "txtVBACode";
            this.txtVBACode.Size = new System.Drawing.Size(424, 398);
            this.txtVBACode.TabIndex = 0;
            // 
            // btnCreateMacro
            // 
            this.btnCreateMacro.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateMacro.Location = new System.Drawing.Point(19, 507);
            this.btnCreateMacro.Name = "btnCreateMacro";
            this.btnCreateMacro.Size = new System.Drawing.Size(156, 34);
            this.btnCreateMacro.TabIndex = 1;
            this.btnCreateMacro.Text = "Save Macro";
            this.btnCreateMacro.UseVisualStyleBackColor = true;
            this.btnCreateMacro.Click += new System.EventHandler(this.btnCreateMacro_Click);
            // 
            // btnRunMacro
            // 
            this.btnRunMacro.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRunMacro.Location = new System.Drawing.Point(214, 507);
            this.btnRunMacro.Name = "btnRunMacro";
            this.btnRunMacro.Size = new System.Drawing.Size(156, 34);
            this.btnRunMacro.TabIndex = 2;
            this.btnRunMacro.Text = "Run Macro";
            this.btnRunMacro.UseVisualStyleBackColor = true;
            this.btnRunMacro.Click += new System.EventHandler(this.btnRunMacro_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(17, 78);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(186, 16);
            this.label1.TabIndex = 3;
            this.label1.Text = "Write Macro Code in VBA:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(17, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 16);
            this.label2.TabIndex = 4;
            this.label2.Text = "Enter Macro Name:";
            // 
            // txtMacroName
            // 
            this.txtMacroName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMacroName.Location = new System.Drawing.Point(20, 39);
            this.txtMacroName.Name = "txtMacroName";
            this.txtMacroName.Size = new System.Drawing.Size(203, 23);
            this.txtMacroName.TabIndex = 5;
            // 
            // cmbMacrosNames
            // 
            this.cmbMacrosNames.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMacrosNames.FormattingEnabled = true;
            this.cmbMacrosNames.Location = new System.Drawing.Point(262, 38);
            this.cmbMacrosNames.Name = "cmbMacrosNames";
            this.cmbMacrosNames.Size = new System.Drawing.Size(182, 24);
            this.cmbMacrosNames.TabIndex = 6;
            this.cmbMacrosNames.SelectedIndexChanged += new System.EventHandler(this.cmbMacrosNames_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(259, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(148, 16);
            this.label3.TabIndex = 7;
            this.label3.Text = "Select Macro Name:";
            // 
            // saveFileDialog1
            // 
            this.saveFileDialog1.Filter = "\"Macro Enabled Files|*.xlsm";
            // 
            // frmMacros
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(470, 556);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmbMacrosNames);
            this.Controls.Add(this.txtMacroName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnRunMacro);
            this.Controls.Add(this.btnCreateMacro);
            this.Controls.Add(this.txtVBACode);
            this.Name = "frmMacros";
            this.Text = "frmMacros";
            this.Deactivate += new System.EventHandler(this.frmMacros_Deactivate);
            this.Load += new System.EventHandler(this.frmMacros_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion 	Export to Excel.exe!Export_to_Excel.Program.Main() Line 19	C#


        private System.Windows.Forms.TextBox txtVBACode;
        private System.Windows.Forms.Button btnCreateMacro;
        private System.Windows.Forms.Button btnRunMacro;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMacroName;
        private System.Windows.Forms.ComboBox cmbMacrosNames;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}