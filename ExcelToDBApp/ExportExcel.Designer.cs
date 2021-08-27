namespace ExcelToDBApp
{
    partial class ExportExcel
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
            this.btnBrowse = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.cboSheet = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.Sheet = new System.Windows.Forms.Label();
            this.btnImport = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnBrowse
            // 
            this.btnBrowse.Location = new System.Drawing.Point(741, 345);
            this.btnBrowse.Name = "btnBrowse";
            this.btnBrowse.Size = new System.Drawing.Size(62, 23);
            this.btnBrowse.TabIndex = 0;
            this.btnBrowse.Text = "...";
            this.btnBrowse.UseVisualStyleBackColor = true;
            this.btnBrowse.Click += new System.EventHandler(this.btnBrowse_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(791, 325);
            this.dataGridView1.TabIndex = 1;
            // 
            // cboSheet
            // 
            this.cboSheet.FormattingEnabled = true;
            this.cboSheet.Location = new System.Drawing.Point(75, 381);
            this.cboSheet.Name = "cboSheet";
            this.cboSheet.Size = new System.Drawing.Size(225, 21);
            this.cboSheet.TabIndex = 2;
            this.cboSheet.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 348);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "File Name:";
            // 
            // txtFileName
            // 
            this.txtFileName.Location = new System.Drawing.Point(75, 346);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.ReadOnly = true;
            this.txtFileName.Size = new System.Drawing.Size(660, 20);
            this.txtFileName.TabIndex = 4;
            // 
            // Sheet
            // 
            this.Sheet.AutoSize = true;
            this.Sheet.Location = new System.Drawing.Point(15, 388);
            this.Sheet.Name = "Sheet";
            this.Sheet.Size = new System.Drawing.Size(38, 13);
            this.Sheet.TabIndex = 5;
            this.Sheet.Text = "Sheet:";
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(306, 379);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(75, 23);
            this.btnImport.TabIndex = 6;
            this.btnImport.Text = "&Import";
            this.btnImport.UseVisualStyleBackColor = true;
            this.btnImport.Click += new System.EventHandler(this.btnImport_Click);
            // 
            // ExportExcel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(815, 423);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.Sheet);
            this.Controls.Add(this.txtFileName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboSheet);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnBrowse);
            this.Name = "ExportExcel";
            this.Text = "Read Excel";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnBrowse;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.ComboBox cboSheet;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Label Sheet;
        private System.Windows.Forms.Button btnImport;
    }
}

