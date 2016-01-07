namespace HarmonizedApp
{
    partial class Import_Rubrique
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Import_Rubrique));
            this.btn_Import = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_Emplacement = new System.Windows.Forms.Button();
            this.txt_Emplacement = new System.Windows.Forms.TextBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.CodeRub = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Libelles = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.idSoc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.OrdreSeq = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // btn_Import
            // 
            this.btn_Import.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Import.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Import.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Import.Location = new System.Drawing.Point(258, 53);
            this.btn_Import.Name = "btn_Import";
            this.btn_Import.Size = new System.Drawing.Size(143, 33);
            this.btn_Import.TabIndex = 86;
            this.btn_Import.Text = "Import File";
            this.btn_Import.UseVisualStyleBackColor = true;
            this.btn_Import.Click += new System.EventHandler(this.btn_Import_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label2.Image = ((System.Drawing.Image)(resources.GetObject("label2.Image")));
            this.label2.Location = new System.Drawing.Point(12, 19);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(68, 30);
            this.label2.TabIndex = 85;
            this.label2.Text = "Path of the \r\nExcel File";
            // 
            // btn_Emplacement
            // 
            this.btn_Emplacement.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Emplacement.Location = new System.Drawing.Point(552, 19);
            this.btn_Emplacement.Name = "btn_Emplacement";
            this.btn_Emplacement.Size = new System.Drawing.Size(48, 23);
            this.btn_Emplacement.TabIndex = 84;
            this.btn_Emplacement.Text = "...";
            this.btn_Emplacement.UseVisualStyleBackColor = true;
            this.btn_Emplacement.Click += new System.EventHandler(this.btn_Emplacement_Click);
            // 
            // txt_Emplacement
            // 
            this.txt_Emplacement.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_Emplacement.Location = new System.Drawing.Point(145, 19);
            this.txt_Emplacement.Name = "txt_Emplacement";
            this.txt_Emplacement.Size = new System.Drawing.Size(402, 21);
            this.txt_Emplacement.TabIndex = 83;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CodeRub,
            this.Libelles,
            this.idSoc,
            this.OrdreSeq});
            this.dataGridView1.Location = new System.Drawing.Point(15, 92);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(442, 150);
            this.dataGridView1.TabIndex = 87;
            this.dataGridView1.Visible = false;
            // 
            // CodeRub
            // 
            this.CodeRub.HeaderText = "CodeRub";
            this.CodeRub.Name = "CodeRub";
            // 
            // Libelles
            // 
            this.Libelles.HeaderText = "Libelles";
            this.Libelles.Name = "Libelles";
            // 
            // idSoc
            // 
            this.idSoc.HeaderText = "idSoc";
            this.idSoc.Name = "idSoc";
            // 
            // OrdreSeq
            // 
            this.OrdreSeq.HeaderText = "ordreseq";
            this.OrdreSeq.Name = "OrdreSeq";
            // 
            // Import_Rubrique
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(615, 90);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btn_Import);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_Emplacement);
            this.Controls.Add(this.txt_Emplacement);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Import_Rubrique";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Rubrique";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Import;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_Emplacement;
        private System.Windows.Forms.TextBox txt_Emplacement;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn CodeRub;
        private System.Windows.Forms.DataGridViewTextBoxColumn Libelles;
        private System.Windows.Forms.DataGridViewTextBoxColumn idSoc;
        private System.Windows.Forms.DataGridViewTextBoxColumn OrdreSeq;
    }
}