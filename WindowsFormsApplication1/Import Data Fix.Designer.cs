namespace HarmonizedApp
{
    partial class Import_Data_Fix
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Import_Data_Fix));
            this.btn_Import = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_Emplacement = new System.Windows.Forms.Button();
            this.txt_Emplacement = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btn_Import
            // 
            this.btn_Import.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Import.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Import.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Import.Location = new System.Drawing.Point(258, 43);
            this.btn_Import.Name = "btn_Import";
            this.btn_Import.Size = new System.Drawing.Size(143, 33);
            this.btn_Import.TabIndex = 91;
            this.btn_Import.Text = "Import DATA";
            this.btn_Import.UseVisualStyleBackColor = true;
            this.btn_Import.Click += new System.EventHandler(this.btn_Import_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(15, 82);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(588, 23);
            this.progressBar1.TabIndex = 90;
            this.progressBar1.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label2.Image = ((System.Drawing.Image)(resources.GetObject("label2.Image")));
            this.label2.Location = new System.Drawing.Point(12, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(112, 30);
            this.label2.TabIndex = 89;
            this.label2.Text = "Initial Location\r\nFile of Gross to Net";
            // 
            // btn_Emplacement
            // 
            this.btn_Emplacement.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Emplacement.Location = new System.Drawing.Point(552, 9);
            this.btn_Emplacement.Name = "btn_Emplacement";
            this.btn_Emplacement.Size = new System.Drawing.Size(48, 23);
            this.btn_Emplacement.TabIndex = 88;
            this.btn_Emplacement.Text = "...";
            this.btn_Emplacement.UseVisualStyleBackColor = true;
            this.btn_Emplacement.Click += new System.EventHandler(this.btn_Emplacement_Click);
            // 
            // txt_Emplacement
            // 
            this.txt_Emplacement.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_Emplacement.Location = new System.Drawing.Point(145, 9);
            this.txt_Emplacement.Name = "txt_Emplacement";
            this.txt_Emplacement.Size = new System.Drawing.Size(402, 21);
            this.txt_Emplacement.TabIndex = 87;
            // 
            // Import_Data_Fix
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(634, 113);
            this.Controls.Add(this.btn_Import);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_Emplacement);
            this.Controls.Add(this.txt_Emplacement);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Import_Data_Fix";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Data Fix";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Import;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_Emplacement;
        private System.Windows.Forms.TextBox txt_Emplacement;
    }
}