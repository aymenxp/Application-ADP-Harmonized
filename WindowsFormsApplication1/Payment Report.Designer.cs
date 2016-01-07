namespace HarmonizedApp
{
    partial class Payment_Report
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Payment_Report));
            this.btn_Import = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.btn_Emplacement = new System.Windows.Forms.Button();
            this.txt_Emplacement = new System.Windows.Forms.TextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btn_Delete = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.chk_delete = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btn_Import
            // 
            this.btn_Import.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Import.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Import.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Import.Location = new System.Drawing.Point(145, 154);
            this.btn_Import.Name = "btn_Import";
            this.btn_Import.Size = new System.Drawing.Size(143, 33);
            this.btn_Import.TabIndex = 86;
            this.btn_Import.Text = "Import DATA";
            this.btn_Import.UseVisualStyleBackColor = true;
            this.btn_Import.Click += new System.EventHandler(this.btn_Import_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label2.Image = ((System.Drawing.Image)(resources.GetObject("label2.Image")));
            this.label2.Location = new System.Drawing.Point(12, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 30);
            this.label2.TabIndex = 85;
            this.label2.Text = "Initial Location\r\nFile of Payment Report";
            // 
            // btn_Emplacement
            // 
            this.btn_Emplacement.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Emplacement.Location = new System.Drawing.Point(552, 20);
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
            this.txt_Emplacement.Location = new System.Drawing.Point(145, 20);
            this.txt_Emplacement.Name = "txt_Emplacement";
            this.txt_Emplacement.Size = new System.Drawing.Size(402, 21);
            this.txt_Emplacement.TabIndex = 83;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 203);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(585, 23);
            this.progressBar1.TabIndex = 87;
            this.progressBar1.Visible = false;
            // 
            // btn_Delete
            // 
            this.btn_Delete.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Delete.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Delete.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Delete.Location = new System.Drawing.Point(332, 154);
            this.btn_Delete.Name = "btn_Delete";
            this.btn_Delete.Size = new System.Drawing.Size(143, 33);
            this.btn_Delete.TabIndex = 88;
            this.btn_Delete.Text = "Delete DATA";
            this.btn_Delete.UseVisualStyleBackColor = true;
            this.btn_Delete.Click += new System.EventHandler(this.button1_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label7.Location = new System.Drawing.Point(12, 114);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(76, 15);
            this.label7.TabIndex = 89;
            this.label7.Text = "Start of Date";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.CalendarFont = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker1.CalendarForeColor = System.Drawing.Color.DodgerBlue;
            this.dateTimePicker1.CalendarTitleForeColor = System.Drawing.Color.DodgerBlue;
            this.dateTimePicker1.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateTimePicker1.Location = new System.Drawing.Point(145, 109);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(183, 22);
            this.dateTimePicker1.TabIndex = 90;
            // 
            // chk_delete
            // 
            this.chk_delete.AutoSize = true;
            this.chk_delete.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("chk_delete.BackgroundImage")));
            this.chk_delete.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chk_delete.ForeColor = System.Drawing.Color.DodgerBlue;
            this.chk_delete.Location = new System.Drawing.Point(145, 70);
            this.chk_delete.Name = "chk_delete";
            this.chk_delete.Size = new System.Drawing.Size(97, 19);
            this.chk_delete.TabIndex = 91;
            this.chk_delete.Text = "Delete DATA";
            this.chk_delete.UseVisualStyleBackColor = true;
            this.chk_delete.CheckedChanged += new System.EventHandler(this.chk_delete_CheckedChanged);
            // 
            // Payment_Report
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(610, 238);
            this.Controls.Add(this.chk_delete);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.dateTimePicker1);
            this.Controls.Add(this.btn_Delete);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btn_Import);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btn_Emplacement);
            this.Controls.Add(this.txt_Emplacement);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Payment_Report";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Import Payment Report";
            this.Load += new System.EventHandler(this.Payment_Report_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Import;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btn_Emplacement;
        private System.Windows.Forms.TextBox txt_Emplacement;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btn_Delete;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.CheckBox chk_delete;
    }
}