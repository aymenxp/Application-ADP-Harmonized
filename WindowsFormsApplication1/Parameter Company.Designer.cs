namespace HarmonizedApp
{
    partial class Parameter_Company
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Parameter_Company));
            this.btn_Initialiser = new System.Windows.Forms.Button();
            this.btn_Insertion = new System.Windows.Forms.Button();
            this.txt_CodeSoc = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_Soc = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_Initialiser
            // 
            this.btn_Initialiser.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btn_Initialiser.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Initialiser.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Initialiser.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Initialiser.Location = new System.Drawing.Point(215, 83);
            this.btn_Initialiser.Name = "btn_Initialiser";
            this.btn_Initialiser.Size = new System.Drawing.Size(110, 31);
            this.btn_Initialiser.TabIndex = 79;
            this.btn_Initialiser.Text = "Initialize";
            this.btn_Initialiser.UseVisualStyleBackColor = true;
            this.btn_Initialiser.Click += new System.EventHandler(this.btn_Initialiser_Click);
            // 
            // btn_Insertion
            // 
            this.btn_Insertion.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Insertion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Insertion.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Insertion.Location = new System.Drawing.Point(49, 83);
            this.btn_Insertion.Name = "btn_Insertion";
            this.btn_Insertion.Size = new System.Drawing.Size(110, 31);
            this.btn_Insertion.TabIndex = 78;
            this.btn_Insertion.Text = "Save";
            this.btn_Insertion.UseVisualStyleBackColor = true;
            this.btn_Insertion.Click += new System.EventHandler(this.btn_Insertion_Click);
            // 
            // txt_CodeSoc
            // 
            this.txt_CodeSoc.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_CodeSoc.Location = new System.Drawing.Point(189, 47);
            this.txt_CodeSoc.Name = "txt_CodeSoc";
            this.txt_CodeSoc.Size = new System.Drawing.Size(166, 21);
            this.txt_CodeSoc.TabIndex = 77;
            this.txt_CodeSoc.TextChanged += new System.EventHandler(this.txt_CodeSoc_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label3.Image = ((System.Drawing.Image)(resources.GetObject("label3.Image")));
            this.label3.Location = new System.Drawing.Point(30, 48);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(117, 17);
            this.label3.TabIndex = 76;
            this.label3.Text = "Code Company :";
            // 
            // txt_Soc
            // 
            this.txt_Soc.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_Soc.Location = new System.Drawing.Point(189, 13);
            this.txt_Soc.Name = "txt_Soc";
            this.txt_Soc.Size = new System.Drawing.Size(166, 21);
            this.txt_Soc.TabIndex = 75;
            this.txt_Soc.TextChanged += new System.EventHandler(this.txt_CompanyName_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label2.Image = ((System.Drawing.Image)(resources.GetObject("label2.Image")));
            this.label2.Location = new System.Drawing.Point(30, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 17);
            this.label2.TabIndex = 74;
            this.label2.Text = "Company Name :";
            // 
            // Parameter_Company
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(385, 156);
            this.Controls.Add(this.btn_Initialiser);
            this.Controls.Add(this.btn_Insertion);
            this.Controls.Add(this.txt_CodeSoc);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txt_Soc);
            this.Controls.Add(this.label2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Parameter_Company";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Parameter Company";
            this.Load += new System.EventHandler(this.Parameter_Company_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Initialiser;
        private System.Windows.Forms.Button btn_Insertion;
        private System.Windows.Forms.TextBox txt_CodeSoc;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_Soc;
        private System.Windows.Forms.Label label2;

    }
}