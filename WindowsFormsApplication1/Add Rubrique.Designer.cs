namespace HarmonizedApp
{
    partial class Add_Rubrique
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Add_Rubrique));
            this.btn_Initialiser = new System.Windows.Forms.Button();
            this.btn_Insertion = new System.Windows.Forms.Button();
            this.txt_NameRub = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_CodeRub = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_OrdreSeq = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btn_Initialiser
            // 
            this.btn_Initialiser.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btn_Initialiser.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Initialiser.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Initialiser.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Initialiser.Location = new System.Drawing.Point(214, 78);
            this.btn_Initialiser.Name = "btn_Initialiser";
            this.btn_Initialiser.Size = new System.Drawing.Size(110, 31);
            this.btn_Initialiser.TabIndex = 91;
            this.btn_Initialiser.Text = "Initialize";
            this.btn_Initialiser.UseVisualStyleBackColor = true;
            this.btn_Initialiser.Click += new System.EventHandler(this.btn_Initialiser_Click);
            // 
            // btn_Insertion
            // 
            this.btn_Insertion.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Insertion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Insertion.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Insertion.Location = new System.Drawing.Point(48, 78);
            this.btn_Insertion.Name = "btn_Insertion";
            this.btn_Insertion.Size = new System.Drawing.Size(110, 31);
            this.btn_Insertion.TabIndex = 90;
            this.btn_Insertion.Text = "Save";
            this.btn_Insertion.UseVisualStyleBackColor = true;
            this.btn_Insertion.Click += new System.EventHandler(this.btn_Insertion_Click);
            // 
            // txt_NameRub
            // 
            this.txt_NameRub.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_NameRub.Location = new System.Drawing.Point(188, 42);
            this.txt_NameRub.Name = "txt_NameRub";
            this.txt_NameRub.Size = new System.Drawing.Size(166, 21);
            this.txt_NameRub.TabIndex = 89;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label3.Image = ((System.Drawing.Image)(resources.GetObject("label3.Image")));
            this.label3.Location = new System.Drawing.Point(29, 43);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(103, 17);
            this.label3.TabIndex = 88;
            this.label3.Text = "Name Rubric :";
            // 
            // txt_CodeRub
            // 
            this.txt_CodeRub.Font = new System.Drawing.Font("Times New Roman", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txt_CodeRub.Location = new System.Drawing.Point(188, 8);
            this.txt_CodeRub.Name = "txt_CodeRub";
            this.txt_CodeRub.Size = new System.Drawing.Size(166, 21);
            this.txt_CodeRub.TabIndex = 87;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label2.Image = ((System.Drawing.Image)(resources.GetObject("label2.Image")));
            this.label2.Location = new System.Drawing.Point(29, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 17);
            this.label2.TabIndex = 86;
            this.label2.Text = "Code Rubric :";
            // 
            // txt_OrdreSeq
            // 
            this.txt_OrdreSeq.Location = new System.Drawing.Point(331, 95);
            this.txt_OrdreSeq.Name = "txt_OrdreSeq";
            this.txt_OrdreSeq.Size = new System.Drawing.Size(100, 20);
            this.txt_OrdreSeq.TabIndex = 92;
            this.txt_OrdreSeq.Visible = false;
            // 
            // Add_Rubrique
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(382, 127);
            this.Controls.Add(this.txt_OrdreSeq);
            this.Controls.Add(this.btn_Initialiser);
            this.Controls.Add(this.btn_Insertion);
            this.Controls.Add(this.txt_NameRub);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txt_CodeRub);
            this.Controls.Add(this.label2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Add_Rubrique";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Add Rubrique";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Initialiser;
        private System.Windows.Forms.Button btn_Insertion;
        private System.Windows.Forms.TextBox txt_NameRub;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_CodeRub;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_OrdreSeq;
    }
}