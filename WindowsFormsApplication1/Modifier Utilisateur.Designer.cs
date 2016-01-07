namespace HarmonizedApp
{
    partial class Modifier_Utilisateur
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Modifier_Utilisateur));
            this.btn_Initialiser = new System.Windows.Forms.Button();
            this.btn_Insertion = new System.Windows.Forms.Button();
            this.txt_Pass = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txt_LOGIN = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_Initialiser
            // 
            this.btn_Initialiser.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btn_Initialiser.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Initialiser.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Initialiser.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Initialiser.Location = new System.Drawing.Point(222, 111);
            this.btn_Initialiser.Name = "btn_Initialiser";
            this.btn_Initialiser.Size = new System.Drawing.Size(110, 31);
            this.btn_Initialiser.TabIndex = 82;
            this.btn_Initialiser.Text = "Initialize";
            this.btn_Initialiser.UseVisualStyleBackColor = true;
            // 
            // btn_Insertion
            // 
            this.btn_Insertion.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btn_Insertion.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Insertion.ForeColor = System.Drawing.Color.DodgerBlue;
            this.btn_Insertion.Location = new System.Drawing.Point(56, 111);
            this.btn_Insertion.Name = "btn_Insertion";
            this.btn_Insertion.Size = new System.Drawing.Size(110, 31);
            this.btn_Insertion.TabIndex = 81;
            this.btn_Insertion.Text = "Save";
            this.btn_Insertion.UseVisualStyleBackColor = true;
            this.btn_Insertion.Click += new System.EventHandler(this.btn_Insertion_Click);
            // 
            // txt_Pass
            // 
            this.txt_Pass.Location = new System.Drawing.Point(167, 57);
            this.txt_Pass.Name = "txt_Pass";
            this.txt_Pass.PasswordChar = '*';
            this.txt_Pass.Size = new System.Drawing.Size(166, 20);
            this.txt_Pass.TabIndex = 79;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label3.Image = ((System.Drawing.Image)(resources.GetObject("label3.Image")));
            this.label3.Location = new System.Drawing.Point(53, 57);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(79, 17);
            this.label3.TabIndex = 78;
            this.label3.Text = "Password :";
            // 
            // txt_LOGIN
            // 
            this.txt_LOGIN.Location = new System.Drawing.Point(167, 23);
            this.txt_LOGIN.Name = "txt_LOGIN";
            this.txt_LOGIN.Size = new System.Drawing.Size(166, 20);
            this.txt_LOGIN.TabIndex = 77;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.DodgerBlue;
            this.label2.Image = ((System.Drawing.Image)(resources.GetObject("label2.Image")));
            this.label2.Location = new System.Drawing.Point(53, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(67, 17);
            this.label2.TabIndex = 76;
            this.label2.Text = "LOGIN :";
            // 
            // Modifier_Utilisateur
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("$this.BackgroundImage")));
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(398, 165);
            this.Controls.Add(this.btn_Initialiser);
            this.Controls.Add(this.btn_Insertion);
            this.Controls.Add(this.txt_Pass);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txt_LOGIN);
            this.Controls.Add(this.label2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Modifier_Utilisateur";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Edit User";
            this.Load += new System.EventHandler(this.Modifier_Utilisateur_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Initialiser;
        private System.Windows.Forms.Button btn_Insertion;
        private System.Windows.Forms.TextBox txt_Pass;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txt_LOGIN;
        private System.Windows.Forms.Label label2;
    }
}