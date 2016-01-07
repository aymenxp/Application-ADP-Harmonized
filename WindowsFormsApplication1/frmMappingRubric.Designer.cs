namespace HarmonizedApp
{
    partial class frmMappingRubric
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
            this.components = new System.ComponentModel.Container();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.btAddRubric = new System.Windows.Forms.Button();
            this.btDetailsRub = new System.Windows.Forms.Button();
            this.btQuit = new System.Windows.Forms.Button();
            this.btDeleteRub = new System.Windows.Forms.Button();
            this.applicationADPDataSet = new HarmonizedApp.ApplicationADPDataSet();
            this.rubriqueStructureGrossToNetBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.rubriqueStructureGrossToNetTableAdapter = new HarmonizedApp.ApplicationADPDataSetTableAdapters.RubriqueStructureGrossToNetTableAdapter();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.applicationADPDataSet)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rubriqueStructureGrossToNetBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dataGridView1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(769, 322);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Liste des rubriques";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7});
            this.dataGridView1.DataSource = this.rubriqueStructureGrossToNetBindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(7, 20);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(752, 296);
            this.dataGridView1.TabIndex = 0;
            // 
            // btAddRubric
            // 
            this.btAddRubric.Location = new System.Drawing.Point(16, 344);
            this.btAddRubric.Name = "btAddRubric";
            this.btAddRubric.Size = new System.Drawing.Size(171, 36);
            this.btAddRubric.TabIndex = 1;
            this.btAddRubric.Text = "Ajouter rubrique";
            this.btAddRubric.UseVisualStyleBackColor = true;
            this.btAddRubric.Click += new System.EventHandler(this.btAddRubric_Click);
            // 
            // btDetailsRub
            // 
            this.btDetailsRub.Location = new System.Drawing.Point(193, 344);
            this.btDetailsRub.Name = "btDetailsRub";
            this.btDetailsRub.Size = new System.Drawing.Size(171, 36);
            this.btDetailsRub.TabIndex = 2;
            this.btDetailsRub.Text = "Détails rubrique";
            this.btDetailsRub.UseVisualStyleBackColor = true;
            // 
            // btQuit
            // 
            this.btQuit.Location = new System.Drawing.Point(547, 344);
            this.btQuit.Name = "btQuit";
            this.btQuit.Size = new System.Drawing.Size(145, 36);
            this.btQuit.TabIndex = 3;
            this.btQuit.Text = "Quitter";
            this.btQuit.UseVisualStyleBackColor = true;
            // 
            // btDeleteRub
            // 
            this.btDeleteRub.Location = new System.Drawing.Point(370, 344);
            this.btDeleteRub.Name = "btDeleteRub";
            this.btDeleteRub.Size = new System.Drawing.Size(171, 36);
            this.btDeleteRub.TabIndex = 4;
            this.btDeleteRub.Text = "Supprimer rubrique";
            this.btDeleteRub.UseVisualStyleBackColor = true;
            // 
            // applicationADPDataSet
            // 
            this.applicationADPDataSet.DataSetName = "ApplicationADPDataSet";
            this.applicationADPDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // rubriqueStructureGrossToNetBindingSource
            // 
            this.rubriqueStructureGrossToNetBindingSource.DataMember = "RubriqueStructureGrossToNet";
            this.rubriqueStructureGrossToNetBindingSource.DataSource = this.applicationADPDataSet;
            // 
            // rubriqueStructureGrossToNetTableAdapter
            // 
            this.rubriqueStructureGrossToNetTableAdapter.ClearBeforeFill = true;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "Label";
            this.dataGridViewTextBoxColumn1.HeaderText = "Label";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "ColonneType";
            this.dataGridViewTextBoxColumn2.HeaderText = "ColonneType";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "StartString";
            this.dataGridViewTextBoxColumn3.HeaderText = "StartString";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "EndString";
            this.dataGridViewTextBoxColumn4.HeaderText = "EndString";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "ConvertOut";
            this.dataGridViewTextBoxColumn5.HeaderText = "ConvertOut";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "ReplaceString";
            this.dataGridViewTextBoxColumn6.HeaderText = "ReplaceString";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "Order";
            this.dataGridViewTextBoxColumn7.HeaderText = "Order";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            // 
            // frmMappingRubric
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(790, 392);
            this.Controls.Add(this.btDeleteRub);
            this.Controls.Add(this.btQuit);
            this.Controls.Add(this.btDetailsRub);
            this.Controls.Add(this.btAddRubric);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmMappingRubric";
            this.Text = "Liste des rubriques (GrossToNet)";
            this.Load += new System.EventHandler(this.frmMappingRubric_Load);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.applicationADPDataSet)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rubriqueStructureGrossToNetBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btAddRubric;
        private System.Windows.Forms.Button btDetailsRub;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button btQuit;
        private System.Windows.Forms.Button btDeleteRub;

      
        private System.Windows.Forms.DataGridViewTextBoxColumn labelDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn colonneTypeDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn startStringDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn endStringDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn convertOutDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn replaceStringDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn orderDataGridViewTextBoxColumn;
        private ApplicationADPDataSet applicationADPDataSet;
        private System.Windows.Forms.BindingSource rubriqueStructureGrossToNetBindingSource;
        private HarmonizedApp.ApplicationADPDataSetTableAdapters.RubriqueStructureGrossToNetTableAdapter rubriqueStructureGrossToNetTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
    }
}