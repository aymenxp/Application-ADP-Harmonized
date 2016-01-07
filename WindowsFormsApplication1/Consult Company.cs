using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using HarmonizedApp.Manager;

namespace HarmonizedApp
{
    public partial class Consult_Company : Form
    {

        ServiceDAO service = new ServiceDAO();

        public Consult_Company()
        {
            InitializeComponent();
        }

        private void Consult_Company_Load(object sender, EventArgs e)
        {
            RemplirGrid();
        }

        public void RemplirGrid()
        {
            OleDbDataReader dr = service.AffichageListeCompany();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(dr);

            foreach (System.Data.DataRow row in dt.Rows)
            {
                dataGridView1.Rows.Add(row["Num"].ToString(), row["NomSoc"].ToString(),
                row["CodeSoc"].ToString());

            }

        }

        private void btn_Supprimer_Click(object sender, EventArgs e)
        {
            int i = dataGridView1.CurrentRow.Index;
            if (MessageBox.Show("Would you like to delete this company: " + dataGridView1.Rows[i].Cells[1].Value.ToString(), "Warning", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                if (service.Delete_Company(int.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString())) == true)
                {
                    MessageBox.Show("Deleting completed successfully");
                    dataGridView1.Rows.Clear();
                    RemplirGrid();
                }
                else
                {
                    MessageBox.Show(" Error of Delete", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btn_Modifier_Click(object sender, EventArgs e)
        {
            Edit_Company modif = new Edit_Company();
            int i = dataGridView1.CurrentRow.Index;
            modif.Refer = int.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
            modif.CompanyName = dataGridView1.Rows[i].Cells[1].Value.ToString();
            modif.CodeCompany = dataGridView1.Rows[i].Cells[2].Value.ToString();
            modif.Show();
        }

        public void ViderGrid()
        {
            dataGridView1.Rows.Clear();
        }

    }
}
