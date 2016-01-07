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
    public partial class List_Of_Rubrique : Form
    {
        public int idSoc = 0;
        ServiceDAO service = new ServiceDAO();

        public List_Of_Rubrique()
        {
            InitializeComponent();
        }

        private void List_Of_Rubrique_Load(object sender, EventArgs e)
        {
            RemplirGrid();
        }

        public void ViderGrid()
        {
            dataGridView1.Rows.Clear();
        }

        public void RemplirGrid()
        {
            OleDbDataReader dr = service.AffichageListeRubrique(idSoc);
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(dr);

            foreach (System.Data.DataRow row in dt.Rows)
            {
                dataGridView1.Rows.Add(row["Num"].ToString(), row["CodeRub"].ToString(),
                row["Libelles"].ToString());

            }

        }

        private void btn_Supprimer_Click(object sender, EventArgs e)
        {
            int i = dataGridView1.CurrentRow.Index;
            if (MessageBox.Show("Would you like to delete the rubric : " + dataGridView1.Rows[i].Cells[1].Value.ToString(), "Attention", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                if (service.Delete_Rubric(int.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString())) == true)
                {
                    MessageBox.Show("Delete realised with success");
                    dataGridView1.Rows.Clear();
                    RemplirGrid();
                }
                else
                {
                    MessageBox.Show(" Error while delete this rubric ", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btn_Modifier_Click(object sender, EventArgs e)
        {
            Edit_Rubric modif = new Edit_Rubric();
            int i = dataGridView1.CurrentRow.Index;
            modif.Refer = int.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
            modif.CodeRub = dataGridView1.Rows[i].Cells[1].Value.ToString();
            modif.LibellesRub = dataGridView1.Rows[i].Cells[2].Value.ToString();
            modif.Show();
        }

    }
}
