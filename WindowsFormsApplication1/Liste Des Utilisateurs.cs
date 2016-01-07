using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HarmonizedApp.Manager;
using System.Data.SqlClient;
using System.Data.OleDb;

namespace HarmonizedApp
{
    public partial class Liste_Des_Utilisateurs : Form
    {

        ServiceDAO service = new ServiceDAO();

        public Liste_Des_Utilisateurs()
        {
            InitializeComponent();
        }

        private void Liste_Des_Utilisateurs_Load(object sender, EventArgs e)
        {
            RemplirGrid();
        }

        public void RemplirGrid()
        {
            OleDbDataReader dr = service.AffichageListeUsers();
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(dr);

            foreach (System.Data.DataRow row in dt.Rows)
            {
                dataGridView1.Rows.Add(row["N°"].ToString(), row["Login"].ToString(),
                row["Pass"].ToString());
              
            }

        }

        private void btn_Modifier_Click(object sender, EventArgs e)
        {
            Modifier_Utilisateur modif = new Modifier_Utilisateur();
            int i = dataGridView1.CurrentRow.Index;
            modif.RefUser = int.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
            modif.Login = dataGridView1.Rows[i].Cells[1].Value.ToString();
            modif.Pass = dataGridView1.Rows[i].Cells[2].Value.ToString();           
            modif.Show();

        }

        private void btn_Supprimer_Click(object sender, EventArgs e)
        {
            int i = dataGridView1.CurrentRow.Index;
            if (MessageBox.Show("Would you like to delete the user : " + dataGridView1.Rows[i].Cells[1].Value.ToString(), "Attention", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                if (service.Delete_User(int.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString())) == true)
                {
                    MessageBox.Show("Delete realized with success");
                    dataGridView1.Rows.Clear();
                    RemplirGrid();
                }
                else
                {
                    MessageBox.Show(" Error while delete this user", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public void ViderGrid()
        {
            dataGridView1.Rows.Clear();
        }

        private void btn_Ajouter_Click(object sender, EventArgs e)
        {
           
        }

    }
}
