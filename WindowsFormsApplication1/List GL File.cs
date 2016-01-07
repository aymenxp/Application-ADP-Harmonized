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
    public partial class List_GL_File : Form
    {
        public List_GL_File()
        {
            InitializeComponent();
        }

        ServiceDAO service = new ServiceDAO();
        public int idSoc;

        private void List_GL_File_Load(object sender, EventArgs e)
        {
            RemplirGrid();
        }

        public void RemplirGrid()
        {
            OleDbDataReader dr = service.AffichageListeGLFILE(idSoc);
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Load(dr);

            foreach (System.Data.DataRow row in dt.Rows)
            {
                dataGridView1.Rows.Add(row["NUM"].ToString(), row["CodeRub"].ToString(), row["LibCodeRub"].ToString(),
                row["CodeCpt"].ToString(), row["Libelles"].ToString(), row["Sens"].ToString());

            }

        }

        private void btn_Supprimer_Click(object sender, EventArgs e)
        {
            int i = dataGridView1.CurrentRow.Index;
            if (MessageBox.Show("Would you like to delete the rubric : " + dataGridView1.Rows[i].Cells[1].Value.ToString(), "Attention", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                if (service.Delete_GLFILE(int.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString())) == true)
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
            Update_Data_GL_File modif = new Update_Data_GL_File();
            int i = dataGridView1.CurrentRow.Index;
            modif.Refer = int.Parse(dataGridView1.Rows[i].Cells[0].Value.ToString());
            modif.CodeRub = dataGridView1.Rows[i].Cells[1].Value.ToString();
            modif.LibRub = dataGridView1.Rows[i].Cells[2].Value.ToString();
            modif.CodeCpt = dataGridView1.Rows[i].Cells[3].Value.ToString();
            modif.Libelles = dataGridView1.Rows[i].Cells[4].Value.ToString();
            modif.Sens = dataGridView1.Rows[i].Cells[5].Value.ToString();
            modif.Show();
        }

        public void ViderGrid()
        {
            dataGridView1.Rows.Clear();
        }

    }
}
