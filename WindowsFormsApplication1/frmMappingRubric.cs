using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace HarmonizedApp
{
    public partial class frmMappingRubric : Form
    {
        public frmMappingRubric()
        {
            InitializeComponent();
        }

        private void frmMappingRubric_Load(object sender, EventArgs e)
        {
            // TODO : cette ligne de code charge les données dans la table 'applicationADPDataSet.RubriqueStructureGrossToNet'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
          //  this.rubriqueStructureGrossToNetTableAdapter.Fill(this.applicationADPDataSet.RubriqueStructureGrossToNet);
            // TODO : cette ligne de code charge les données dans la table 'applicationADPDataSet1.RubriqueStructureGrossToNet'. Vous pouvez la déplacer ou la supprimer selon vos besoins.
        
        }

        private void btAddRubric_Click(object sender, EventArgs e)
        {
            frmNewRubric f = new frmNewRubric();
            f.ShowDialog();
        }

     
    }
}
