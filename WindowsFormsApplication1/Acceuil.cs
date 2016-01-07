using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using HarmonizedApp.Manager;

namespace HarmonizedApp
{
    public partial class Acceuil : Form
    {
        ServiceDAO service = new ServiceDAO();
        private Liste_Des_Utilisateurs user = null;
        public string User = string.Empty;
        private LesUtilisateurs ajout = null;
        private Parameter_Extraction_PDF param = null;
        private Structure struc = null;
        private Code_Memo code = null;
        private Add_Rubrique add = null;
        private Parameter_Company company = null;
        private Consult_Company cons = null;
        private Import_Data_Fix fix = null;
        private SRF srf = null;
        private SVE sve = null;
        private Synchronisation_Code_Memo syn = null;
        Import_of_Gross_to_Net imp = null;
        List_Of_Rubrique lst = null;
        Gross_To_Net gross = null;
        SVR svr = null;
        Payment_Report pay = null;
        Pay_Report sep = null;
        public int idSociete = 0;
        public string CodeSoc = string.Empty;
        public string NomSoc = string.Empty;
        Account_Plan_about_GL_File gl = null;
        List_GL_File GLFile = null;
        Extract_SGL_File sgl = null;
        OPR8 opr8 = null;
        OPR11 opr11 = null;
        SGR sgr = null;
        Import_Rubrique imprub = null;
        OPR17 opr17 = null;
        OPR18 opr18 = null;
        OPR15 opr15 = null;

        public Acceuil()
        {
            InitializeComponent();
        }

        private void CutToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        public void AfficherUser()
        {
            toolStripStatusLabel4.Text = User;
            toolStripStatusLabel4.ForeColor = Color.Red;
        }

        private void StatusBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            statusStrip.Visible = statusBarToolStripMenuItem.Checked;
        }

        private void CascadeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }

        private void TileVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileVertical);
        }

        private void TileHorizontalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void ArrangeIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.ArrangeIcons);
        }

        private void CloseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        private void gérerLesUtilisateursToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (user == null || user.IsDisposed)
            {
                user = new Liste_Des_Utilisateurs();
                user.MdiParent = this;
                user.Show();
            }
            else
            {
                user.WindowState = FormWindowState.Maximized;
                user.Activate();
            }
        }

        private void ShowNewForm(object sender, EventArgs e)
        {
            Identification childForm = new Identification();
            childForm.Owner = this;
            childForm.Text = "Identification ";
            childForm.Souris();
            childForm.Show();
        }

        private void Acceuil_Load(object sender, EventArgs e)
        {
            ShowNewForm(sender, e);
            MPanel.Width = 180;
            GMenu1.Dock = DockStyle.Top;
            GMenu2.Dock = DockStyle.Bottom;
            MPanel1.Dock = DockStyle.Fill;
            MPanel1.Visible = true;
            MPanel2.Visible = false;  
        }

        public void AfficherApropos()
        {
            gestionDesUtilisateursToolStripMenuItem.Visible = true;
            managementOfCompanyToolStripMenuItem.Visible = true;
            managementToolStripMenuItem.Visible = true;
            gLFileConfigurationToolStripMenuItem.Visible = true;
        }

        public void Disable()
        {
            fileMenu.Enabled = true;
            windowsMenu.Enabled = true;
            GMenu1.Enabled = true;
            Btt_employers.Enabled = true;
            btt_affect.Enabled = true;
            brFrs.Enabled = true;
            Btt_Pointage.Enabled = true;
            Button1.Enabled = true;
            GMenu2.Enabled = true;
            Button5.Enabled = true;
            btt_affect.Enabled = true;
            importDATAToolStripMenuItem.Enabled = true;
            bChantier.Enabled = true;
            bFacReg.Enabled = true;
            button2.Enabled = true;
            oPRBGToolStripMenuItem.Enabled = true;
        }

        private void toolBarToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        public void AfficherUser(string Login,string NomSoc)
        {
            toolStripStatusLabel7.Text =  "                                           L'utilisateur est :"; 
            toolStripStatusLabel8.Text =  Login; 
            string DateSys = DateTime.Today.ToString().Substring(0, 10);
            toolStripStatusLabel9.Text = service.ConcChaine("x", 100, DateSys);
            toolStripStatusLabel10.Text = service.ConcChaine("x", 100, " La Société est :"); 
            toolStripStatusLabel11.Text = NomSoc;
            toolStripStatusLabel5.Visible = true;
            fileMenu.Enabled = true;
            articleToolStripMenuItem.Enabled = true;
            editMenu.Enabled = true;
            GMenu1.Enabled = true;
            Btt_employers.Enabled = true;
            Btt_Pointage.Enabled = true;
            button2.Enabled = true;

        }

        private void demandeDeCongéUtilisateurToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }

        private void Acceuil_Click(object sender, EventArgs e)
        {
        }

        private void consulterClientToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sve == null || sve.IsDisposed)
            {
                sve = new SVE();
                sve.MdiParent = this;
                sve.codeSoc = CodeSoc;
                sve.idSoc = idSociete;
                sve.NomSoc = NomSoc;
                sve.Show();
            }
            else
            {
                sve.WindowState = FormWindowState.Maximized;
                sve.Activate();
            }
        }
        private void ajouterClientToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void ajouterUtilisateurToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LesUtilisateurs users = new LesUtilisateurs();
            users.Show();
        }

        private void ajouterArticleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (struc == null || struc.IsDisposed)
            {
                struc = new Structure();
                struc.MdiParent = this;
                struc.Show();
            }
            else
            {
                struc.WindowState = FormWindowState.Maximized;
                struc.Activate();
            }
        }

        private void consulterArticleToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void lesEntréesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sgl == null || sgl.IsDisposed)
            {
                sgl = new Extract_SGL_File();
                sgl.idSoc = idSociete;
                sgl.codeSoc = CodeSoc;
                sgl.NomSoc = NomSoc;
                sgl.MdiParent = this;
                sgl.Show();
            }
            else
            {
                sgl.WindowState = FormWindowState.Maximized;
                sgl.Activate();
            }
        }

        private void lesSortiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sgr == null || sgr.IsDisposed)
            {
                sgr = new SGR();
                sgr.idSoc = idSociete;
                sgr.codeSoc = CodeSoc;
                sgr.NomSoc = NomSoc;
                sgr.MdiParent = this;
                sgr.Show();
            }
            else
            {
                sgr.WindowState = FormWindowState.Maximized;
                sgr.Activate();
            }
        }

        private void historiqueDeVenteToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void GMenu1_Click(object sender, EventArgs e)
        {

        }

        private void MPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void toolTip1_Popup(object sender, PopupEventArgs e)
        {

        }

        private void toolBarToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (toolBarToolStripMenuItem.Checked)
            {
                MPanel.Visible = true;
            }
            else
            {
                MPanel.Visible = false;
            }
        }

        private void Btt_employers_Click(object sender, EventArgs e)
        {

        }

        private void Btt_Pointage_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void statusStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void GMenu1_Click_1(object sender, EventArgs e)
        {
            GMenu1.Dock = DockStyle.Top;
            GMenu2.Dock = DockStyle.Bottom;
            MPanel1.Dock = DockStyle.Fill;
            MPanel1.Visible = true;
            MPanel2.Visible = false;
        }

        private void GMenu2_Click(object sender, EventArgs e)
        {
            GMenu1.Dock = DockStyle.Top;
            GMenu2.Dock = DockStyle.Top;
            MPanel2.Dock = DockStyle.Fill;
            MPanel1.Visible = false;
            MPanel2.Visible = true;
        }

        private void addNewUserToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (ajout == null || ajout.IsDisposed)
            {
                ajout = new LesUtilisateurs();
                ajout.MdiParent = this;
                ajout.Show();
            }
            else
            {
                ajout.WindowState = FormWindowState.Maximized;
                ajout.Activate();
            }
        }

        private void consultUsersToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (user == null || user.IsDisposed)
            {
                user = new Liste_Des_Utilisateurs();
                user.MdiParent = this;
                user.Show();
            }
            else
            {
                user.WindowState = FormWindowState.Maximized;
                user.Activate();
            }
        }

        private void extractingPDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (param == null || param.IsDisposed)
            {
                param = new Parameter_Extraction_PDF();
                param.MdiParent = this;
                param.Show();
            }
            else
            {
                param.WindowState = FormWindowState.Maximized;
                param.Activate();
            }
        }

        private void addNewCompanyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (company == null || company.IsDisposed)
            {
                company = new Parameter_Company();
                company.MdiParent = this;
                company.Show();
            }
            else
            {
                company.WindowState = FormWindowState.Maximized;
                company.Activate();
            }
        }

        private void consultCompanyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (cons == null || cons.IsDisposed)
            {
                cons = new Consult_Company();
                cons.MdiParent = this;
                cons.Show();
            }
            else
            {
                cons.WindowState = FormWindowState.Maximized;
                cons.Activate();
            }
        }

        private void menuStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void summaryReportFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (srf == null || srf.IsDisposed)
            {
                srf = new SRF();
                srf.MdiParent = this;
                srf.idSoc = idSociete;
                srf.NomSoc = NomSoc;
                srf.CodeSoc = CodeSoc;
                srf.Show();
            }
            else
            {
                srf.WindowState = FormWindowState.Maximized;
                srf.Activate();
            }
        }

        private void importGrossToNetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (imp == null || imp.IsDisposed)
            {
                imp = new Import_of_Gross_to_Net();
                imp.MdiParent = this;
                imp.CodeSoc = CodeSoc;
                imp.RefSoc = idSociete;
                imp.NomSoc = NomSoc;
                imp.Show();
            }
            else
            {
                imp.WindowState = FormWindowState.Maximized;
                imp.Activate();
            }
        }

        private void addRubriqueToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (add == null || add.IsDisposed)
            {
                add = new Add_Rubrique();
                add.MdiParent = this;
                add.idSoc = idSociete;
                add.Show();
            }
            else
            {
                add.WindowState = FormWindowState.Maximized;
                add.Activate();
            }
        }

        private void addRubriqueToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (add == null || add.IsDisposed)
            {
                add = new Add_Rubrique();
                add.MdiParent = this;
                add.idSoc = idSociete;
                add.Show();
            }
            else
            {
                add.WindowState = FormWindowState.Maximized;
                add.Activate();
            }
        }

        private void addRubriqueToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (add == null || add.IsDisposed)
            {
                add = new Add_Rubrique();
                add.MdiParent = this;
                add.idSoc = idSociete;
                add.Show();
            }
            else
            {
                add.WindowState = FormWindowState.Maximized;
                add.Activate();
            }
        }

        private void consultRubriqueToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lst == null || lst.IsDisposed)
            {
                lst = new List_Of_Rubrique();
                lst.MdiParent = this;
                lst.idSoc = idSociete;
                lst.Show();
            }
            else
            {
                lst.WindowState = FormWindowState.Maximized;
                lst.Activate();
            }
        }

        private void summaryVarienceReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (svr == null || svr.IsDisposed)
            {
                svr = new SVR();
                svr.MdiParent = this;
                svr.idSoc = idSociete;
                svr.codeSoc = CodeSoc;
                svr.NomSoc = NomSoc;
                svr.Show();
            }
            else
            {
                lst.WindowState = FormWindowState.Maximized;
                lst.Activate();
            }
        }

        private void grossToNetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (gross == null || gross.IsDisposed)
            {
                gross = new Gross_To_Net();
                gross.MdiParent = this;
                gross.idSoc = idSociete;
                gross.codeSoc = CodeSoc;
                gross.NomSoc = NomSoc;
                gross.Show();
            }
            else
            {
                gross.WindowState = FormWindowState.Maximized;
                gross.Activate();
            }
        }

        private void importPaymentReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (pay == null || pay.IsDisposed)
            {
                pay = new Payment_Report();
                pay.MdiParent = this;
                pay.CodeSoc = CodeSoc;
                pay.RefSoc = idSociete;
                pay.NomSoc = NomSoc;
                pay.Show();
            }
            else
            {
                pay.WindowState = FormWindowState.Maximized;
                pay.Activate();
            }
        }

        private void paymentReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sep == null || sep.IsDisposed)
            {
                sep = new Pay_Report();
                sep.MdiParent = this;
                sep.codeSoc = CodeSoc;
                sep.idSoc = idSociete;
                sep.NomSoc = NomSoc;
                sep.Show();
            }
            else
            {
                sep.WindowState = FormWindowState.Maximized;
                sep.Activate();
            }
        }

        private void importDataFixeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (fix == null || fix.IsDisposed)
            {
                fix = new Import_Data_Fix();
                fix.MdiParent = this;
                fix.CodeSoc = CodeSoc;
                fix.RefSoc = idSociete;
                fix.NomSoc = NomSoc;
                fix.Show();
            }
            else
            {
                fix.WindowState = FormWindowState.Maximized;
                fix.Activate();
            }
        }

        private void affectationCodeMemoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (code == null || code.IsDisposed)
            {
                code = new Code_Memo();
                code.MdiParent = this;
                code.idSoc = idSociete;
                code.Show();
            }
            else
            {
                code.WindowState = FormWindowState.Maximized;
                code.Activate();
            }
        }

        private void synchronisationCodeMemoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (syn == null || syn.IsDisposed)
            {
                syn = new Synchronisation_Code_Memo();
                syn.MdiParent = this;
                syn.idSoc = idSociete;
                syn.Show();
            }
            else
            {
                syn.WindowState = FormWindowState.Maximized;
                syn.Activate();
            }
        }

        private void Button5_Click_1(object sender, EventArgs e)
        {
            if (gross == null || gross.IsDisposed)
            {
                gross = new Gross_To_Net();
                gross.MdiParent = this;
                gross.idSoc = idSociete;
                gross.codeSoc = CodeSoc;
                gross.NomSoc = NomSoc;
                gross.Show();
            }
            else
            {
                gross.WindowState = FormWindowState.Maximized;
                gross.Activate();
            }
        }

        private void btt_affect_Click(object sender, EventArgs e)
        {
            if (sep == null || sep.IsDisposed)
            {
                sep = new Pay_Report();
                sep.MdiParent = this;
                sep.codeSoc = CodeSoc;
                sep.idSoc = idSociete;
                sep.NomSoc = NomSoc;
                sep.Show();
            }
            else
            {
                sep.WindowState = FormWindowState.Maximized;
                sep.Activate();
            }
        }

        private void bChantier_Click(object sender, EventArgs e)
        {
            if (svr == null || svr.IsDisposed)
            {
                svr = new SVR();
                svr.MdiParent = this;
                svr.idSoc = idSociete;
                svr.codeSoc = CodeSoc;
                svr.NomSoc = NomSoc;
                svr.Show();
            }
            else
            {
                lst.WindowState = FormWindowState.Maximized;
                lst.Activate();
            }
        }

        private void bFacReg_Click(object sender, EventArgs e)
        {
            if (sve == null || sve.IsDisposed)
            {
                sve = new SVE();
                sve.MdiParent = this;
                sve.codeSoc = CodeSoc;
                sve.idSoc = idSociete;
                sve.NomSoc = NomSoc;
                sve.Show();
            }
            else
            {
                sve.WindowState = FormWindowState.Maximized;
                sve.Activate();
            }
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (param == null || param.IsDisposed)
            {
                param = new Parameter_Extraction_PDF();
                param.MdiParent = this;
                param.Show();
            }
            else
            {
                param.WindowState = FormWindowState.Maximized;
                param.Activate();
            }
        }

        private void addCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (gl == null || gl.IsDisposed)
            {
                gl = new Account_Plan_about_GL_File();
                gl.MdiParent = this;
                gl.IdSoc = idSociete;
                gl.Show();
            }
            else
            {
                gl.WindowState = FormWindowState.Maximized;
                gl.Activate();
            }
        }

        private void consultConfigurationToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (GLFile == null || GLFile.IsDisposed)
            {
                GLFile = new List_GL_File();
                GLFile.MdiParent = this;
                GLFile.idSoc = idSociete;
                GLFile.Show();
            }
            else
            {
                GLFile.WindowState = FormWindowState.Maximized;
                GLFile.Activate();
            }
        }

        private void oPR8ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (opr8 == null || opr8.IsDisposed)
            {
                opr8 = new OPR8();
                opr8.MdiParent = this;
                opr8.NomSoc = NomSoc;
                opr8.Show();
            }
            else
            {
                opr8.WindowState = FormWindowState.Maximized;
                opr8.Activate();
            }
        }

        private void oPR11ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (opr11 == null || opr11.IsDisposed)
            {
                opr11 = new OPR11();
                opr11.MdiParent = this;
                opr11.NomSoc = NomSoc;
                opr11.Show();
            }
            else
            {
                opr11.WindowState = FormWindowState.Maximized;
                opr11.Activate();
            }
        }

        private void importRubriqueToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (imprub == null || imprub.IsDisposed)
            {
                imprub = new Import_Rubrique();
                imprub.MdiParent = this;
                imprub.Show();
            }
            else
            {
                imprub.WindowState = FormWindowState.Maximized;
                imprub.Activate();
            }
        }

        private void oPR17ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (opr17 == null || opr17.IsDisposed)
            {
                opr17 = new OPR17();
                opr17.MdiParent = this;
                opr17.NomSoc = NomSoc;
                opr17.Show();
            }
            else
            {
                opr17.WindowState = FormWindowState.Maximized;
                opr17.Activate();
            }
        }

        private void oPR18ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (opr18 == null || opr18.IsDisposed)
            {
                opr18 = new OPR18();
                opr18.MdiParent = this;
                opr18.NomSoc = NomSoc;
                opr18.Show();
            }
            else
            {
                opr18.WindowState = FormWindowState.Maximized;
                opr18.Activate();
            }
        }

        private void oPR15ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (opr15 == null || opr15.IsDisposed)
            {
                opr15 = new OPR15();
                opr15.MdiParent = this;
                opr15.NomSoc = NomSoc;
                opr15.Show();
            }
            else
            {
                opr15.WindowState = FormWindowState.Maximized;
                opr15.Activate();
            }
        }

        private void rubricsStructureToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmMappingRubric f = new frmMappingRubric();
            f.ShowDialog();
        }

    }
}
