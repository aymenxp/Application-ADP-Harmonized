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
    public partial class Gross_To_Net : Form
    {

        public int idSoc;
        public string codeSoc = string.Empty;
        public string NomSoc = string.Empty;
        ExtractSGN sgn = new ExtractSGN();
        FormatDate form = new FormatDate();
        ServiceDAO serv = new ServiceDAO();

        public Gross_To_Net()
        {
            InitializeComponent();
        }

        private void btn_Gross_Click(object sender, EventArgs e)
        {
            string Date1 = form.LA_DATE(dateTimePicker1.Value.ToString());
            string Date2 = form.LA_DATE(dateTimePicker2.Value.ToString());
            int MoisPe = int.Parse(dateTimePicker1.Value.ToString().Substring(3, 2));
            int AnneePe = int.Parse(dateTimePicker1.Value.ToString().Substring(6, 4));
            sgn.ConstructSGN(Date1, Date2, codeSoc, NomSoc, progressBar1, idSoc, MoisPe,AnneePe);
            if (idSoc == 2)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA001").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA001").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA001").ToString();
            }
            if (idSoc == 5)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA002").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA002").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA002").ToString();
            }
            if (idSoc == 8)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA003").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA003").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA003").ToString();
            }
            if (idSoc == 7)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA004").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA004").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA004").ToString();
            }
            if (idSoc == 14)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA021").ToString();
                txt_NetPay.Text = serv.SumNetPayA021(MoisPe, AnneePe, "GROSSTONETA021").ToString();
                txt_SalBrut.Text = serv.SumBrutPayA021(MoisPe, AnneePe, "GROSSTONETA021").ToString();
            }
            if (idSoc == 9)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA010").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA010").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA010").ToString();
            }

            if (idSoc == 10)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA011").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA011").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA011").ToString();
            }
            if (idSoc == 3)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA014").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA014").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA014").ToString();
            }
            if (idSoc == 12)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA015").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA015").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA015").ToString();
            }
            if (idSoc == 11)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA013").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA013").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA013").ToString();
            }
            if (idSoc == 4)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA020").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA020").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA020").ToString();
            }
            if (idSoc == 13)
            {
                txt_TotalEmp.Text = serv.NumberSalary(MoisPe, AnneePe, "GROSSTONETA016").ToString();
                txt_NetPay.Text = serv.SumNetPay(MoisPe, AnneePe, "GROSSTONETA016").ToString();
                txt_SalBrut.Text = serv.SumBrutPay(MoisPe, AnneePe, "GROSSTONETA016").ToString();
            }

        }

        private void Gross_To_Net_Load(object sender, EventArgs e)
        {

        }
    }
}
