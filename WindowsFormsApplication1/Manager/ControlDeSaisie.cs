using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace HarmonizedApp.Manager
{
    class ControlDeSaisie
    {
         public void ControleNumérique(object sender, KeyPressEventArgs e, string Valeur)
        {
            string s = "";
            int i;
            e.Handled = (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar));
            if (e.Handled)
            {
                MessageBox.Show("Chiffre Invalide.",
      "Important Note",
      MessageBoxButtons.OK,
      MessageBoxIcon.Exclamation,
      MessageBoxDefaultButton.Button1);
            }
            else
            {
                if (!char.IsDigit(e.KeyChar))
                {
                    s = Valeur + e.KeyChar.ToString();
                }
                int.TryParse(s, out i);
                e.Handled = (i > 255);
            }
        }
        public double Valeur(string val)
        {
            double Valeur = 0.0;
            try
            {
                Valeur = double.Parse(val);
            }
            catch (Exception)
            {
                Valeur = 0.0;
            }
            return Valeur;
        }

        public void ControleDate(object sender, KeyPressEventArgs e, string Valeur)
        {
            e.Handled = (!Char.IsDigit(e.KeyChar) && (e.KeyChar != 58) && !char.IsControl(e.KeyChar));
            if (e.Handled)
            {
                MessageBox.Show("Caractère non valide", "Erreur Saisie", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        public void ControleTexte(object sender, KeyPressEventArgs e,string Valeur)
        {
            e.Handled = (!Char.IsLetter(e.KeyChar) && (e.KeyChar != 32) && !char.IsControl(e.KeyChar));
            if (e.Handled)
            {
                MessageBox.Show("Caractère non valide", "Erreur Saisie", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        public void ControleDeDouble(object sender, KeyPressEventArgs e, string Valeur)
        {
            string s = "";
            int i;
            e.Handled = (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar)&&(e.KeyChar !=46));
            if (e.Handled)
            {
                MessageBox.Show("Chiffre Invalide.",
      "Important Note",
      MessageBoxButtons.OK,
      MessageBoxIcon.Exclamation,
      MessageBoxDefaultButton.Button1);
            }
            else
            {
                if (!char.IsDigit(e.KeyChar))
                {
                    s = Valeur + e.KeyChar.ToString();
                }
                int.TryParse(s, out i);
                e.Handled = (i > 255);
            }
        }

    }
}
