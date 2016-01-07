using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HarmonizedApp.Manager
{
    class Personne
    {
        private string matricule;

        public string Matricule
        {
            get { return matricule; }
            set { matricule = value; }
        }
        private string nom;

        public string Nom
        {
            get { return nom; }
            set { nom = value; }
        }
        private string prenom;

        public string Prenom
        {
            get { return prenom; }
            set { prenom = value; }
        }
    }
}
