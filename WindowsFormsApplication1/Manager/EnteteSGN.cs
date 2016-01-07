using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using System.Data;

namespace HarmonizedApp.Manager
{
    class EnteteSGN
    {
        ServiceDAO service = new ServiceDAO();
        FormatDate form = new FormatDate();

        public Worksheet EnteteSGNLigne1(Worksheet worksheet)
        {
            worksheet.Cells[0, 0] = new Cell("Gross to Net");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne2(Worksheet worksheet)
        {
            worksheet.Cells[1, 0] = new Cell("Report Date and time");
            string Year = DateTime.Today.Year.ToString();
            string Month = DateTime.Today.Month.ToString();
            string Day = DateTime.Today.Day.ToString();
            if (Month.Length == 1)
            {
                Month = "0" + Month;
            }
            if (Day.Length == 1)
            {
                Day = "0" + Day;
            }

            string DATE1 = Year + "/" + Month + "/" + Day;
            worksheet.Cells[1, 1] = new Cell(DATE1);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne3(Worksheet worksheet)
        {
            worksheet.Cells[2, 0] = new Cell("Country code");
            worksheet.Cells[2, 1] = new Cell("TN");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne4(Worksheet worksheet,string NomSoc)
        {
            worksheet.Cells[3, 0] = new Cell("Company name");
            worksheet.Cells[3, 1] = new Cell(NomSoc);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne5(Worksheet worksheet, string CodeSoc)
        {
            worksheet.Cells[4, 0] = new Cell("Entity ID");
            worksheet.Cells[4, 1] = new Cell(CodeSoc);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne6(Worksheet worksheet, string NomSoc)
        {
            worksheet.Cells[5, 0] = new Cell("Entity name");
            worksheet.Cells[5, 1] = new Cell(NomSoc);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne7(Worksheet worksheet)
        {
            worksheet.Cells[6, 0] = new Cell("Currency");
            worksheet.Cells[6, 1] = new Cell("TND");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne8(Worksheet worksheet, string Date1, string Date2)
        {
            string DT1 = form.LA_DATE_ADP(Date1);
            string DT2 = form.LA_DATE_ADP(Date2);
            worksheet.Cells[7, 0] = new Cell("Pay Cycle");
            worksheet.Cells[7, 1] = new Cell(DT1 + "-" + DT2);
            return worksheet;
        }

        public Worksheet EnteteSVRLigne9(Worksheet worksheet)
        {
            worksheet.Cells[8, 0] = new Cell("breakdown per employee, per pay item with gross amount & deductions & net pay & employer deduction");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne10(Worksheet worksheet)
        {
            worksheet.Cells[9, 0] = new Cell("");
            worksheet.Cells[9, 1] = new Cell("");
            return worksheet;
        }

        public Worksheet EnteteSVRLigne11(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Identifiant du salarié");
            worksheet.Cells[10, 1] = new Cell("Nom");
            worksheet.Cells[10, 2] = new Cell("Prénom");
            worksheet.Cells[10, 3] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 4] = new Cell("Date de départ société");
            worksheet.Cells[10, 5] = new Cell("Code service");
            worksheet.Cells[10, 6] = new Cell("Intitulé service");
            worksheet.Cells[10, 7] = new Cell("Nature du contrat");
            worksheet.Cells[10, 8] = new Cell("Emploi occupé");
            worksheet.Cells[10, 9] = new Cell("Salaire de base");
            worksheet.Cells[10, 10] = new Cell("Nb de jour de présence");
            worksheet.Cells[10, 11] = new Cell("Salaires de base jours travaillés");
            worksheet.Cells[10, 12] = new Cell("Indemnité de transport");
            worksheet.Cells[10, 13] = new Cell("Indemnité de présence");
            worksheet.Cells[10, 14] = new Cell("Indemnité de fonction");
            worksheet.Cells[10, 15] = new Cell("Ind Manipulation de fond");
            worksheet.Cells[10, 16] = new Cell("Prime de panier");
            worksheet.Cells[10, 17] = new Cell("Prime de transport");
            worksheet.Cells[10, 18] = new Cell("Prime commercial CIP");
            worksheet.Cells[10, 19] = new Cell("Prime Award");
            worksheet.Cells[10, 20] = new Cell("Prime de voiture fixe");
            worksheet.Cells[10, 21] = new Cell("Ind forfait téléphonique");
            worksheet.Cells[10, 22] = new Cell("Prime astreinte réccurente");
            worksheet.Cells[10, 23] = new Cell("Prime de carburant");
            worksheet.Cells[10, 24] = new Cell("Heures supplémentaires");
            worksheet.Cells[10, 25] = new Cell("Congés payés");
            worksheet.Cells[10, 26] = new Cell("Prime Astreinte");
            worksheet.Cells[10, 27] = new Cell("Prime de risque");
            worksheet.Cells[10, 28] = new Cell("Prime fin d'année");
            worksheet.Cells[10, 29] = new Cell("Prime de naissance");
            worksheet.Cells[10, 30] = new Cell("Prime de mariage");
            worksheet.Cells[10, 31] = new Cell("Prime Aid");
            worksheet.Cells[10, 32] = new Cell("Regularisation Brut");
            worksheet.Cells[10, 33] = new Cell("Rappel sur salaire");
            worksheet.Cells[10, 34] = new Cell("Avantage en nature");
            worksheet.Cells[10, 35] = new Cell("Indemnité de préavis");
            worksheet.Cells[10, 36] = new Cell("Gratification fin de service");
            worksheet.Cells[10, 37] = new Cell("Indemnité de licenciement");
            worksheet.Cells[10, 38] = new Cell("Salaire Brut");
            worksheet.Cells[10, 39] = new Cell("Ajustement Cotisation");
            worksheet.Cells[10, 40] = new Cell("Salaire brut cotisable");
            worksheet.Cells[10, 41] = new Cell("CNSS employé");
            worksheet.Cells[10, 42] = new Cell("CNSS Complémentaire employé");
            worksheet.Cells[10, 43] = new Cell("Rtn assurance groupe maladie");
            worksheet.Cells[10, 44] = new Cell("Salaire Imposable");
            worksheet.Cells[10, 45] = new Cell("Rtn assurance groupe inc et décès");
            worksheet.Cells[10, 46] = new Cell("IRPP");
            worksheet.Cells[10, 47] = new Cell("Redevance 1%");
            worksheet.Cells[10, 48] = new Cell("Contribution Conjoncturelle");
            worksheet.Cells[10, 49] = new Cell("Avance sur salaire");
            worksheet.Cells[10, 50] = new Cell("Prêt Société");
            worksheet.Cells[10, 51] = new Cell("Pret CNSS");
            worksheet.Cells[10, 52] = new Cell("Pret Aid");
            worksheet.Cells[10, 53] = new Cell("Rtn Avantage en nature");
            worksheet.Cells[10, 54] = new Cell("Remboursement");
            worksheet.Cells[10, 55] = new Cell("Total Déduction salarié");
            worksheet.Cells[10, 56] = new Cell("Salaire net à payer");
            worksheet.Cells[10, 57] = new Cell("CNSS employeur");
            worksheet.Cells[10, 58] = new Cell("CNSS Complémentaire employeur");
            worksheet.Cells[10, 59] = new Cell("Accident de travail");
            worksheet.Cells[10, 60] = new Cell("Assurance groupe maladie employeur");
            worksheet.Cells[10, 61] = new Cell("Assurance grpe inc et décès employeur");
            worksheet.Cells[10, 62] = new Cell("TFP");
            worksheet.Cells[10, 63] = new Cell("FOPROLOS");
            worksheet.Cells[10, 64] = new Cell("Total cotisation patronale");
            worksheet.Cells[10, 65] = new Cell("Mois de paie");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A002(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Identifiant du salarié");
            worksheet.Cells[10, 1] = new Cell("Nom");
            worksheet.Cells[10, 2] = new Cell("Prénom");
            worksheet.Cells[10, 3] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 4] = new Cell("Date de départ société");
            worksheet.Cells[10, 5] = new Cell("Code service");
            worksheet.Cells[10, 6] = new Cell("Intitulé service");
            worksheet.Cells[10, 7] = new Cell("Nature du contrat");
            worksheet.Cells[10, 8] = new Cell("Emploi occupé");
            worksheet.Cells[10, 9] = new Cell("Salaire de base");
            worksheet.Cells[10, 10] = new Cell("Nb de jour de présence");
            worksheet.Cells[10, 11] = new Cell("Nb Heures de présence");
            worksheet.Cells[10, 12] = new Cell("Salaires de base jours travaillés");
            worksheet.Cells[10, 13] = new Cell("Indemnité de transport");
            worksheet.Cells[10, 14] = new Cell("Indemnité de présence");
            worksheet.Cells[10, 15] = new Cell("Congés payer");
            worksheet.Cells[10, 16] = new Cell("Congés spéciaux");
            worksheet.Cells[10, 17] = new Cell("Congés Absences non rémunérés");
            worksheet.Cells[10, 18] = new Cell("Prime exceptionnelle fixe");
            worksheet.Cells[10, 19] = new Cell("Congés à payer");
            worksheet.Cells[10, 20] = new Cell("Heures supplémentaires 125%");
            worksheet.Cells[10, 21] = new Cell("Heures supplémentaires 150%");
            worksheet.Cells[10, 22] = new Cell("Heures supplémentaires 200%");
            worksheet.Cells[10, 23] = new Cell("Prime Astreinte");
            worksheet.Cells[10, 24] = new Cell("Prime de dimanche");
            worksheet.Cells[10, 25] = new Cell("Prime Différentielle");
            worksheet.Cells[10, 26] = new Cell("Prime Différentielle de formation");
            worksheet.Cells[10, 27] = new Cell("Prime de parrainage");
            worksheet.Cells[10, 28] = new Cell("Prime de formation");
            worksheet.Cells[10, 29] = new Cell("Prime exceptionnelle");
            worksheet.Cells[10, 30] = new Cell("Prime sur Objectif");
            worksheet.Cells[10, 31] = new Cell("Prime de naissance");
            worksheet.Cells[10, 32] = new Cell("Prime de mariage");
            worksheet.Cells[10, 33] = new Cell("Prime de décès");
            worksheet.Cells[10, 34] = new Cell("Prime Aid");
            worksheet.Cells[10, 35] = new Cell("Rappel élément réccurent");
            worksheet.Cells[10, 36] = new Cell("Rappel élément non réccurent");
            worksheet.Cells[10, 37] = new Cell("Avantage en nature TR");
            worksheet.Cells[10, 38] = new Cell("Rappel sur avantage en nature TR");
            worksheet.Cells[10, 39] = new Cell("Avantage en nature Ass groupe");
            worksheet.Cells[10, 40] = new Cell("Avantage en nature ticket aid");
            worksheet.Cells[10, 41] = new Cell("Indemnité de préavis");
            worksheet.Cells[10, 42] = new Cell("Gratification fin de service");
            worksheet.Cells[10, 43] = new Cell("Indemnité de licenciement");
            worksheet.Cells[10, 44] = new Cell("Dommages et intérêts");
            worksheet.Cells[10, 45] = new Cell("Salaire Brut");
            worksheet.Cells[10, 46] = new Cell("Salaire brut cotisable");
            worksheet.Cells[10, 47] = new Cell("CNSS employé");
            worksheet.Cells[10, 48] = new Cell("Rtn assurance groupe employé");
            worksheet.Cells[10, 49] = new Cell("Salaire Imposable");
            worksheet.Cells[10, 50] = new Cell("IRPP");
            worksheet.Cells[10, 51] = new Cell("Redevance 1%");
            worksheet.Cells[10, 52] = new Cell("Contribution Conjoncturelle");
            worksheet.Cells[10, 53] = new Cell("Avance sur salaire");
            worksheet.Cells[10, 54] = new Cell("Prêt Société");
            worksheet.Cells[10, 55] = new Cell("Retenue Tickets Restaurant");
            worksheet.Cells[10, 56] = new Cell("Autes déductions");
            worksheet.Cells[10, 57] = new Cell("Déduction Avantage en nature TR");
            worksheet.Cells[10, 58] = new Cell("Déduction avantage en nature rappel TR");
            worksheet.Cells[10, 59] = new Cell("Déduction Avantage en nature ass groupe");
            worksheet.Cells[10, 60] = new Cell("Deduction avantage ticket Aid");
            worksheet.Cells[10, 61] = new Cell("Remboursement Frais Judiciaires");
            worksheet.Cells[10, 62] = new Cell("Total Déduction employé");
            worksheet.Cells[10, 63] = new Cell("Salaire net à payer");
            worksheet.Cells[10, 64] = new Cell("CNSS employeur");
            worksheet.Cells[10, 65] = new Cell("Accident de travail");
            worksheet.Cells[10, 66] = new Cell("Assurance groupe employeur");
            worksheet.Cells[10, 67] = new Cell("Total contribution patronale");
            worksheet.Cells[10, 68] = new Cell("Mois de paie");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A003(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Identifiant du salarié");
            worksheet.Cells[10, 1] = new Cell("Nom");
            worksheet.Cells[10, 2] = new Cell("Prénom");
            worksheet.Cells[10, 3] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 4] = new Cell("Date de départ société");
            worksheet.Cells[10, 5] = new Cell("Code service");
            worksheet.Cells[10, 6] = new Cell("Intitulé service");
            worksheet.Cells[10, 7] = new Cell("Nature du contrat");
            worksheet.Cells[10, 8] = new Cell("Emploi occupé");
            worksheet.Cells[10, 9] = new Cell("Salaire de base");
            worksheet.Cells[10, 10] = new Cell("Nb de jour de présence");
            worksheet.Cells[10, 11] = new Cell("Salaires de base jours travaillés");
            worksheet.Cells[10, 12] = new Cell("Prime de transport");
            worksheet.Cells[10, 13] = new Cell("Prime de présence");
            worksheet.Cells[10, 14] = new Cell("Prime d'assiduité");
            worksheet.Cells[10, 15] = new Cell("Prime de panier");
            worksheet.Cells[10, 16] = new Cell("Prime de douche");
            worksheet.Cells[10, 17] = new Cell("Prime de caisse");
            worksheet.Cells[10, 18] = new Cell("Prime spécifique");
            worksheet.Cells[10, 19] = new Cell("Augmentation Provisoire 2014");
            worksheet.Cells[10, 20] = new Cell("Prime de redevance de 1%");
            worksheet.Cells[10, 21] = new Cell("Prime Fin Année");
            worksheet.Cells[10, 22] = new Cell("Prime de Peréquation");
            worksheet.Cells[10, 23] = new Cell("Heures supplémentaires 175%");
            worksheet.Cells[10, 24] = new Cell("Heures supplémentaires 200%");
            worksheet.Cells[10, 25] = new Cell("Heures supplémentaires de nuit");
            worksheet.Cells[10, 26] = new Cell("Primes exceptionnelles");
            worksheet.Cells[10, 27] = new Cell("Prime d'inventaire");
            worksheet.Cells[10, 28] = new Cell("Prime de productivité");
            worksheet.Cells[10, 29] = new Cell("Prime de déclaration");
            worksheet.Cells[10, 30] = new Cell("Prime de vente");
            worksheet.Cells[10, 31] = new Cell("Commission de recouvrement");
            worksheet.Cells[10, 32] = new Cell("Prime de Taxi");
            worksheet.Cells[10, 33] = new Cell("Gratification fin Sce");
            worksheet.Cells[10, 34] = new Cell("Rappel");
            worksheet.Cells[10, 35] = new Cell("Avantage en nature TR");
            worksheet.Cells[10, 36] = new Cell("Congés payés");
            worksheet.Cells[10, 37] = new Cell("Salaire Brut");
            worksheet.Cells[10, 38] = new Cell("Ajustement Cotisation");
            worksheet.Cells[10, 39] = new Cell("Salaire brut cotisable");
            worksheet.Cells[10, 40] = new Cell("CNSS employé");
            worksheet.Cells[10, 41] = new Cell("CNSS Complémentaire employé");
            worksheet.Cells[10, 42] = new Cell("Salaire Imposable");
            worksheet.Cells[10, 43] = new Cell("IRPP");
            worksheet.Cells[10, 44] = new Cell("Redevance 1%");
            worksheet.Cells[10, 45] = new Cell("Contribution Conjoncturelle");
            worksheet.Cells[10, 46] = new Cell("Assurance groupe employé");
            worksheet.Cells[10, 47] = new Cell("Avance sur salaire");
            worksheet.Cells[10, 48] = new Cell("Retenue Ticket restaurant");
            worksheet.Cells[10, 49] = new Cell("Prêt Société");
            worksheet.Cells[10, 50] = new Cell("Pret CNSS");
            worksheet.Cells[10, 51] = new Cell("Pret Aid");
            worksheet.Cells[10, 52] = new Cell("Avantage en nature TR");
            worksheet.Cells[10, 53] = new Cell("Autres Déductions");
            worksheet.Cells[10, 54] = new Cell("Total déduction salariale");
            worksheet.Cells[10, 55] = new Cell("Salaire net à payer");
            worksheet.Cells[10, 56] = new Cell("CNSS employeur");
            worksheet.Cells[10, 57] = new Cell("CNSS Complémentaire employeur");
            worksheet.Cells[10, 58] = new Cell("Accident de travail");
            worksheet.Cells[10, 59] = new Cell("Assurance groupe employeur");
            worksheet.Cells[10, 60] = new Cell("Total contribution employeur");
            worksheet.Cells[10, 61] = new Cell("Mois de paie");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A004(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Identifiant du salarié");
            worksheet.Cells[10, 1] = new Cell("Nom");
            worksheet.Cells[10, 2] = new Cell("Prénom");
            worksheet.Cells[10, 3] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 4] = new Cell("Date de départ société");
            worksheet.Cells[10, 5] = new Cell("Code service");
            worksheet.Cells[10, 6] = new Cell("Intitulé service");
            worksheet.Cells[10, 7] = new Cell("Nature du contrat");
            worksheet.Cells[10, 8] = new Cell("Emploi occupé");
            worksheet.Cells[10, 9] = new Cell("Salaire de base");
            worksheet.Cells[10, 10] = new Cell("Nb de jour de présence");
            worksheet.Cells[10, 11] = new Cell("Salaires de base jours travaillés");
            worksheet.Cells[10, 12] = new Cell("Prime de transport");
            worksheet.Cells[10, 13] = new Cell("Prime de présence");
            worksheet.Cells[10, 14] = new Cell("Prime d'assiduité");
            worksheet.Cells[10, 15] = new Cell("Prime de panier");
            worksheet.Cells[10, 16] = new Cell("Prime de douche");
            worksheet.Cells[10, 17] = new Cell("Prime de caisse");
            worksheet.Cells[10, 18] = new Cell("Prime spécifique");
            worksheet.Cells[10, 19] = new Cell("Augmentation Provisoire 2014");
            worksheet.Cells[10, 20] = new Cell("Prime de redevance de 1%");
            worksheet.Cells[10, 21] = new Cell("Prime de Peréquation");
            worksheet.Cells[10, 22] = new Cell("Heures supplémentaires 175%");
            worksheet.Cells[10, 23] = new Cell("Heures supplémentaires 200%");
            worksheet.Cells[10, 24] = new Cell("Heures supplémentaires de nuit");
            worksheet.Cells[10, 25] = new Cell("Prime d'inventaire");
            worksheet.Cells[10, 26] = new Cell("Primes exceptionnelles");
            worksheet.Cells[10, 27] = new Cell("Prime de productivité");
            worksheet.Cells[10, 28] = new Cell("Prime de déclaration");
            worksheet.Cells[10, 39] = new Cell("Prime de vente");
            worksheet.Cells[10, 30] = new Cell("Commission de recouvrement");
            worksheet.Cells[10, 31] = new Cell("Prime de Taxi");
            worksheet.Cells[10, 32] = new Cell("Conges Payes");
            worksheet.Cells[10, 33] = new Cell("Prime Fin Annee");
            worksheet.Cells[10, 34] = new Cell("Indemnité de préavis");
            worksheet.Cells[10, 35] = new Cell("Gratification Fin Service");
            worksheet.Cells[10, 36] = new Cell("Avantage en nature TR");
            worksheet.Cells[10, 37] = new Cell("Salaire Brut");
            worksheet.Cells[10, 38] = new Cell("Ajustement Cotisation");
            worksheet.Cells[10, 39] = new Cell("Salaire brut cotisable");
            worksheet.Cells[10, 40] = new Cell("CNSS employé");
            worksheet.Cells[10, 41] = new Cell("CNSS Complémentaire employé");
            worksheet.Cells[10, 42] = new Cell("Salaire Imposable");
            worksheet.Cells[10, 43] = new Cell("IRPP");
            worksheet.Cells[10, 44] = new Cell("Redevance 1%");
            worksheet.Cells[10, 45] = new Cell("Contribution Conjoncturelle");
            worksheet.Cells[10, 46] = new Cell("Assurance groupe employé");
            worksheet.Cells[10, 47] = new Cell("Avance sur salaire");
            worksheet.Cells[10, 48] = new Cell("Retenue Ticket restaurant");
            worksheet.Cells[10, 49] = new Cell("Prêt Société");
            worksheet.Cells[10, 50] = new Cell("Pret CNSS");
            worksheet.Cells[10, 51] = new Cell("Pret Aid");
            worksheet.Cells[10, 52] = new Cell("Avantage en nature TR");
            worksheet.Cells[10, 53] = new Cell("Autres Déductions");
            worksheet.Cells[10, 54] = new Cell("Total déduction employé");
            worksheet.Cells[10, 55] = new Cell("Salaire net à payer");
            worksheet.Cells[10, 56] = new Cell("CNSS employeur");
            worksheet.Cells[10, 57] = new Cell("CNSS Complémentaire employeur");
            worksheet.Cells[10, 58] = new Cell("Accident de travail");
            worksheet.Cells[10, 59] = new Cell("Assurance groupe employeur");
            worksheet.Cells[10, 60] = new Cell("Total contribution employeur");
            worksheet.Cells[10, 61] = new Cell("Mois de paie");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A021(Worksheet worksheet)
        {

            worksheet.Cells[10, 0] = new Cell(" matricule");
            worksheet.Cells[10, 1] = new Cell("nom");
            worksheet.Cells[10, 2] = new Cell("prenom");
            worksheet.Cells[10, 3] = new Cell("Datedembauchesociete");
            worksheet.Cells[10, 4] = new Cell("Datededepartsociete");
            worksheet.Cells[10, 5] = new Cell("Intituledepartement");
            worksheet.Cells[10, 6] = new Cell("  CodeCostcenter");
            worksheet.Cells[10, 7] = new Cell("  Intitulecostcenter");
            worksheet.Cells[10, 8] = new Cell("  Intituledelanatureducontrat");
            worksheet.Cells[10, 9] = new Cell("  Emploioccupe");
            worksheet.Cells[10, 10] = new Cell("  Salairedebase");
            worksheet.Cells[10, 11] = new Cell("  Nbjoursdepresence");
            worksheet.Cells[10, 12] = new Cell("  SalairedebaseJtrvaille");
            worksheet.Cells[10, 13] = new Cell("  Inddepresence");
            worksheet.Cells[10, 14] = new Cell("  InddeTransport");
            worksheet.Cells[10, 15] = new Cell("  PrimeComplementairedePresence");
            worksheet.Cells[10, 16] = new Cell("  IndemnitedeVoiture");
            worksheet.Cells[10, 17] = new Cell("  IndFixederepresentation");
            worksheet.Cells[10, 18] = new Cell("  Indemnitedexpatriation");
            worksheet.Cells[10, 19] = new Cell("HS125HS0");
            worksheet.Cells[10, 20] = new Cell("ValeurHS0");
            worksheet.Cells[10, 21] = new Cell("HS1502");
            worksheet.Cells[10, 22] = new Cell(" ValeurHS02");
            worksheet.Cells[10, 23] = new Cell(" HS175HS06");
            worksheet.Cells[10, 24] = new Cell("ValeurHS06");
            worksheet.Cells[10, 25] = new Cell("HS200HS03");
            worksheet.Cells[10, 26] = new Cell("ValeurHS03");
            worksheet.Cells[10, 27] = new Cell("HsupplementairenuitHS04");
            worksheet.Cells[10, 28] = new Cell("valeurHsuplementairenuitHS04");
            worksheet.Cells[10, 29] = new Cell("primeLVP");
            worksheet.Cells[10, 30] = new Cell("PrimeSIP");
            worksheet.Cells[10, 31] = new Cell("PrimeAOS");
            worksheet.Cells[10, 32] = new Cell("Bonus");
            worksheet.Cells[10, 33] = new Cell("Compensationsupport");
            worksheet.Cells[10, 34] = new Cell("Rappelsursalaire");
            worksheet.Cells[10, 35] = new Cell("Rappelprimedetransport");
            worksheet.Cells[10, 36] = new Cell("PrimeperequationRC");
            worksheet.Cells[10, 37] = new Cell("PrimeperequationAV");
            worksheet.Cells[10, 38] = new Cell("PrimeAid");
            worksheet.Cells[10, 39] = new Cell("Primemariage");
            worksheet.Cells[10, 40] = new Cell("Primedenaissance");
            worksheet.Cells[10, 41] = new Cell("Congespayes");
            worksheet.Cells[10, 42] = new Cell("Indemnitedepreavis");
            worksheet.Cells[10, 43] = new Cell("Gratificationfindeservice");
            worksheet.Cells[10, 44] = new Cell("Indemnitedelicenciement");
            worksheet.Cells[10, 45] = new Cell("Avticketsrestaurant");
            worksheet.Cells[10, 46] = new Cell("AvAssurancemaladie");
            worksheet.Cells[10, 47] = new Cell("Avassuranceretraitecomplementaire");
            worksheet.Cells[10, 48] = new Cell("Avcarburant");
            worksheet.Cells[10, 49] = new Cell("Avvoiture");
            worksheet.Cells[10, 50] = new Cell("Avscolariteenfants");
            worksheet.Cells[10, 51] = new Cell("Avutiliteexpat");
            worksheet.Cells[10, 52] = new Cell("Avlogement");
            worksheet.Cells[10, 53] = new Cell("SalaireBrut");
            worksheet.Cells[10, 54] = new Cell("BrutEricsson");
            worksheet.Cells[10, 55] = new Cell("CNSSEmploye");
            worksheet.Cells[10, 56] = new Cell("AssuranceretraitecomplemCNRPSemploye");
            worksheet.Cells[10, 57] = new Cell("AssuranceretraiteComplementaire");
            worksheet.Cells[10, 58] = new Cell("Salaireimposable");
            worksheet.Cells[10, 59] = new Cell("IRPP");
            worksheet.Cells[10, 60] = new Cell("Redevance1");
            worksheet.Cells[10, 61] = new Cell("PrêtCNSS");
            worksheet.Cells[10, 62] = new Cell("AutresRetenuesursalaire");
            worksheet.Cells[10, 63] = new Cell("DeductionAvTR");
            worksheet.Cells[10, 64] = new Cell("DeductionAvassurancemaladie");
            worksheet.Cells[10, 65] = new Cell("DeductionAvassurancecomplementaire");
            worksheet.Cells[10, 66] = new Cell("DeductionAvVoiture");
            worksheet.Cells[10, 67] = new Cell("DeductionAvCarburant");
            worksheet.Cells[10, 68] = new Cell("DeductionAvloyer");
            worksheet.Cells[10, 69] = new Cell("DeductionAvIndExpat");
            worksheet.Cells[10, 70] = new Cell("Deductionutiliteexpat");
            worksheet.Cells[10, 71] = new Cell("NetàPayer");
            worksheet.Cells[10, 72] = new Cell("CNSSEmployeur");
            worksheet.Cells[10, 73] = new Cell("Accidentdetravail");
            worksheet.Cells[10, 74] = new Cell("AssurancemaladieEmployeur");
            worksheet.Cells[10, 75] = new Cell("AssuranceComplementaireEmployeur");
            worksheet.Cells[10, 76] = new Cell("AssuranceretraitcompleCNRPSemplye");
            worksheet.Cells[10, 77] = new Cell("TFP");
            worksheet.Cells[10, 78] = new Cell("FOPROLOS");
            worksheet.Cells[10, 79] = new Cell("TotalContributionEmployeur");
            worksheet.Cells[10, 80] = new Cell("Moisdepaie");

      

            return worksheet;
        }



        public Worksheet EnteteSVRLigne11A011(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Identifiant du salarié");
            worksheet.Cells[10, 1] = new Cell("Nom");
            worksheet.Cells[10, 2] = new Cell("Prénom");
            worksheet.Cells[10, 3] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 4] = new Cell("Date de départ société");
            worksheet.Cells[10, 5] = new Cell("Intitulé département");
            worksheet.Cells[10, 6] = new Cell("Intitulé service");
            worksheet.Cells[10, 7] = new Cell("Intitulé catégorie");
            worksheet.Cells[10, 8] = new Cell("Nature du contrat");
            worksheet.Cells[10, 9] = new Cell("Emploi occupé");
            worksheet.Cells[10, 10] = new Cell("Salaire de base du salarié");
            worksheet.Cells[10, 11] = new Cell("Présence");
            worksheet.Cells[10, 12] = new Cell("Salaire de base jours travaillés");
            worksheet.Cells[10, 13] = new Cell("Prime de Présence");
            worksheet.Cells[10, 14] = new Cell("Prime de Transport");
            worksheet.Cells[10, 15] = new Cell("Indemnité de fonction");
            worksheet.Cells[10, 16] = new Cell("Prime allocation voiture");
            worksheet.Cells[10, 17] = new Cell("Prime de Panier Restauration");
            worksheet.Cells[10, 18] = new Cell("Prime de Panier");
            worksheet.Cells[10, 19] = new Cell("Prime de Bienvenue");
            worksheet.Cells[10, 20] = new Cell("Oncall Duty");
            worksheet.Cells[10, 21] = new Cell("Prime de mission");
            worksheet.Cells[10, 22] = new Cell("Prime de productivité");
            worksheet.Cells[10, 23] = new Cell("Prime de rendement annuelle");
            worksheet.Cells[10, 24] = new Cell("Prime fin d'année");
            worksheet.Cells[10, 25] = new Cell("Prime Aid Idha");
            worksheet.Cells[10, 26] = new Cell("Indemnité de Scolarité");
            worksheet.Cells[10, 27] = new Cell("Heures Supp 125%");
            worksheet.Cells[10, 28] = new Cell("Heures Supp 150%");
            worksheet.Cells[10, 29] = new Cell("Heures Supp 200%");
            worksheet.Cells[10, 30] = new Cell("Heures Compensation Réunion Qualité");
            worksheet.Cells[10, 31] = new Cell("Gratification fin de Sce");
            worksheet.Cells[10, 32] = new Cell("Conges A Payer");
            worksheet.Cells[10, 33] = new Cell("Total Brut");
            worksheet.Cells[10, 34] = new Cell("CNSS");
            worksheet.Cells[10, 35] = new Cell("Salaire Imposable");
            worksheet.Cells[10, 36] = new Cell("IRPP");
            worksheet.Cells[10, 37] = new Cell("Redevance");
            worksheet.Cells[10, 38] = new Cell("Total Déduction employé");
            worksheet.Cells[10, 39] = new Cell("Net à Payer");
            worksheet.Cells[10, 40] = new Cell("CNSS Employeur");
            worksheet.Cells[10, 41] = new Cell("Accident de travail");
            worksheet.Cells[10, 42] = new Cell("TFP");
            worksheet.Cells[10, 43] = new Cell("FOPROLOS");
            worksheet.Cells[10, 44] = new Cell("Total cotisation Employeur");
            worksheet.Cells[10, 45] = new Cell("Total cost of labor");
            worksheet.Cells[10, 46] = new Cell("Mois de paie");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A014(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Id Oracle");
            worksheet.Cells[10, 1] = new Cell("Matricule");
            worksheet.Cells[10, 2] = new Cell("Nom");
            worksheet.Cells[10, 3] = new Cell("Prénom");
            worksheet.Cells[10, 4] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 5] = new Cell("Date de sortie établissement");
            worksheet.Cells[10, 6] = new Cell("EXPAT / LOCAL");
            worksheet.Cells[10, 7] = new Cell("Emploi Occupé");
            worksheet.Cells[10, 8] = new Cell("Intitulé département");
            worksheet.Cells[10, 9] = new Cell("Centre de coût");
            worksheet.Cells[10, 10] = new Cell("Salaire de Base");
            worksheet.Cells[10, 11] = new Cell("Indemnité d'ancienneté");
            worksheet.Cells[10, 12] = new Cell("Indemnité ICP");
            worksheet.Cells[10, 13] = new Cell("Indemnité de fonction");
            worksheet.Cells[10, 14] = new Cell("Indemnité de caisse");
            worksheet.Cells[10, 15] = new Cell("Prime  Charges Supportées");
            worksheet.Cells[10, 16] = new Cell("HS 125% (HS01)");
            worksheet.Cells[10, 17] = new Cell("Valeur HS 125% (4110)");
            worksheet.Cells[10, 18] = new Cell("HS 150% (HS02)");
            worksheet.Cells[10, 19] = new Cell("Valeur HS 150% (4111)");
            worksheet.Cells[10, 20] = new Cell("HS 100% (HS03)");
            worksheet.Cells[10, 21] = new Cell("Valeur HS 100% (4113)");
            worksheet.Cells[10, 22] = new Cell("HS 125% EX");
            worksheet.Cells[10, 23] = new Cell("Indemnité Forfaitaire HS");
            worksheet.Cells[10, 24] = new Cell("Congés Pris");
            worksheet.Cells[10, 25] = new Cell("Indemnité Prés Siège Férié");
            worksheet.Cells[10, 26] = new Cell("Retard");
            worksheet.Cells[10, 27] = new Cell("Allocation Personnelle");
            worksheet.Cells[10, 28] = new Cell("Indemnité Transport Prés Siège Férié");
            worksheet.Cells[10, 29] = new Cell("Indemnité Répatriation  Allowance");
            worksheet.Cells[10, 30] = new Cell("Indemnité à l'Amenagement");
            worksheet.Cells[10, 31] = new Cell("Prime Exceptionnelle");
            worksheet.Cells[10, 32] = new Cell("Prime De Performance");
            worksheet.Cells[10, 33] = new Cell("Prime De Rendement");
            worksheet.Cells[10, 34] = new Cell("Prime De Productivité");
            worksheet.Cells[10, 35] = new Cell("Prime Fin D'Année");
            worksheet.Cells[10, 36] = new Cell("Prime de Bilan");
            worksheet.Cells[10, 37] = new Cell("Indemnité De Déplacement  Avec Nuité");
            worksheet.Cells[10, 38] = new Cell("Indemnité De Déplacement Journ");
            worksheet.Cells[10, 39] = new Cell("Indemnité De Voyage");
            worksheet.Cells[10, 40] = new Cell("Indemnité  Fort  Jour Dispatch");
            worksheet.Cells[10, 41] = new Cell("Prime Pour Vêtement De Travail");
            worksheet.Cells[10, 42] = new Cell("Prime De Naissance");
            worksheet.Cells[10, 43] = new Cell("Prime De Mariage");
            worksheet.Cells[10, 44] = new Cell("Prime De Décès");
            worksheet.Cells[10, 45] = new Cell("Prime De Scolarité");
            worksheet.Cells[10, 46] = new Cell("Prime De  L'aid");
            worksheet.Cells[10, 47] = new Cell("Prime D'évènements Familiaux");
            worksheet.Cells[10, 48] = new Cell("Indemnité  De Préavis");
            worksheet.Cells[10, 49] = new Cell("Economie banque");
            worksheet.Cells[10, 50] = new Cell("Gratification Fin De Services");
            worksheet.Cells[10, 51] = new Cell("Indemnité De Licenciment");
            worksheet.Cells[10, 52] = new Cell("IFRT");
            worksheet.Cells[10, 53] = new Cell("Indemnité De Retraite Anticipé");
            worksheet.Cells[10, 54] = new Cell("Rappel");
            worksheet.Cells[10, 55] = new Cell("Rappel sur Allocation Personnelle");
            worksheet.Cells[10, 56] = new Cell("Rappel En Nbr De Jour");
            worksheet.Cells[10, 57] = new Cell("Prime trimesterielle");
            worksheet.Cells[10, 58] = new Cell("Rappel sur element non recurrent");
            worksheet.Cells[10, 59] = new Cell("Solde congé payé");
            worksheet.Cells[10, 60] = new Cell("Valeur congés payés");
            worksheet.Cells[10, 61] = new Cell("Av charges connexées");
            worksheet.Cells[10, 62] = new Cell("Av en nature bon restauration");
            worksheet.Cells[10, 63] = new Cell("Av en nature Tantum");
            worksheet.Cells[10, 64] = new Cell("Av Nat Frais  Medicaux");
            worksheet.Cells[10, 65] = new Cell("Av nature loyer");
            worksheet.Cells[10, 66] = new Cell("Av Charges Connexées Billet d'Avion");
            worksheet.Cells[10, 67] = new Cell("Av Indemnité De Voyage");
            worksheet.Cells[10, 68] = new Cell("Av Housing");
            worksheet.Cells[10, 69] = new Cell("Av Scolarité enfants");
            worksheet.Cells[10, 70] = new Cell("Av Redevance");
            worksheet.Cells[10, 71] = new Cell("Av Assurance Groupe");
            worksheet.Cells[10, 72] = new Cell("Av Assurance Complémentaire");
            worksheet.Cells[10, 73] = new Cell("Salaire Brut");
            worksheet.Cells[10, 74] = new Cell("Salaire Brut Cotisable");
            worksheet.Cells[10, 75] = new Cell("Rtn CNSS Salariale");
            worksheet.Cells[10, 76] = new Cell("Rtn CNSS Complémentaire Salariale");
            worksheet.Cells[10, 77] = new Cell("Rtn ASS Complémentaire Salariale");
            worksheet.Cells[10, 78] = new Cell("Retenue INPS");
            worksheet.Cells[10, 79] = new Cell("Regularisation INPS");
            worksheet.Cells[10, 80] = new Cell("Salaire Imposable");
            worksheet.Cells[10, 81] = new Cell("IRPP");
            worksheet.Cells[10, 82] = new Cell("Redevance 1%");
            worksheet.Cells[10, 83] = new Cell("Rtn Avantage en nature ASS Group (7115)");
            worksheet.Cells[10, 84] = new Cell("Rtn Avantage en nature ASS Compl");
            worksheet.Cells[10, 85] = new Cell("Rtn Charges Connexées");
            worksheet.Cells[10, 86] = new Cell("Rtn Avantage bon restauration");
            worksheet.Cells[10, 87] = new Cell("Rtn Av indemnité de voyage");
            worksheet.Cells[10, 88] = new Cell("Rtn Avantage Frais Medicaux");
            worksheet.Cells[10, 89] = new Cell("Rtn Avantage Charg Connexées Billet Avion");
            worksheet.Cells[10, 90] = new Cell("Rtn Avantage en nature Loyer");
            worksheet.Cells[10, 91] = new Cell("Rtn Avantage Housing");
            worksheet.Cells[10, 92] = new Cell("Rtn Av Nat Scolarité Enfants");
            worksheet.Cells[10, 93] = new Cell("Remboursement IRPP");
            worksheet.Cells[10, 94] = new Cell("Avance");
            worksheet.Cells[10, 95] = new Cell("Retenue bon restauration");
            worksheet.Cells[10, 96] = new Cell("Retenue Amicale Sergaz");
            worksheet.Cells[10, 97] = new Cell("Divers Retenues");
            worksheet.Cells[10, 98] = new Cell("Contribution Conjoncturelle");
            worksheet.Cells[10, 99] = new Cell("Prêt logement");
            worksheet.Cells[10, 100] = new Cell("Prêt Voiture");
            worksheet.Cells[10, 101] = new Cell("Prêt bancaire");
            worksheet.Cells[10, 102] = new Cell("Prêt Exceptionnel");
            worksheet.Cells[10, 103] = new Cell("Remboursement divers");
            worksheet.Cells[10, 104] = new Cell("Remboursement Frais Mission à l'étranger");
            worksheet.Cells[10, 105] = new Cell("Remboursement Frais Mission TN");
            worksheet.Cells[10, 106] = new Cell("Total déduction employé");
            worksheet.Cells[10, 107] = new Cell("Net à Payer");
            worksheet.Cells[10, 108] = new Cell("CNSS Patronale");
            worksheet.Cells[10, 109] = new Cell("CNSS Compélmentaire employeur");
            worksheet.Cells[10, 110] = new Cell("Accident de Travail");
            worksheet.Cells[10, 111] = new Cell("Ass Maladie Patronale");
            worksheet.Cells[10, 112] = new Cell("Ass retraite complémentaire Patronale");
            worksheet.Cells[10, 113] = new Cell("Total Contribution patronale");
            worksheet.Cells[10, 114] = new Cell("Mois de paie");


            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A015(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Employee Local ID");
            worksheet.Cells[10, 1] = new Cell("Last Name");
            worksheet.Cells[10, 2] = new Cell("First name");
            worksheet.Cells[10, 3] = new Cell("Date of Hire company");
            worksheet.Cells[10, 4] = new Cell("Departure Date Company");
            worksheet.Cells[10, 5] = new Cell("Department code");
            worksheet.Cells[10, 6] = new Cell("Title department");
            worksheet.Cells[10, 7] = new Cell("Nature of Contract");
            worksheet.Cells[10, 8] = new Cell("Qualification");
            worksheet.Cells[10, 9] = new Cell("Annual basic salary");
            worksheet.Cells[10, 10] = new Cell("Monthly Basic salary");
            worksheet.Cells[10, 11] = new Cell("Presence");
            worksheet.Cells[10, 12] = new Cell("Basic salary working days");
            worksheet.Cells[10, 13] = new Cell("Meal Allowance");
            worksheet.Cells[10, 14] = new Cell("Transport allowance");
            worksheet.Cells[10, 15] = new Cell("Equalization Allowance");
            worksheet.Cells[10, 16] = new Cell("Exeptional Allowance");
            worksheet.Cells[10, 17] = new Cell("13th month");
            worksheet.Cells[10, 18] = new Cell("Lump-sum");
            worksheet.Cells[10, 19] = new Cell("Gym Allowance");
            worksheet.Cells[10, 20] = new Cell("Birth Allowance");
            worksheet.Cells[10, 21] = new Cell("Aid Idha Allowance");
            worksheet.Cells[10, 22] = new Cell("Aid Fitr Allowance");
            worksheet.Cells[10, 23] = new Cell("Spot Bonus");
            worksheet.Cells[10, 24] = new Cell("Redundancy allowance related to leave days");
            worksheet.Cells[10, 25] = new Cell("Recall paydate");
            worksheet.Cells[10, 26] = new Cell("Company car Benefit in kind");
            worksheet.Cells[10, 27] = new Cell("Petrol benefit in kind");
            worksheet.Cells[10, 28] = new Cell("Housing benefit in kind");
            worksheet.Cells[10, 29] = new Cell("Total Gross");
            worksheet.Cells[10, 30] = new Cell("Employee social contribution");
            worksheet.Cells[10, 31] = new Cell("Employee social complementary contribution");
            worksheet.Cells[10, 32] = new Cell("Taxable salary");
            worksheet.Cells[10, 33] = new Cell("Withholding tax");
            worksheet.Cells[10, 34] = new Cell("Tax 1%");
            worksheet.Cells[10, 35] = new Cell("Refund");
            worksheet.Cells[10, 36] = new Cell("Advantage car");
            worksheet.Cells[10, 37] = new Cell("Group insurance deduction");
            worksheet.Cells[10, 38] = new Cell("Total  employee deduction");
            worksheet.Cells[10, 39] = new Cell("NET TO PAY");
            worksheet.Cells[10, 40] = new Cell("Employer social contribution");
            worksheet.Cells[10, 41] = new Cell("Employer social complementary contribution");
            worksheet.Cells[10, 42] = new Cell("work accident");
            worksheet.Cells[10, 43] = new Cell("Total Employer Contribution");
            worksheet.Cells[10, 44] = new Cell("Month Pay");
          
            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A013(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Id Oracle");
            worksheet.Cells[10, 1] = new Cell("Matricule");
            worksheet.Cells[10, 2] = new Cell("Nom");
            worksheet.Cells[10, 3] = new Cell("Prénom");
            worksheet.Cells[10, 4] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 5] = new Cell("Date de départ société");
            worksheet.Cells[10, 6] = new Cell("Emploi occupé");
            worksheet.Cells[10, 7] = new Cell("Nature du contrat");
            worksheet.Cells[10, 8] = new Cell("Code Département");
            worksheet.Cells[10, 9] = new Cell("Département");
            worksheet.Cells[10, 10] = new Cell("Horaire de base du salarié");
            worksheet.Cells[10, 11] = new Cell("Bulletin Mod");
            worksheet.Cells[10, 12] = new Cell("Catégorie");
            worksheet.Cells[10, 13] = new Cell("Mode de paiement");
            worksheet.Cells[10, 14] = new Cell("Salaire de Base");
            worksheet.Cells[10, 15] = new Cell("Bourse Formation (1300)");
            worksheet.Cells[10, 16] = new Cell("Remb Bourse Formation");
            worksheet.Cells[10, 17] = new Cell("Ind Présence (2100-CL63)");
            worksheet.Cells[10, 18] = new Cell("Ind Transport (2200-CL64)");
            worksheet.Cells[10, 19] = new Cell("Prime de fonction");
            worksheet.Cells[10, 20] = new Cell("Prime de transport");
            worksheet.Cells[10, 21] = new Cell("Nbr Jours Absence (HA03+congé négatif)");
            worksheet.Cells[10, 22] = new Cell("Absence Non Rémunerés (4174)");
            worksheet.Cells[10, 23] = new Cell("HS 125% (HS01)");
            worksheet.Cells[10, 24] = new Cell("Valeur HS 125% (4110)");
            worksheet.Cells[10, 25] = new Cell("HS 150% (HS02)");
            worksheet.Cells[10, 26] = new Cell("Valeur HS 150% (4111)");
            worksheet.Cells[10, 27] = new Cell("HS 200% (HS03)");
            worksheet.Cells[10, 28] = new Cell("Valeur HS 200% (4113)");
            worksheet.Cells[10, 29] = new Cell("Heures de Nuit (HS04)");
            worksheet.Cells[10, 30] = new Cell("Majoration de Nuit (4120)");
            worksheet.Cells[10, 31] = new Cell("Nbre de dimanche travaillés");
            worksheet.Cells[10, 32] = new Cell("Prime jour de dimanche travaillé");
            worksheet.Cells[10, 33] = new Cell("Rappel Nbre Dimanche Trav");
            worksheet.Cells[10, 34] = new Cell("Jour Fériés (HS05)");
            worksheet.Cells[10, 35] = new Cell("Valeur Jours Fériés Payés (4114)");
            worksheet.Cells[10, 36] = new Cell("Congé Pris (HA04)");
            worksheet.Cells[10, 37] = new Cell("Jour Fériés ETRANGER Non Payés (HA06)");
            worksheet.Cells[10, 38] = new Cell("Valeur Jour Fériés ETRANGER non payés (4175)");
            worksheet.Cells[10, 39] = new Cell("Jour Trop Perçu (HA07)");
            worksheet.Cells[10, 40] = new Cell("Valeur Jours Trop Perçu (4176)");
            worksheet.Cells[10, 41] = new Cell("Prime Astreinte (4130-CL46)");
            worksheet.Cells[10, 42] = new Cell("Maj Prime Astreinte (4135-CL47)");
            worksheet.Cells[10, 43] = new Cell("Prime d'assiduité conventionnelle");
            worksheet.Cells[10, 44] = new Cell("Prime Exceptionnelle (4222-CL05)");
            worksheet.Cells[10, 45] = new Cell("Prime de Cooptation (3309-CL40)");
            worksheet.Cells[10, 46] = new Cell("Prime de Mission");
            worksheet.Cells[10, 47] = new Cell("Prime de Commisssion (4371-CL58)");
            worksheet.Cells[10, 48] = new Cell("Prime MIP TR (4817-CL53)");
            worksheet.Cells[10, 49] = new Cell("Prime MIP AN (4818-CL54)");
            worksheet.Cells[10, 50] = new Cell("Prime de Performance Trimestrielle");
            worksheet.Cells[10, 51] = new Cell("Prime de Challenge");
            worksheet.Cells[10, 52] = new Cell("Rappel Jours (HA09)");
            worksheet.Cells[10, 53] = new Cell("Valeur Rappel Jours (5500)");
            worksheet.Cells[10, 54] = new Cell("Rap Salaire de Base (5100-CL24)");
            worksheet.Cells[10, 55] = new Cell("Rap Prime (5400-CL25)");
            worksheet.Cells[10, 56] = new Cell("Rap Augmentation (5200-CL28)");
            worksheet.Cells[10, 57] = new Cell("Rap Divers (5900-CL27)");
            worksheet.Cells[10, 58] = new Cell("STC (HA01)");
            worksheet.Cells[10, 59] = new Cell("Prime STC (4816-CL06)");
            worksheet.Cells[10, 60] = new Cell("Ind Licenciement (4810-CL19)");
            worksheet.Cells[10, 61] = new Cell("Ind Préavis (4811-CL20)");
            worksheet.Cells[10, 62] = new Cell("Gratification Fin Service (4812-CL21)");
            worksheet.Cells[10, 63] = new Cell("Nbr Jours congé (teststc)");
            worksheet.Cells[10, 64] = new Cell("Valeur Congés Payés (4800)");
            worksheet.Cells[10, 65] = new Cell("Nbr Ticket Rest (pres8)");
            worksheet.Cells[10, 66] = new Cell("Av en Nature Tickets Restaurant");
            worksheet.Cells[10, 67] = new Cell("Rappel Nbre TR");
            worksheet.Cells[10, 68] = new Cell("Rappel Ticket restaurant");
            worksheet.Cells[10, 69] = new Cell("Avantage Ass Groupe (7115)");
            worksheet.Cells[10, 70] = new Cell("Salaire Brut");
            worksheet.Cells[10, 71] = new Cell("CNSS Employé");
            worksheet.Cells[10, 72] = new Cell("Salaire Imposable");
            worksheet.Cells[10, 73] = new Cell("IRPP");
            worksheet.Cells[10, 74] = new Cell("Redevance 1%");
            worksheet.Cells[10, 75] = new Cell("Contribution Conjoncturelle");
            worksheet.Cells[10, 76] = new Cell("Avance (6410)");
            worksheet.Cells[10, 77] = new Cell("Retenues Trop Perçu (6490-CL30)");
            worksheet.Cells[10, 78] = new Cell("Retenus Déterioration perte (6491-CL31)");
            worksheet.Cells[10, 79] = new Cell("Remb Détérioration perte (6491-CL32)");
            worksheet.Cells[10, 80] = new Cell("Retenue UGTT");
            worksheet.Cells[10, 81] = new Cell("Déduction Avantage en Nature TR");
            worksheet.Cells[10, 82] = new Cell("Déduction Rappel Avantage en Nature TR");
            worksheet.Cells[10, 83] = new Cell("Avantage Ass Groupe (6840)");
            worksheet.Cells[10, 84] = new Cell("Total déduction employé");
            worksheet.Cells[10, 85] = new Cell("Net à payer");
            worksheet.Cells[10, 86] = new Cell("CNSS employeur");
            worksheet.Cells[10, 87] = new Cell("Accident de travail");
            worksheet.Cells[10, 88] = new Cell("Assurance groupe employeur");
            worksheet.Cells[10, 89] = new Cell("Total contribution employeur");
            worksheet.Cells[10, 90] = new Cell("Mois de paie");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A020(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Id Oracle");
            worksheet.Cells[10, 1] = new Cell("Matricule");
            worksheet.Cells[10, 2] = new Cell("Nom");
            worksheet.Cells[10, 3] = new Cell("Prénom");
            worksheet.Cells[10, 4] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 5] = new Cell("Date de départ société");
            worksheet.Cells[10, 6] = new Cell("Code département");
            worksheet.Cells[10, 7] = new Cell("Intitulé département");
            worksheet.Cells[10, 8] = new Cell("EXPAT / LOCAL");
            worksheet.Cells[10, 9] = new Cell("Emploi occupé");
            worksheet.Cells[10, 10] = new Cell("Centre de Coût");
            worksheet.Cells[10, 11] = new Cell("Salaire de Base");
            worksheet.Cells[10, 12] = new Cell("Indemnité d'ancienneté");
            worksheet.Cells[10, 13] = new Cell("Indemnité  ICP");
            worksheet.Cells[10, 14] = new Cell("Indemnité de Fonction");
            worksheet.Cells[10, 15] = new Cell("Indemnité de caisse");
            worksheet.Cells[10, 16] = new Cell("Indemnité de disponibilité");
            worksheet.Cells[10, 17] = new Cell("Prime Charges Supportées");
            worksheet.Cells[10, 18] = new Cell("HS 25% (HS01)");
            worksheet.Cells[10, 19] = new Cell("Valeur HS 25% (4110)");
            worksheet.Cells[10, 20] = new Cell("HS 50% (HS02)");
            worksheet.Cells[10, 21] = new Cell("Valeur HS 50% (4111)");
            worksheet.Cells[10, 22] = new Cell("HS 100% (HS03)");
            worksheet.Cells[10, 23] = new Cell("Valeur HS 100% (4113)");
            worksheet.Cells[10, 24] = new Cell("HS 25% EX");
            worksheet.Cells[10, 25] = new Cell("Indemnité Forfaitaire HS");
            worksheet.Cells[10, 26] = new Cell("Indemnité Présence Siège Férié");
            worksheet.Cells[10, 27] = new Cell("Retard");
            worksheet.Cells[10, 28] = new Cell("Allocation Personnelle");
            worksheet.Cells[10, 29] = new Cell("Indemnité Transp Prés Sieg Férié");
            worksheet.Cells[10, 30] = new Cell("Prime  Exceptionnelle");
            worksheet.Cells[10, 31] = new Cell("Prime de  Performance");
            worksheet.Cells[10, 32] = new Cell("Prime de Rendement");
            worksheet.Cells[10, 33] = new Cell("Prime de  Productivité");
            worksheet.Cells[10, 34] = new Cell("Prime Fin D'année");
            worksheet.Cells[10, 35] = new Cell("Prime de Bilan");
            worksheet.Cells[10, 36] = new Cell("Una Tantum");
            worksheet.Cells[10, 37] = new Cell("Indemnité de Dépl avec Nuité");
            worksheet.Cells[10, 38] = new Cell("Indemnité Repatriation Allowance");
            worksheet.Cells[10, 39] = new Cell("Indemnité à l'Aménagement");
            worksheet.Cells[10, 40] = new Cell("Indemnité de Déplacement  Journ");
            worksheet.Cells[10, 41] = new Cell("Indemnité  de Voyage");
            worksheet.Cells[10, 42] = new Cell("Indemnité  Fort  Jour  Dispatch");
            worksheet.Cells[10, 43] = new Cell("Prime pour  Vêtement  de Travail");
            worksheet.Cells[10, 44] = new Cell("Prime de Naissance");
            worksheet.Cells[10, 45] = new Cell("Prime de Mariage");
            worksheet.Cells[10, 46] = new Cell("Prime de Décés");
            worksheet.Cells[10, 47] = new Cell("Prime de Scolarité");
            worksheet.Cells[10, 48] = new Cell("Prime de l 'Aid");
            worksheet.Cells[10, 49] = new Cell("Prime Evénèments Familiaux");
            worksheet.Cells[10, 50] = new Cell("Indemnité de Préavis");
            worksheet.Cells[10, 51] = new Cell("Gratification Fin de Service");
            worksheet.Cells[10, 52] = new Cell("Indemnité de Licenciment");
            worksheet.Cells[10, 53] = new Cell("Indemnité IFRT");
            worksheet.Cells[10, 54] = new Cell("Indemnité de Retraite Anticipée");
            worksheet.Cells[10, 55] = new Cell("Liquidation Ifrt");
            worksheet.Cells[10, 56] = new Cell("Rappel");
            worksheet.Cells[10, 57] = new Cell("Rappel en Nbre de Jour");
            worksheet.Cells[10, 58] = new Cell("Rappel sur Allocation personnelle");
            worksheet.Cells[10, 59] = new Cell("Rappel sur element non recurrent");
            worksheet.Cells[10, 60] = new Cell("Prime Trimesterielle");
            worksheet.Cells[10, 61] = new Cell("Congé Pris");
            worksheet.Cells[10, 62] = new Cell("Congés Payés");
            worksheet.Cells[10, 63] = new Cell("Solde Congé Payé");
            worksheet.Cells[10, 64] = new Cell("Valeur Congés Payés");
            worksheet.Cells[10, 65] = new Cell("Av en Nature bon restauration");
            worksheet.Cells[10, 66] = new Cell("Av Nat Frais Medicaux");
            worksheet.Cells[10, 67] = new Cell("Rtn Av Charges Connexées");
            worksheet.Cells[10, 68] = new Cell(" Av Indemnité  de Voyage");
            worksheet.Cells[10, 69] = new Cell(" Av Redevance");
            worksheet.Cells[10, 70] = new Cell("Av Assurance Maladie");
            worksheet.Cells[10, 71] = new Cell("Av Assurance Complémentaire");
            worksheet.Cells[10, 72] = new Cell("Salaire Brut");
            worksheet.Cells[10, 73] = new Cell("Salaire Brut Cotisable");
            worksheet.Cells[10, 74] = new Cell("CNSS Employé");
            worksheet.Cells[10, 75] = new Cell("CNSS Complémentaire  Employé");
            worksheet.Cells[10, 76] = new Cell("Ass Retraite Complémentaire  Employé");
            worksheet.Cells[10, 77] = new Cell("Retenue INPS");
            worksheet.Cells[10, 78] = new Cell("Salaire Imposable");
            worksheet.Cells[10, 79] = new Cell("IRPP");
            worksheet.Cells[10, 80] = new Cell("Redevance 1%");
            worksheet.Cells[10, 81] = new Cell("Rtn Av Assurance Group");
            worksheet.Cells[10, 82] = new Cell("Rtn Av en nature Ass Complémentaire");
            worksheet.Cells[10, 83] = new Cell("Rtn Av Charges Connexées");
            worksheet.Cells[10, 84] = new Cell("Rtn Av en nature bon restauration");
            worksheet.Cells[10, 85] = new Cell("Rtn Av Indemnité de voyage");
            worksheet.Cells[10, 86] = new Cell("Rtn Av Frais mèdicaux");
            worksheet.Cells[10, 87] = new Cell("Rtn Av Charges Connexées-Billet AVIO");
            worksheet.Cells[10, 88] = new Cell("Remb IRPP");
            worksheet.Cells[10, 89] = new Cell("Avance");
            worksheet.Cells[10, 90] = new Cell("Retenue bon restauration");
            worksheet.Cells[10, 91] = new Cell("Retenue Amicale");
            worksheet.Cells[10, 92] = new Cell("Divers retenues");
            worksheet.Cells[10, 93] = new Cell("Prêt Logement");
            worksheet.Cells[10, 94] = new Cell("Prêt voiture");
            worksheet.Cells[10, 95] = new Cell("Prêt Bancaire");
            worksheet.Cells[10, 96] = new Cell("Prêt exceptionnel");
            worksheet.Cells[10, 97] = new Cell("Contribution Conjoncturelle");
            worksheet.Cells[10, 98] = new Cell("Remboursement divers");
            worksheet.Cells[10, 99] = new Cell("Rem Frais Mission à l'etranger");
            worksheet.Cells[10, 100] = new Cell("Rem Frais Mission TN");
            worksheet.Cells[10, 101] = new Cell("Total Déduction Employé");
            worksheet.Cells[10, 102] = new Cell("Salaire Net");
            worksheet.Cells[10, 103] = new Cell("CNSS employeur");
            worksheet.Cells[10, 104] = new Cell("CNSS compémentaire employeur");
            worksheet.Cells[10, 105] = new Cell("Accident de travail");
            worksheet.Cells[10, 106] = new Cell("Ass maladie employeur");
            worksheet.Cells[10, 107] = new Cell("Ass retraite complémentaire employeur");
            worksheet.Cells[10, 108] = new Cell("Total contribution employeur");
            worksheet.Cells[10, 109] = new Cell("Mois de paie");


            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A016(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("ID Oracle");
            worksheet.Cells[10, 1] = new Cell("Last Name");
            worksheet.Cells[10, 2] = new Cell("First Name");
            worksheet.Cells[10, 3] = new Cell("Hire date");
            worksheet.Cells[10, 4] = new Cell("Date of Departure");
            worksheet.Cells[10, 5] = new Cell("Cost center");
            worksheet.Cells[10, 6] = new Cell("Departement");
            worksheet.Cells[10, 7] = new Cell("Contract Nature");
            worksheet.Cells[10, 8] = new Cell("Employement Title");
            worksheet.Cells[10, 9] = new Cell("Basic Salary");
            worksheet.Cells[10, 10] = new Cell("Hourly presence");
            worksheet.Cells[10, 11] = new Cell("Basic salary working days");
            worksheet.Cells[10, 12] = new Cell("Attendance allowance");
            worksheet.Cells[10, 13] = new Cell("Transportation allowance");
            worksheet.Cells[10, 14] = new Cell("Shift allowance ");
            worksheet.Cells[10, 15] = new Cell("Site allowance ");
            worksheet.Cells[10, 16] = new Cell("Offshores allowance ");
            worksheet.Cells[10, 17] = new Cell("Off shore I Allowance");
            worksheet.Cells[10, 18] = new Cell("Offshores astreinte allowance ");
            worksheet.Cells[10, 19] = new Cell("Cashier allowance ");
            worksheet.Cells[10, 20] = new Cell("Child care allowance 1st children");
            worksheet.Cells[10, 21] = new Cell("Child care allowance  2nd children");
            worksheet.Cells[10, 22] = new Cell("Child care allowance 3rd children");
            worksheet.Cells[10, 23] = new Cell("Duty  allowance");
            worksheet.Cells[10, 24] = new Cell("Voluntary additional contribution ");
            worksheet.Cells[10, 25] = new Cell("President driver  allowance");
            worksheet.Cells[10, 26] = new Cell("Rotational allowance");
            worksheet.Cells[10, 27] = new Cell("Banking allowance");
            worksheet.Cells[10, 28] = new Cell("Mobility allowance LT");
            worksheet.Cells[10, 29] = new Cell("Mobility allowance ST");
            worksheet.Cells[10, 30] = new Cell("Hardship  premium allowance  LT");
            worksheet.Cells[10, 31] = new Cell("Hardship  premium allowance  ST");
            worksheet.Cells[10, 32] = new Cell("Repatriation  allowance");
            worksheet.Cells[10, 33] = new Cell("C&S allowance LT");
            worksheet.Cells[10, 34] = new Cell("C&S allowance ST");
            worksheet.Cells[10, 35] = new Cell("Cash in lieu of car (for overseas expartiates)");
            worksheet.Cells[10, 36] = new Cell("International peer spendable");
            worksheet.Cells[10, 37] = new Cell("Disruption allowance LT");
            worksheet.Cells[10, 38] = new Cell("Disruption allowance ST");
            worksheet.Cells[10, 39] = new Cell("Foreign spend allowance LT");
            worksheet.Cells[10, 40] = new Cell("Home leave allowance LT");
            worksheet.Cells[10, 41] = new Cell("Bawab allowance LT");
            worksheet.Cells[10, 42] = new Cell("Fixed gross salary");
            worksheet.Cells[10, 43] = new Cell("Partner allowance LT");
            worksheet.Cells[10, 44] = new Cell("Currency paid overseas");
            worksheet.Cells[10, 45] = new Cell("Nbr of shift daily allowance");
            worksheet.Cells[10, 46] = new Cell("Shift daily allowance");
            worksheet.Cells[10, 47] = new Cell("Nbr Site daily allowance");
            worksheet.Cells[10, 48] = new Cell("Site daily allowance");
            worksheet.Cells[10, 49] = new Cell("Nbr of Miskar offshores daily variable allowance");
            worksheet.Cells[10, 50] = new Cell("Miskar offshores daily variable allowance");
            worksheet.Cells[10, 51] = new Cell("Nbr of Hasdrubal offshores daily allowance");
            worksheet.Cells[10, 52] = new Cell("Hasdrubal offshores daily allowance");
            worksheet.Cells[10, 53] = new Cell("Nbr of Mahras transportation daily allowance");
            worksheet.Cells[10, 54] = new Cell("Mahras transportation allowance ");
            worksheet.Cells[10, 55] = new Cell("Nbr of Standby daily allowance ");
            worksheet.Cells[10, 56] = new Cell("Standby allowance");
            worksheet.Cells[10, 57] = new Cell("Standby allowance extra-offshore");
            worksheet.Cells[10, 58] = new Cell("Nbr of Callout- In daily Allowance");
            worksheet.Cells[10, 59] = new Cell("Callout- In Allowance");
            worksheet.Cells[10, 60] = new Cell("Nbr of Callout-  Out daily Allowance");
            worksheet.Cells[10, 61] = new Cell("Callout-  Out Allowance");
            worksheet.Cells[10, 62] = new Cell("Nbr of Deputizing daily allowance");
            worksheet.Cells[10, 63] = new Cell("Deputizing allowance");
            worksheet.Cells[10, 64] = new Cell("Relocation allowance");
            worksheet.Cells[10, 65] = new Cell("Shutdown Bonus");
            worksheet.Cells[10, 66] = new Cell("Other Bonus");
            worksheet.Cells[10, 67] = new Cell("AIS- Annual Incentive Scheme");
            worksheet.Cells[10, 68] = new Cell("Prorata of AIS");
            worksheet.Cells[10, 69] = new Cell("Biannual Bonus");
            worksheet.Cells[10, 70] = new Cell("Prorata of mid year Bonus");
            worksheet.Cells[10, 71] = new Cell("13th month Bonus");
            worksheet.Cells[10, 72] = new Cell("Prorata of 13the month Bonus");
            worksheet.Cells[10, 73] = new Cell("Aid Idha sheep Allowance");
            worksheet.Cells[10, 74] = new Cell("Aid Fitr  Allowance");
            worksheet.Cells[10, 75] = new Cell("Peligrinage allowance");
            worksheet.Cells[10, 76] = new Cell("School allowance");
            worksheet.Cells[10, 77] = new Cell("Nbr of Hourly Night shift majoration 10.5%");
            worksheet.Cells[10, 78] = new Cell("Hourly Night shift majoration 10.5%");
            worksheet.Cells[10, 79] = new Cell("Nbr of Overtime 25%");
            worksheet.Cells[10, 80] = new Cell("Overtime 25%");
            worksheet.Cells[10, 81] = new Cell("Nbr of Over time 100% at normal rate");
            worksheet.Cells[10, 82] = new Cell("Over time 100% at normal rate");
            worksheet.Cells[10, 83] = new Cell("Nbr of Over time 150% ");
            worksheet.Cells[10, 84] = new Cell("Over time 150%");
            worksheet.Cells[10, 85] = new Cell("Nbr of Over time 175%");
            worksheet.Cells[10, 86] = new Cell("Over time 175%");
            worksheet.Cells[10, 87] = new Cell("Nbr of Over time 200%");
            worksheet.Cells[10, 88] = new Cell("Over time 200%");
            worksheet.Cells[10, 89] = new Cell("Nbr of Over time 225%");
            worksheet.Cells[10, 90] = new Cell("Over time 225%");
            worksheet.Cells[10, 91] = new Cell("Salary absence deduction");
            worksheet.Cells[10, 92] = new Cell("Sickness deduction");
            worksheet.Cells[10, 93] = new Cell("Ecart shift/site");
            worksheet.Cells[10, 94] = new Cell("Ecart offshores shift/site");
            worksheet.Cells[10, 95] = new Cell("Pay back");
            worksheet.Cells[10, 96] = new Cell("Pay back in days");
            worksheet.Cells[10, 97] = new Cell("Pay back in hours");
            worksheet.Cells[10, 98] = new Cell("Pay back transportation allowance");
            worksheet.Cells[10, 99] = new Cell("Payback of variables or non-recurring premiums");
            worksheet.Cells[10, 100] = new Cell("Pay back period benefits in kind");
            worksheet.Cells[10, 101] = new Cell("Other pay back period");
            worksheet.Cells[10, 102] = new Cell("Car/ cash equivalent assigned car");
            worksheet.Cells[10, 103] = new Cell("Benefits in kind : Oil card");
            worksheet.Cells[10, 104] = new Cell("Housing");
            worksheet.Cells[10, 105] = new Cell("Scholarship benefit in kind");
            worksheet.Cells[10, 106] = new Cell("Utilities ");
            worksheet.Cells[10, 107] = new Cell("Meal voucher benefit in kind");
            worksheet.Cells[10, 108] = new Cell("Salary adjustment (+)");
            worksheet.Cells[10, 109] = new Cell("Salary adjustment (-)");
            worksheet.Cells[10, 110] = new Cell("Health insurance benefit in kind");
            worksheet.Cells[10, 111] = new Cell("Benefits in Kind expatriates");
            worksheet.Cells[10, 112] = new Cell("Share allowance");
            worksheet.Cells[10, 113] = new Cell("Voluntary deduction");
            worksheet.Cells[10, 114] = new Cell("Payment outstanding holydays");
            worksheet.Cells[10, 115] = new Cell("Notice period for redundancy");
            worksheet.Cells[10, 116] = new Cell("Redundancy allowance");
            worksheet.Cells[10, 117] = new Cell("End of service gratification");
            worksheet.Cells[10, 118] = new Cell("3 month retirement gift");
            worksheet.Cells[10, 119] = new Cell("Total gross salary");
            worksheet.Cells[10, 120] = new Cell("Gross salary BG");
            worksheet.Cells[10, 121] = new Cell("Total gross cotisable");//
            worksheet.Cells[10, 122] = new Cell("Social security contribution");
            worksheet.Cells[10, 123] = new Cell("Complementary Social security contribution -");
            worksheet.Cells[10, 124] = new Cell("Taxable salary");
            worksheet.Cells[10, 125] = new Cell("Witholding tax on revenue");
            worksheet.Cells[10, 126] = new Cell("Hypotax");
            worksheet.Cells[10, 127] = new Cell("Redevance tax 1% (RCGC)");
            worksheet.Cells[10, 128] = new Cell("Sport & social subscription");
            worksheet.Cells[10, 129] = new Cell("CNSS lodging loan");

            worksheet.Cells[10, 130] = new Cell("Contribution meal voucher");
            worksheet.Cells[10, 131] = new Cell("Aid Idha sheep deduction");
            worksheet.Cells[10, 132] = new Cell("School allowance deduction");
            worksheet.Cells[10, 133] = new Cell("Health insurance benefit in kind Deduction");
            worksheet.Cells[10, 134] = new Cell("Car Cash Equivalent deduction benefit in kind");
            worksheet.Cells[10, 135] = new Cell("Meal voucher benefit in kind deduction");
            worksheet.Cells[10, 136] = new Cell("Benefits in kind expatriates deduction");
            worksheet.Cells[10, 137] = new Cell("Salary deductions");
            worksheet.Cells[10, 138] = new Cell("Currency paid overseas");
            worksheet.Cells[10, 139] = new Cell("Share  Deduction");
            worksheet.Cells[10, 140] = new Cell("Account reimbursment");
            worksheet.Cells[10, 141] = new Cell("Total Deduction employee");
            worksheet.Cells[10, 142] = new Cell("Net to pay");
            worksheet.Cells[10, 143] = new Cell(" Social security contribution-employer");
            worksheet.Cells[10, 144] = new Cell("Complementary Social security contribution -");
            worksheet.Cells[10, 145] = new Cell("Work accident");
            worksheet.Cells[10, 146] = new Cell("Insurance Group Contribution-employer");
            worksheet.Cells[10, 147] = new Cell("TFP");
            worksheet.Cells[10, 148] = new Cell("Total employer contribution");
            worksheet.Cells[10, 149] = new Cell("fee Overseas");
            worksheet.Cells[10, 150] = new Cell("Month of pay");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne11A010(Worksheet worksheet)
        {
            worksheet.Cells[10, 0] = new Cell("Id Oracle");
            worksheet.Cells[10, 1] = new Cell("Matricule");
            worksheet.Cells[10, 2] = new Cell("Nom");
            worksheet.Cells[10, 3] = new Cell("Prénom");
            worksheet.Cells[10, 4] = new Cell("Date d'embauche société");
            worksheet.Cells[10, 5] = new Cell("Date de départ société");
            worksheet.Cells[10, 6] = new Cell("Emploi occupé");
            worksheet.Cells[10, 7] = new Cell("Nature du contrat");
            worksheet.Cells[10, 8] = new Cell("Code département");
            worksheet.Cells[10, 9] = new Cell("Département");
            worksheet.Cells[10, 10] = new Cell("Horaire de base du salarié");
            worksheet.Cells[10, 11] = new Cell("Bulletin Mod");
            worksheet.Cells[10, 12] = new Cell("Catégorie");
            worksheet.Cells[10, 13] = new Cell("Salaire de Base");
            worksheet.Cells[10, 14] = new Cell("Bourse Formation (1300)");
            worksheet.Cells[10, 15] = new Cell("Ind Présence (2100-CL63)");
            worksheet.Cells[10, 16] = new Cell("Ind Transport (2200-CL64)");
            worksheet.Cells[10, 17] = new Cell("P Fonction (3310)");
            worksheet.Cells[10, 18] = new Cell("Nbr Jours Absence (HA03+congé négatif)");
            worksheet.Cells[10, 19] = new Cell("Absence Non Rémunerés (4174)");
            worksheet.Cells[10, 20] = new Cell("Arrondi Absences Rémunérées");
            worksheet.Cells[10, 21] = new Cell("Indemnité de mission (4373-CL49)");
            worksheet.Cells[10, 22] = new Cell("HS 125% (HS01)");
            worksheet.Cells[10, 23] = new Cell("Valeur HS 125% (4110)");
            worksheet.Cells[10, 24] = new Cell("HS 150% (HS02)");
            worksheet.Cells[10, 25] = new Cell("Valeur HS 150% (4111)");
            worksheet.Cells[10, 26] = new Cell("HS 200% (HS03)");
            worksheet.Cells[10, 27] = new Cell("Valeur HS 200% (4113)");
            worksheet.Cells[10, 28] = new Cell("HS 175% (HS06)");
            worksheet.Cells[10, 29] = new Cell("Valeur HS 175% (4112)");
            worksheet.Cells[10, 30] = new Cell("Heures de Nuit (HS04)");
            worksheet.Cells[10, 31] = new Cell("Majoration de Nuit (4120)");
            worksheet.Cells[10, 32] = new Cell("Nombre Dimanche Travaillé(HS07)");
            worksheet.Cells[10, 33] = new Cell("Prime Nbr Dimanche Travaillé(4372)");
            worksheet.Cells[10, 34] = new Cell("Jour Fériés (HS05)");
            worksheet.Cells[10, 35] = new Cell("Rappel Dimanche");
            worksheet.Cells[10, 36] = new Cell("Valeur Jours Fériés Payés (4114)");
            worksheet.Cells[10, 37] = new Cell("Congé Pris (HA04)");
            worksheet.Cells[10, 38] = new Cell("Jour Fériés ETRANGER Non Payés (HA06)");
            worksheet.Cells[10, 39] = new Cell("Valeur Jour Fériés ETRANGER non payés (4175)");
            worksheet.Cells[10, 40] = new Cell("Jour Trop Perçu (HA07)");
            worksheet.Cells[10, 41] = new Cell("Valeur Jours Trop Perçu (4176)");
            worksheet.Cells[10, 42] = new Cell("Rappel Jours (HA09)");
            worksheet.Cells[10, 43] = new Cell("Valeur Rappel Jours (5500)");
            worksheet.Cells[10, 44] = new Cell("Prime Exceptionnelle (4222-CL05)");
            worksheet.Cells[10, 45] = new Cell("Prime de Commisssion (4371-CL58)");
            worksheet.Cells[10, 46] = new Cell("Prime Astreinte (4130-CL46)");
            worksheet.Cells[10, 47] = new Cell("Maj Prime Astreinte (4135-CL47)");
            worksheet.Cells[10, 48] = new Cell("Prime Assiduité (2361)");
            worksheet.Cells[10, 49] = new Cell("Prime STC (4816-CL06)");
            worksheet.Cells[10, 50] = new Cell("Ind Licenciement (4810-CL19)");
            worksheet.Cells[10, 51] = new Cell("STC (HA01)");
            worksheet.Cells[10, 52] = new Cell("Ind Préavis (4811-CL20)");
            worksheet.Cells[10, 53] = new Cell("Gratification Fin Service (4812-CL21)");
            worksheet.Cells[10, 54] = new Cell("Rap Salaire de Base (5100-CL24)");
            worksheet.Cells[10, 55] = new Cell("Rap Prime (5400-CL25)");
            worksheet.Cells[10, 56] = new Cell("Rap Divers (5900-CL27)");
            worksheet.Cells[10, 57] = new Cell("Rap Augmentation (5200-CL28)");
            worksheet.Cells[10, 58] = new Cell("Opposition CNSS (6510-CL29)");
            worksheet.Cells[10, 59] = new Cell("Prime de Cooptation (3309-CL40)");
            worksheet.Cells[10, 60] = new Cell("Prime MIP TR (4817-CL53)");
            worksheet.Cells[10, 61] = new Cell("Prime MIP AN (4818-CL54)");
            worksheet.Cells[10, 62] = new Cell("Prime trimestrielle(CL52 4313)");
            worksheet.Cells[10, 63] = new Cell("Prime de Performance(4312-CL51)");
            worksheet.Cells[10, 64] = new Cell("Nbr Jours STC (teststc)");
            worksheet.Cells[10, 65] = new Cell("Valeur Congés Payés (4800)");
            worksheet.Cells[10, 66] = new Cell("Nbr Ticket Rest (pres8)");
            worksheet.Cells[10, 67] = new Cell("Valeur Ticket Restaurant (pres2)");
            worksheet.Cells[10, 68] = new Cell("Avantage Ass Groupe (7115)");
            worksheet.Cells[10, 69] = new Cell("Salaire Brut");
            worksheet.Cells[10, 70] = new Cell("CNSS Salariale");
            worksheet.Cells[10, 71] = new Cell("Salaire Imposable");
            worksheet.Cells[10, 72] = new Cell("IRPP (6310)");
            worksheet.Cells[10, 73] = new Cell("Redevance 1%(6311)");
            worksheet.Cells[10, 74] = new Cell("Retenues Trop Perçu (6490-CL30)");
            worksheet.Cells[10, 75] = new Cell("Retenus Déterioration perte (6491-CL31)");
            worksheet.Cells[10, 76] = new Cell("Remb Détérioration perte (6491-CL32)");
            worksheet.Cells[10, 77] = new Cell("Avantage Ass Groupe (6840)");
            worksheet.Cells[10, 78] = new Cell("Remb Redevance antérieurs (6610)");
            worksheet.Cells[10, 79] = new Cell("Retenus Redevance antérieurs (6611)");
            worksheet.Cells[10, 80] = new Cell("Remb IRPP antérieures (6613)");
            worksheet.Cells[10, 81] = new Cell("Retenue IRPP antérieures (6612)");
            worksheet.Cells[10, 82] = new Cell("Avance (6410)");
            worksheet.Cells[10, 83] = new Cell("Retenue cotisation UGTT ( 6511-CL11)");
            worksheet.Cells[10, 84] = new Cell("Contribution Conjoncturelle");
            worksheet.Cells[10, 85] = new Cell("Total déduction salariale");
            worksheet.Cells[10, 86] = new Cell("NET A PAYER");
            worksheet.Cells[10, 87] = new Cell("CNSS employeur");
            worksheet.Cells[10, 88] = new Cell("CNSS Accident de travail");
            worksheet.Cells[10, 89] = new Cell("Assurance groupe");
            worksheet.Cells[10, 90] = new Cell("Total contribution patronale");
            worksheet.Cells[10, 91] = new Cell("Mode de paiement");
            worksheet.Cells[10, 92] = new Cell("Mois de paie");

            return worksheet;
        }

        public Worksheet EnteteSVRLigne12(Worksheet worksheet, int refSoc,int MoisPe,int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGN(refSoc, MoisPe,AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Mat = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string CodeSce = rez["CodeSce"].ToString();
                string IntituleSce = rez["IntituleSce"].ToString();
                string NatureContrat = rez["NatureContrat"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string A1100 = rez["A1100"].ToString();
                string ACL01 = rez["CL01"].ToString();
                string A1200 = rez["A1200"].ToString();
                string A2200 = rez["A2200"].ToString();
                string A2100 = rez["A2100"].ToString();
                string A2311 = rez["A2311"].ToString();
                string A2320 = rez["A2320"].ToString();
                string A3318 = rez["A3318"].ToString();
                string A3110 = rez["A3110"].ToString();
                string A4311 = rez["A4311"].ToString();
                string A4330 = rez["A4330"].ToString();
                string A3111 = rez["A3111"].ToString();
                string A3210 = rez["A3210"].ToString();
                string A3317 = rez["A3317"].ToString();
                string A3319 = rez["A3319"].ToString();
                string A4100 = rez["A4100"].ToString();
                string A4171 = rez["A4171"].ToString();
                string A4130 = rez["A4130"].ToString();
                string A4223 = rez["A4223"].ToString();
                string A4380 = rez["A4380"].ToString();
                string A4711 = rez["A4711"].ToString();
                string A4712 = rez["A4712"].ToString();
                string A4717 = rez["A4717"].ToString();
                string A4226 = rez["A4226"].ToString();
                string A5100 = rez["A5100"].ToString();
                string A7600 = rez["A7600"].ToString();
                string A4811 = rez["A4811"].ToString();
                string A4812 = rez["A4812"].ToString();
                string A4813 = rez["A4813"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string AJUSCOT = rez["AJUSCOT"].ToString();
                string A9112 = rez["A9112"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A6120 = rez["A6120"].ToString();
                string A6210 = rez["A6210"].ToString();
                string A6230 = rez["A6230"].ToString();
                string A9121 = rez["A9121"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A6614 = rez["A6614"].ToString();
                string A6410 = rez["A6410"].ToString();
                string A6520 = rez["A6520"].ToString();
                string A6510 = rez["A6510"].ToString();
                string A6530 = rez["A6530"].ToString();
                string A6890 = rez["A6890"].ToString();
                string A8000 = rez["A8000"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A6110R =  rez["A61109"].ToString();
                string A6120R = rez["A61209"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A6210R = rez["A62109"].ToString();
                string A6230R = rez["A62309"].ToString();
                string A6320 = rez["A6320"].ToString();
                string A6330 = rez["A6330"].ToString();
                string RUB_COTEPY = rez["RUB_COTEPY"].ToString();
                string MOIS_PAIE = rez["MOIS_PAIE"].ToString();

                worksheet = EcrireSGN(cpt, worksheet, Mat, Nom, Prenom, DateEmbauche, DateDepart, CodeSce, IntituleSce, NatureContrat,
                    Emploi, A1100, ACL01, A1200, A2200, A2100, A2311, A2320, A3318, A3110,A4311,A4330, A3111, A3210, A3317, A3319, A4100, A4171,
                    A4130, A4223, A4380, A4711,A4712, A4717,A4226, A5100, A7600, A4811, A4812, A4813, BRUT, AJUSCOT, A9112, A6110, A6120, A6210,
                    A6230, A9121, A6310, A6311, A6614, A6410, A6520, A6510, A6530, A6890,A8000, RUB_DEDSA, NETPAIE, A6110R, A6120R, A6150,
                    A6210R, A6230R, A6320, A6330, RUB_COTEPY, MOIS_PAIE);
                cpt = cpt + 1;
            }

            return worksheet;
        }

        public Worksheet EnteteSVRLigne12A002(Worksheet worksheet, int refSoc, int MoisPe,int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA002(refSoc, MoisPe,AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Mat = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string CodeSce = rez["CodeSce"].ToString();
                string IntituleSce = rez["IntituleSce"].ToString();
                string NatureContrat = rez["NatureContrat"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string A1100 = rez["A1100"].ToString();
                string PRESC1J = rez["PRESC1J"].ToString();
                string PRESC2H = rez["PRESC2H"].ToString();
                string A1000 = rez["A1000"].ToString();
                string A2200 = rez["A2200"].ToString();
                string A2100 = rez["A2100"].ToString();
                string A4170 = rez["A4170"].ToString();
                string A4173 = rez["A4173"].ToString();
                string A4174 = rez["A4174"].ToString();
                string A3710 = rez["A3710"].ToString();
                string A4171 = rez["A4171"].ToString();
                string A4110 = rez["A4110"].ToString();
                string A4111 = rez["A4111"].ToString();
                string A4113 = rez["A4113"].ToString();
                string A4130 = rez["A4130"].ToString();
                string A4140 = rez["A4140"].ToString();
                string A4217 = rez["A4217"].ToString();
                string A4218 = rez["A4218"].ToString();
                string A4220 = rez["A4220"].ToString();
                string A4221 = rez["A4221"].ToString();
                string A4222 = rez["A4222"].ToString();
                string A4340 = rez["A4340"].ToString();
                string A4711 = rez["A4711"].ToString();
                string A4712 = rez["A4712"].ToString();
                string A4715 = rez["A4715"].ToString();
                string A4717 = rez["A4717"].ToString();
                string CL25 = rez["CL25"].ToString();
                string CL24 = rez["CL24"].ToString();
                string A7113 = rez["A7113"].ToString();
                string A7114 = rez["A7114"].ToString();
                string A7115 = rez["A7115"].ToString();
                string A4811 = rez["A4811"].ToString();
                string A4812 = rez["A4812"].ToString();
                string A4813 = rez["A4813"].ToString();
                string A4815 = rez["A4815"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string A9112 = rez["A9112"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A6210 = rez["A6210"].ToString();
                string A9121 = rez["A9121"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A6614 = rez["A6614"].ToString();
                string A6410 = rez["A6410"].ToString();
                string A6850 = rez["A6850"].ToString();
                string A6520 = rez["A6520"].ToString();
                string A6430 = rez["A6430"].ToString();
                string A6490 = rez["A6490"].ToString();
                string A6830 = rez["A6830"].ToString();
                string A6831 = rez["A6831"].ToString();
                string A6840 = rez["A6840"].ToString();
                string A8800 = rez["A8800"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A6110A = rez["A61109"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A7115A = rez["A71159"].ToString();
                string A7116 = rez["A7116"].ToString();
                string RUBCOTEPY = rez["RUB_COTEPY"].ToString();
                string MOIS_PAIE = rez["MOIS_PAIE"].ToString();

                worksheet = EcrireSGNA002(cpt, worksheet, Mat, Nom, Prenom, DateEmbauche, DateDepart, CodeSce, IntituleSce, NatureContrat,
                    Emploi, A1100, PRESC1J, PRESC2H, A1000, A2200, A2100, A4170, A4173, A4174, A3710, A4171, A4110, A4111, A4113, A4130,
                    A4140, A4217, A4218, A4220, A4221, A4222, A4340, A4711, A4712, A4715, A4717, CL25, CL24, A7113, A7114, A7115, A4811,
                    A4812, A4813,A4815, BRUT, A9112, A6110, A6210, A9121, A6310, A6311, A6614, A6410,A6850, A6520, A6430, A6490, A6830,
                    A6831, A6840,A8800, RUB_DEDSA, NETPAIE, A6110A, A6150, A7115A,A7116, RUBCOTEPY, MOIS_PAIE);
                cpt = cpt + 1;
            }

            return worksheet;
        }

        public Worksheet EcrireSGN(int cpt, Worksheet worksheet, string Mat, string Nom, string Prenom, string DateEmbauche, string DateDepart,
            string CodeSce, string IntituleSce, string NatureContrat, string Emploi, string A1100, string CL01, string A1200,string A2200,string A2100,
            string A2311, string A2320, string A3318, string A3110,string A4311,string A4330, string A3111, string A3210, string A3317, string A3319, string A4100, string A4171,
            string A4130, string A4223, string A4380, string A4711,string A4712, string A4717,string A4226, string A5100, string A7600, string A4811, string A4812, string A4813,
            string BRUT, string AJUSCOT, string A9112, string A6110, string A6120, string A6210, string A6230, string A9121, string A6310, string A6311,
            string A6614, string A6410, string A6520, string A6510, string A6530, string A6890,string A8000, string RUB_DEDSA, string NETPAIE, string A6110R,
            string A6120R, string A6150, string A6210R, string A6230R, string A6320, string A6330, string RUB_COTEPY, string MOIS_PAIE)
        {
            worksheet.Cells[cpt + 11, 0] = new Cell(Mat);
            worksheet.Cells[cpt + 11, 1] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 2] = new Cell(Prenom);
            worksheet.Cells[cpt + 11, 3] = new Cell(DateEmbauche);
            worksheet.Cells[cpt + 11, 4] = new Cell(DateDepart);
            worksheet.Cells[cpt + 11, 5] = new Cell(CodeSce);
            worksheet.Cells[cpt + 11, 6] = new Cell(IntituleSce);
            worksheet.Cells[cpt + 11, 7] = new Cell(NatureContrat);
            worksheet.Cells[cpt + 11, 8] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 9] = new Cell(A1100);
            worksheet.Cells[cpt + 11, 10] = new Cell(CL01);
            worksheet.Cells[cpt + 11, 11] = new Cell(A1200);
            worksheet.Cells[cpt + 11, 12] = new Cell(A2200);
            worksheet.Cells[cpt + 11, 13] = new Cell(A2100);
            worksheet.Cells[cpt + 11, 14] = new Cell(A2311);
            worksheet.Cells[cpt + 11, 15] = new Cell(A2320);
            worksheet.Cells[cpt + 11, 16] = new Cell(A3318);
            worksheet.Cells[cpt + 11, 17] = new Cell(A3110);
            worksheet.Cells[cpt + 11, 18] = new Cell(A4311);
            worksheet.Cells[cpt + 11, 19] = new Cell(A4330);
            worksheet.Cells[cpt + 11, 20] = new Cell(A3111);
            worksheet.Cells[cpt + 11, 21] = new Cell(A3210);
            worksheet.Cells[cpt + 11, 22] = new Cell(A3317);
            worksheet.Cells[cpt + 11, 23] = new Cell(A3319);
            worksheet.Cells[cpt + 11, 24] = new Cell(A4100);
            worksheet.Cells[cpt + 11, 25] = new Cell(A4171);
            worksheet.Cells[cpt + 11, 26] = new Cell(A4130);
            worksheet.Cells[cpt + 11, 27] = new Cell(A4223);
            worksheet.Cells[cpt + 11, 28] = new Cell(A4380);
            worksheet.Cells[cpt + 11, 29] = new Cell(A4711);
            worksheet.Cells[cpt + 11, 30] = new Cell(A4712);
            worksheet.Cells[cpt + 11, 31] = new Cell(A4717);
            worksheet.Cells[cpt + 11, 32] = new Cell(A4226);
            worksheet.Cells[cpt + 11, 33] = new Cell(A5100);
            worksheet.Cells[cpt + 11, 34] = new Cell(A7600);
            worksheet.Cells[cpt + 11, 35] = new Cell(A4811);
            worksheet.Cells[cpt + 11, 36] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 37] = new Cell(A4813);
            worksheet.Cells[cpt + 11, 38] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 39] = new Cell(AJUSCOT);
            worksheet.Cells[cpt + 11, 40] = new Cell(A9112);
            worksheet.Cells[cpt + 11, 41] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 42] = new Cell(A6120);
            worksheet.Cells[cpt + 11, 43] = new Cell(A6210);
            worksheet.Cells[cpt + 11, 44] = new Cell(A9121);
            worksheet.Cells[cpt + 11, 45] = new Cell(A6230);
            worksheet.Cells[cpt + 11, 46] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 47] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 48] = new Cell(A6614);
            worksheet.Cells[cpt + 11, 49] = new Cell(A6410);
            worksheet.Cells[cpt + 11, 50] = new Cell(A6520);
            worksheet.Cells[cpt + 11, 51] = new Cell(A6510);
            worksheet.Cells[cpt + 11, 52] = new Cell(A6530);
            worksheet.Cells[cpt + 11, 53] = new Cell(A6890);
            worksheet.Cells[cpt + 11, 54] = new Cell(A8000);
            worksheet.Cells[cpt + 11, 55] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 56] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 57] = new Cell(A6110R);
            worksheet.Cells[cpt + 11, 58] = new Cell(A6120R);
            worksheet.Cells[cpt + 11, 59] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 60] = new Cell(A6210R);
            worksheet.Cells[cpt + 11, 61] = new Cell(A6230R);
            worksheet.Cells[cpt + 11, 62] = new Cell(A6320);
            worksheet.Cells[cpt + 11, 63] = new Cell(A6330);
            worksheet.Cells[cpt + 11, 64] = new Cell(RUB_COTEPY);
            worksheet.Cells[cpt + 11, 65] = new Cell(MOIS_PAIE);

            return worksheet;
        }

        public Worksheet EnteteSVRLigne12A003(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA003(refSoc, MoisPe, AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Mat = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string CodeSce = rez["CodeSce"].ToString();
                string IntituleSce = rez["IntituleSce"].ToString();
                string NatureContrat = rez["NatureContrat"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string SALBASE = rez["A1100"].ToString();
                string CL01 = rez["CL01"].ToString();
                string A1200 = rez["A1200"].ToString();
                string A2200 = rez["A2200"].ToString();
                string A2100 = rez["A2100"].ToString();
                string A2360 = rez["A2360"].ToString();
                string A2330 = rez["A2330"].ToString();
                string A3412 = rez["A3412"].ToString();
                string A3317 = rez["A3317"].ToString();
                string A3712 = rez["A3712"].ToString();
                string A3710 = rez["A3710"].ToString();
                string A3711 = rez["A3711"].ToString();
                string A3713 = rez["A3713"].ToString();
                string A4112 = rez["A4112"].ToString();
                string A4113 = rez["A4113"].ToString();
                string A4120 = rez["A4120"].ToString();
                string A4222 = rez["A4222"].ToString();
                string A4216 = rez["A4216"].ToString();
                string A4361 = rez["A4361"].ToString();
                string A4215 = rez["A4215"].ToString();
                string A4320 = rez["A4320"].ToString();
                string A4370 = rez["A4370"].ToString();
                string A4430 = rez["A4430"].ToString();
                string A4812 = rez["A4812"].ToString();
                string A5100 = rez["A5100"].ToString();
                string A7113 = rez["A7113"].ToString();
                string A4171 = rez["A4171"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string AjusCot = rez["AjusCot"].ToString();
                string A9112 = rez["A9112"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A6120 = rez["A6120"].ToString();
                string A9310 = rez["A9310"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A6340 = rez["A6340"].ToString();
                string A6210 = rez["A6210"].ToString();
                string A6410 = rez["A6410"].ToString();
                string A6430 = rez["A6430"].ToString();
                string A6520 = rez["A6520"].ToString();
                string A6510 = rez["A6510"].ToString();
                string A6530 = rez["A6530"].ToString();
                string A6830 = rez["A6830"].ToString();
                string A6490 = rez["A6490"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A6110A = rez["A61109"].ToString();
                string A6120A = rez["A61209"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A6210A = rez["A62109"].ToString();
                string RUB_COTEPY = rez["RUB_COTEPY"].ToString();
                string MoisP = rez["MoisPe"].ToString();
                string A4380 = rez["A4380"].ToString(); ;

                worksheet = EcrireSGNA003(cpt, worksheet, Mat, Nom, Prenom, DateEmbauche, DateDepart, CodeSce, IntituleSce, NatureContrat,
                    Emploi,SALBASE, CL01, A1200, A2200, A2100, A2360, A2330, A3412, A3317, A3712, A3710, A3711, A3713, A4112, A4113, A4120, A4222,
                   A4216,A4361, A4215, A4320, A4370, A4430,A4812,A5100, A7113, A4171, BRUT, AjusCot, A9112, A6110, A6120, A9310, A6310, A6311, A6340, A6210,
                   A6410, A6430, A6520, A6510, A6530, A6830,A6490, RUB_DEDSA, NETPAIE, A6110A, A6120A, A6150, A6210A, RUB_COTEPY, MoisP,A4380);
                cpt = cpt + 1;
            }

            return worksheet;
        }

        public Worksheet EnteteSVRLigne12A004(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA004(refSoc, MoisPe, AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Mat = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string CodeSce = rez["CodeSce"].ToString();
                string IntituleSce = rez["IntituleSce"].ToString();
                string NatureContrat = rez["NatureContrat"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string SALBASE = rez["A1100"].ToString();
                string CL01 = rez["CL01"].ToString();
                string A1100 = rez["A1200"].ToString();
                string A2200 = rez["A2200"].ToString();
                string A2100 = rez["A2100"].ToString();
                string A3411 = rez["A3411"].ToString();
                string A3410 = rez["A3410"].ToString();
                string A3412 = rez["A3412"].ToString();
                string A3317 = rez["A3317"].ToString();
                string A3712 = rez["A3712"].ToString();
                string A3710 = rez["A3710"].ToString();
                string A3711 = rez["A3711"].ToString();
                string A3713 = rez["A3713"].ToString();
                string A4112 = rez["A4112"].ToString();
                string A4113 = rez["A4113"].ToString();
                string A4120 = rez["A4120"].ToString();
                string A4222 = rez["A4222"].ToString();
                string A4216 = rez["A4216"].ToString();
                string A4361 = rez["A4361"].ToString();
                string A4215 = rez["A4215"].ToString();
                string A4320 = rez["A4320"].ToString();
                string A4370 = rez["A4370"].ToString();
                string A4430 = rez["A4430"].ToString();
                string A4171 = rez["A4171"].ToString();
                string A4380 = rez["A4380"].ToString();
                string A4811 = rez["A4811"].ToString();
                string A7113 = rez["A7113"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string AjusCot = rez["AjusCot"].ToString();
                string A9112 = rez["A9112"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A6210 = rez["A6210"].ToString();
                string A9310 = rez["A9310"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A6340 = rez["A6340"].ToString();
                string A6210A = rez["A62109"].ToString();
                string A6410 = rez["A6410"].ToString();
                string A6430 = rez["A6430"].ToString();
                string A6520 = rez["A6520"].ToString();
                string A6510 = rez["A6510"].ToString();
                string A6530 = rez["A6530"].ToString();
                string A6830 = rez["A6830"].ToString();
                string A6490 = rez["A6490"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A6110A = rez["A61109"].ToString();
                string A6120 = rez["A6120"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A6210AA = rez["A621099"].ToString();
                string RUB_COTEPY = rez["RUB_COTEPY"].ToString();
                string MoisP = rez["MoisPe"].ToString();
                string A4812 = rez["A4812"].ToString();

                worksheet = EcrireSGNA004(cpt, worksheet, Mat, Nom, Prenom, DateEmbauche, DateDepart, CodeSce, IntituleSce, NatureContrat,
                    Emploi, SALBASE, CL01, A1100, A2200, A2100, A3411, A3410, A3412, A3317, A3712, A3710, A3711, A3713, A4112, A4113, A4120,A4216, A4222,
                   A4361, A4215, A4320, A4370, A4430,A4171,A4380,A4811, A7113,BRUT, AjusCot, A9112, A6110, A6210, A9310, A6310, A6311, A6340, A6210A,
                   A6410, A6430, A6520, A6510, A6530, A6830,A6490, RUB_DEDSA, NETPAIE, A6110A, A6120, A6150, A6210AA, RUB_COTEPY, MoisP,A4812);
                cpt = cpt + 1;
            }

            return worksheet;
        }

     



        public Worksheet EnteteSVRLigne12A010(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA010(refSoc, MoisPe, AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string IdOracle = rez["IdOracle"].ToString();
                string Matricule = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string NatureContrat = rez["NatureContrat"].ToString();
                string CodeDept = rez["CodeDept"].ToString();
                string Dept = rez["IntituleDept"].ToString();
                string HorBaseSal = rez["HorBaseSal"].ToString();
                string BullMod = rez["BullMod"].ToString();
                string Categorie = rez["Categorie"].ToString();
                string A1100 = rez["A1100"].ToString();
                string A1300 = rez["A1300"].ToString();
                string A2100 = rez["A2100"].ToString();
                string A2200 = rez["A2200"].ToString();
                string A3310 = rez["A3310"].ToString();
                string ABS_CONGNR = rez["ABS_CONGNR"].ToString();
                string AbsNonRem = rez["A4174"].ToString();
                string RUB_ARR = rez["RUB_ARR"].ToString();
                string A4373 = rez["A4373"].ToString();
                string HS01 = rez["HS01"].ToString();
                string A4110 = rez["A4110"].ToString();
                string HS02 = rez["HS02"].ToString();
                string A4111 = rez["A4111"].ToString();
                string HS03 = rez["HS03"].ToString();
                string A4113 = rez["A4113"].ToString();
                string HS06 = rez["HS06"].ToString();
                string A4112 = rez["A4112"].ToString();
                string HS04 = rez["HS04"].ToString();
                string A4120 = rez["A4120"].ToString();
                string HS07 = rez["HS07"].ToString();
                string A4372 = rez["A4372"].ToString();
                string HS05 = rez["HS05"].ToString();
                string A4114= rez["A4114"].ToString();
                string HA04 = rez["HA04"].ToString();
                string HA06 = rez["HA06"].ToString();
                string RUB_JCHOM = rez["RUB_JCHOM"].ToString();
                string HA07 = rez["HA07"].ToString();
                string RUB_TRP = rez["A4176"].ToString();
                string HA09 = rez["HA09"].ToString();
                string A5500 = rez["A5500"].ToString();
                string A4222 = rez["A4222"].ToString();
                string A4371 = rez["A4371"].ToString();
                string A4130 = rez["A4130"].ToString();
                string A4135 = rez["A4135"].ToString();
                string A2361 = rez["A2361"].ToString();
                string A4816 = rez["A4816"].ToString();
                string A4810 = rez["A4810"].ToString();
                string HA01 = rez["HA01"].ToString();
                string A4811 = rez["A4811"].ToString();
                string A4812 = rez["A4812"].ToString();
                string A5100 = rez["A5100"].ToString();
                string A5400 = rez["A5400"].ToString();
                string A5900 = rez["A5900"].ToString();
                string A5200 = rez["A5200"].ToString();
                string A6510 = rez["A6510"].ToString();
                string A3309 = rez["A3309"].ToString();
                string A4817 = rez["A4817"].ToString();
                string A4818 = rez["A4818"].ToString();
                string A4313 = rez["A4313"].ToString();
                string A4312 = rez["A4312"].ToString();
                string TESTSTC = rez["TESTSTC"].ToString();
                string A4800 = rez["A4800"].ToString();
                string PRES8 = rez["PRES8"].ToString();
                string PRES2 = rez["PRES2"].ToString();
                string A7115 = rez["A7115"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A9310 = rez["A9310"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A6490 = rez["A6490"].ToString();
                string A6491 = rez["A6491"].ToString();
                string A6492 = rez["A6492"].ToString();
                string A6840 = rez["A6840"].ToString();
                string A6610 = rez["A6610"].ToString();
                string A6611 = rez["A6611"].ToString();
                string A6613 = rez["A6613"].ToString();
                string A6612 = rez["A6612"].ToString();
                string A6410 = rez["A6410"].ToString();
                string A6511 = rez["A6511"].ToString();
                string A6614 = rez["A6614"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A6110A = rez["A61109"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A7115A = rez["A71159"].ToString();
                string RUB_COTEPY = rez["RUB_COTEPY"].ToString();
                string ModPe = rez["ModPe"].ToString();
                string MoisP = rez["MoisPe"].ToString();
                string A4374 = rez["A4374"].ToString();

                worksheet = EcrireSGNA010(cpt, worksheet, IdOracle,Matricule,Nom,Prenom,DateEmbauche,DateDepart,Emploi,NatureContrat,
                    CodeDept,Dept,HorBaseSal,BullMod,Categorie,A1100,A1300,A2100,A2200,A3310,ABS_CONGNR,AbsNonRem,RUB_ARR,A4373,HS01,
                    A4110,HS02,A4111,HS03,A4113,HS06,A4112,HS04,A4120,HS07,A4372,HS05,A4114,HA04,HA06,RUB_JCHOM,HA07,RUB_TRP,HA09,A5500,
                    A4222,A4371,A4130,A4135,A2361,A4816,A4810,HA01,A4811,A4812,A5100,A5400,A5900,A5200,A6510,A3309,A4817,A4818,A4313,A4312,
                    TESTSTC,A4800,PRES8,PRES2,A7115,BRUT,A6110,A9310,A6310,A6311,A6490,A6491,A6492,A6840,A6610,A6611,A6613,A6612,A6410,
                    A6511,A6614,RUB_DEDSA,NETPAIE,A6110A,A6150,A7115A,RUB_COTEPY,ModPe,MoisP,A4374);
                cpt = cpt + 1;
            }

            return worksheet;
        }

    


        public Worksheet EnteteSVRLigne12A011(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA011(refSoc, MoisPe, AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Matricule = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string IntituleDept = rez["IntituleDept"].ToString();
                string IntituleSce = rez["IntituleSce"].ToString();
                string IntituleCat = rez["IntituleCat"].ToString();
                string NatContrat = rez["NatContrat"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string SalBaseSal = rez["A1100"].ToString();
                string CL01 = rez["CL01"].ToString();
                string A1200 = rez["A1200"].ToString();
                string A2100 = rez["A2100"].ToString();
                string A2200 = rez["A2200"].ToString();
                string A2311 = rez["A2311"].ToString();
                string A3111 = rez["A3111"].ToString();
                string A3413 = rez["A3413"].ToString();
                string A3410 = rez["A3410"].ToString();
                string A4351 = rez["A4351"].ToString();
                string A4220 = rez["A4220"].ToString();
                string A4211 = rez["A4211"].ToString();
                string A4369 = rez["A4369"].ToString();
                string A4360 = rez["A4360"].ToString();
                string A4380 = rez["A4380"].ToString();
                string A4717 = rez["A4717"].ToString();
                string A4716 = rez["A4716"].ToString();
                string A4110 = rez["A4110"].ToString();
                string A4111 = rez["A4111"].ToString();
                string A4113 = rez["A4113"].ToString();
                string A4114 = rez["A4114"].ToString();
                string A4812 = rez["A4812"].ToString();
                string A4171 = rez["A4171"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A9121 = rez["A9121"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string TotDecEmp = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A6110A = rez["A61109"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A6320 = rez["A6320"].ToString();
                string A6330 = rez["A6330"].ToString();
                string RUB_EPYC = rez["RUB_COTEPY"].ToString();
                string TotCostLab = rez["TotCostLab"].ToString();
                string MoisP = rez["MoisPe"].ToString();

                worksheet = EcrireSGNA011(cpt, worksheet, Matricule, Nom, Prenom, DateEmbauche, DateDepart, IntituleDept, IntituleSce,
                 IntituleCat, NatContrat, Emploi, SalBaseSal, CL01, A1200, A2100, A2200, A2311, A3111, A3413,A3410, A4351, A4220, A4211, A4369,
                 A4360, A4380, A4717, A4716, A4110, A4111, A4113, A4114,A4812,A4171, BRUT, A6110, A9121, A6310, A6311, TotDecEmp, NETPAIE, A6110A, A6150,
                 A6320, A6330, RUB_EPYC, TotCostLab, MoisP);
                cpt = cpt + 1;
            }

            return worksheet;
        }

        public string VerifSiVide(string Val)
        {
            string ValF=string.Empty;
            if(Val=="")
            {
                ValF="";
            }
            else
            {
                ValF=Val;
            }
            return ValF;
        }

        public Worksheet EnteteSVRLigne12A013(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA013(refSoc, MoisPe, AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string IdOracle = VerifSiVide(rez["IdOracle"].ToString());
                string Matricule = VerifSiVide(rez["Matricule"].ToString());
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string NatureContrat = rez["NatureContrat"].ToString();
                string CodeDept = rez["CodeDept"].ToString();
                string Dept = rez["IntituleDept"].ToString();
                string HorBaseSal = rez["HorBaseSal"].ToString();
                string BullMod = rez["BullMod"].ToString();
                string Categorie = rez["Categorie"].ToString();
                string ModPaie = rez["ModPaie"].ToString();
                string A1100 = rez["A1100"].ToString();
                string A1300 = rez["A1300"].ToString();
                string A1310 = rez["A1310"].ToString();
                string A2100 = rez["A2100"].ToString();
                string A2200 = rez["A2200"].ToString();
                string A3310 = rez["A3310"].ToString();
                string A3110 = rez["A3110"].ToString();
                string ABS_CONGNR = rez["ABS_CONGNR"].ToString();
                string A4174 = rez["A4174"].ToString();
                string HS01 = rez["HS01"].ToString();
                string A4110 = rez["A4110"].ToString();
                string HS02 = rez["HS02"].ToString();
                string A4111 = rez["A4111"].ToString();
                string HS03 = rez["HS03"].ToString();
                string A4113 = rez["A4113"].ToString();
                string HS04 = rez["HS04"].ToString();
                string A4120 = rez["A4120"].ToString();
                string HS07 = rez["HS07"].ToString();
                string A4372 = rez["A4372"].ToString();
                string A4374 = rez["A4374"].ToString();
                string HS05 = rez["HS05"].ToString();
                string A4114 = rez["A4114"].ToString();
                string HA04 = rez["HA04"].ToString();
                string HA06 = rez["HA06"].ToString();
                string A4175 = rez["A4175"].ToString();
                string HA07 = rez["HA07"].ToString();
                string A4176 = rez["A4176"].ToString();
                string A4130 = rez["A4130"].ToString();
                string A4135 = rez["A4135"].ToString();
                string A2361 = rez["A2361"].ToString();
                string A4222 = rez["A4222"].ToString();
                string A3309 = rez["A3309"].ToString();
                string A4373 = rez["A4373"].ToString();
                string A4371 = rez["A4371"].ToString();
                string A4817 = rez["A4817"].ToString();
                string A4818 = rez["A4818"].ToString();
                string A4313 = rez["A4313"].ToString();
                string A4311 = rez["A4311"].ToString();
                string HA09 = rez["HA09"].ToString();
                string A5500 = rez["A5500"].ToString();
                string A5100 = rez["A5100"].ToString();
                string A5400 = rez["A5400"].ToString();
                string A5200 = rez["A5200"].ToString();
                string A5900 = rez["A5900"].ToString();
                string HA01 = rez["HA01"].ToString();
                string A4816 = rez["A4816"].ToString();
                string A4810 = rez["A4810"].ToString();
                string A4811 = rez["A4811"].ToString();
                string A4812 = rez["A4812"].ToString();
                string TESTSTC = rez["TESTSTC"].ToString();
                string A4800 = rez["A4800"].ToString();
                string PRES8 = rez["PRES8"].ToString();
                string A7113 = rez["A7113"].ToString();
                string CL23 = rez["CL23"].ToString();
                string A7116 = rez["A7116"].ToString();
                string A7115 = rez["A7115"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A9310 = rez["A9310"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A6614 = rez["A6614"].ToString();
                string A6410 = rez["A6410"].ToString();
                string A6490 = rez["A6490"].ToString();
                string A6491 = rez["A6491"].ToString();
                string A6492 = rez["A6492"].ToString();
                string A6511 = rez["A6511"].ToString();
                string A6430 = rez["A6430"].ToString();
                string A6431 = rez["A6431"].ToString();
                string A6840 = rez["A6840"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A61109 = rez["A61109"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A71159 = rez["A71159"].ToString();
                string RUB_COTEPY = rez["RUB_COTEPY"].ToString();
                string MoisP = rez["MOIS_PAIE"].ToString();
                worksheet = EcrireSGNA013(cpt, worksheet, IdOracle, Matricule, Nom, Prenom, DateEmbauche,
                       DateDepart, Emploi, NatureContrat, CodeDept, Dept, HorBaseSal, BullMod,
                       Categorie, ModPaie, A1100, A1300, A1310, A2100, A2200, A3310, A3110, ABS_CONGNR, A4174, HS01, A4110, HS02, A4111, HS03, A4113, HS04, A4120,
                        HS07, A4372,A4374, HS05, A4114, HA04, HA06, A4175, HA07, A4176, A4130, A4135, A2361, A4222, A3309, A4373, A4371, A4817, A4818, A4313,
                        A4311, HA09, A5500, A5100, A5400, A5200, A5900, HA01, A4816, A4810, A4811, A4812, TESTSTC, A4800, PRES8, A7113, CL23, A7116, A7115,
                        BRUT, A6110, A9310, A6310, A6311, A6614, A6410, A6490, A6491, A6492, A6511, A6430, A6431, A6840, RUB_DEDSA, NETPAIE, A61109, A6150, A71159,
                        RUB_COTEPY, MoisP);
                cpt = cpt + 1;
            }

            return worksheet;
        }

        public Worksheet EnteteSVRLigne12A016(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA016(refSoc, MoisPe, AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Matricule = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDept"].ToString();
                string CodeDept = rez["CodeDept"].ToString();
                string IntituleDept = rez["IntituleDept"].ToString();
                string NatureContrat = rez["NatureContrat"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string SALBASE = rez["SALBASE"].ToString();
                string HORAIRE = rez["HORAIRE"].ToString();
                string A1000 = rez["A1000"].ToString();
                string A2100 = rez["A2100"].ToString();
                string A2200 = rez["A2200"].ToString();
                string A3413 = rez["A3413"].ToString();
                string A3414 = rez["A3414"].ToString();
                string A3415 = rez["A3415"].ToString();
                string A3418 = rez["A3418"].ToString();
                string A3416 = rez["A3416"].ToString();
                string A3317 = rez["A3317"].ToString();
                string A3510 = rez["A3510"].ToString();
                string A3520 = rez["A3520"].ToString();
                string A3530 = rez["A3530"].ToString();
                string A3417 = rez["A3417"].ToString();
                string A5910 = rez["A5910"].ToString();
                string A3318 = rez["A3318"].ToString();
                string A3610 = rez["A3610"].ToString();
                string A3620 = rez["A3620"].ToString();
                string A3630 = rez["A3630"].ToString();
                string A3640 = rez["A3640"].ToString();
                string A3651 = rez["A3651"].ToString();
                string A3652 = rez["A3652"].ToString();
                string A3650 = rez["A3650"].ToString();
                string A3660 = rez["A3660"].ToString();
                string A3661 = rez["A3661"].ToString();
                string A3670 = rez["A3670"].ToString();
                string A3680 = rez["A3680"].ToString();
                string A3681 = rez["A3681"].ToString();
                string A3682 = rez["A3682"].ToString();
                string A3683 = rez["A3683"].ToString();
                string A3684 = rez["A3684"].ToString();
                string A3685 = rez["A3685"].ToString();
                string BRUTFX = rez["BRUTFX"].ToString();
                string A3686 = rez["A3686"].ToString();
                string A3690 = rez["A3690"].ToString();
                string CL31 = rez["CL31"].ToString();
                string A4460 = rez["A4460"].ToString();
                string CL32 = rez["CL32"].ToString();
                string A4461 = rez["A4461"].ToString();
                string CL33 = rez["CL33"].ToString();
                string A4462 = rez["A4462"].ToString();
                string CL34 = rez["CL34"].ToString();
                string A4463 = rez["A4463"].ToString();
                string CL35 = rez["CL35"].ToString();
                string A3110 = rez["A3110"].ToString();
                string CL37 = rez["CL37"].ToString();
                string A4292 = rez["A4292"].ToString();
                string A4293 = rez["A4293"].ToString();
                string CL39 = rez["CL39"].ToString();
                string A4294 = rez["A4294"].ToString();
                string HC01 = rez["HC01"].ToString();
                string A4295 = rez["A4295"].ToString();
                string CL40 = rez["CL40"].ToString();
                string  A4296= rez["A4296"].ToString();
                string A4491 = rez["A4491"].ToString();
                string A4162 = rez["A4162"].ToString();
                string A4391 = rez["A4391"].ToString();
                string A4381 = rez["A4381"].ToString();
                string A4840 = rez["A4840"].ToString();
                string A4382 = rez["A4382"].ToString();
                string A4850 = rez["A4850"].ToString();
                string A4383 = rez["A4383"].ToString();
                string A4860 = rez["A4860"].ToString();
                string A4718 = rez["A4718"].ToString();
                string A4717 = rez["A4717"].ToString();
                string A4713 = rez["A4713"].ToString();
                string A4716 = rez["A4716"].ToString();
                string HS01 = rez["HS01"].ToString();
                string A4111 = rez["A4111"].ToString();
                string HS02 = rez["HS02"].ToString();
                string A4112 = rez["A4112"].ToString();
                string HS03 = rez["HS03"].ToString();
                string A4113 = rez["A4113"].ToString();
                string HS04 = rez["HS04"].ToString();
                string A4114 = rez["A4114"].ToString();
                string HS05 = rez["HS05"].ToString();
                string A4115 = rez["A4115"].ToString();
                string  HS06= rez["HS06"].ToString();
                string A4116 = rez["A4116"].ToString();
                string HS07 = rez["HS07"].ToString();
                string A4117 = rez["A4117"].ToString();
                string A4172 = rez["A4172"].ToString();
                string A4173 = rez["A4173"].ToString();
                string HC02 = rez["HC02"].ToString();
                string A4464 = rez["A4464"].ToString();
                string A5100 = rez["A5100"].ToString();
                string A5700 = rez["A5700"].ToString();
                string A5702 = rez["A5702"].ToString();
                string A5220 = rez["A5220"].ToString();
                string A5400 = rez["A5400"].ToString();
                string A5800 = rez["A5800"].ToString();
                string A5900 = rez["A5900"].ToString();
                string A7111 = rez["A7111"].ToString();
                string A7114 = rez["A7114"].ToString();
                string A7112 = rez["A7112"].ToString();
                string A7116 = rez["A7116"].ToString();
                string A7117 = rez["A7117"].ToString();
                string A7113 = rez["A7113"].ToString();
                string A5920 = rez["A5920"].ToString();
                string A59209 = rez["A59209"].ToString();
                string A7115 = rez["A7115"].ToString();
                string A7600 = rez["A7600"].ToString();
                string A7670 = rez["A7670"].ToString();
                string A3910 = rez["A3910"].ToString();
                string A4170 = rez["A4170"].ToString();
                string A4811 = rez["A4811"].ToString();
                string A4813 = rez["A4813"].ToString();
                string A4812 = rez["A4812"].ToString();
                string A4830 = rez["A4830"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string A9110 = rez["A9110"].ToString();
                string A9112 = rez["A9112"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A6120 = rez["A6120"].ToString();
                string A9121 = rez["A9121"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6312 = rez["A6312"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A6462 = rez["A6462"].ToString();
                string A6511 = rez["A6511"].ToString();
                string A6430 = rez["A6430"].ToString();
                string A6465 = rez["A6465"].ToString();
                string A6467 = rez["A6467"].ToString();
                string A6840 = rez["A6840"].ToString();
                string A6810 = rez["A6810"].ToString();
                string A6830 = rez["A6830"].ToString();
                string A6860 = rez["A6860"].ToString();
                string A6463 = rez["A6463"].ToString();
                string A6610 = rez["A6610"].ToString();
                string A6870 = rez["A6870"].ToString();
                string A8140 = rez["A8140"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A61109 = rez["A61109"].ToString();
                string A61209 = rez["A61209"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A6210 = rez["A6210"].ToString();
                string TFP = rez["TFP"].ToString();
                string RUB_COTEPY = rez["RUB_COTEPY"].ToString();
                string RD_NET = rez["RD_NET"].ToString();
                string MoisPaie = rez["MoisPe"].ToString();

           worksheet = EcrireSGNA016(cpt, worksheet, Matricule, Nom, Prenom, DateEmbauche, DateDepart, CodeDept, IntituleDept,
           NatureContrat,  Emploi,  SALBASE,  HORAIRE,  A1000,  A2100,  A2200,  A3413,  A3414,  A3415,  A3418,
           A3416,  A3317,  A3510,  A3520,  A3530,  A3417,  A5910,  A3318,  A3610,  A3620,  A3630,  A3640,  A3651,
           A3652,  A3650,  A3660,  A3661,  A3670,  A3680,  A3681,  A3682,  A3683,  A3684,  A3685,  BRUTFX,  A3686,
           A3690,  CL31,  A4460,  CL32,  A4461,  CL33,  A4462,  CL34,  A4463,  CL35,  A3110,
           CL37,  A4292,  A4293,  CL39,  A4294,  HC01,  A4295,  CL40,  A4296,  A4491,  A4162,  A4391,  A4381,
           A4840,  A4382,  A4850,  A4383,  A4860,  A4718,  A4717,  A4713,  A4716,  HS01,  A4111,  HS02,  A4112,
           HS03,  A4113,  HS04,  A4114,  HS05,  A4115,  HS06,  A4116,  HS07,  A4117,  A4172,  A4173,
           HC02,  A4464,  A5100,  A5700,  A5702,  A5220,  A5400,  A5800,  A5900,  A7111,  A7114,  A7112,
           A7116,  A7117,  A7113,  A5920,  A59209,  A7115,  A7600,  A7670,  A3910,  A4170,  A4811,  A4813,  A4812,
           A4830,  BRUT,  A9110,  A9112,  A6110,  A6120,  A9121,  A6310,  A6312,  A6311,  A6462,  A6511,  A6430,
           A6465,  A6467,  A6840,  A6810,  A6830,  A6860,  A6463,  A6610,  A6870,  A8140,  RUB_DEDSA,  NETPAIE,
           A61109,  A61209,  A6150,  A6210,  TFP,  RUB_COTEPY,  RD_NET, MoisPaie);
                cpt = cpt + 1;
            }

            return worksheet;
        }

        public Worksheet EnteteSVRLigne12A021(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA021(refSoc, MoisPe, AnnePe);
            int index = 0;

            foreach (System.Data.DataRow rez in dt.Rows)
            {
                index = 1;
                string matricule = rez[index++].ToString();
                string Datedembauchesociete = rez[index++].ToString();
                string nom = rez[index++].ToString();
                string prenom = rez[index++].ToString();
                string CodeCostcenter = rez[index++].ToString();

                string Datededepartsociete = rez[index++].ToString();
                string Intituledepartement = rez[index++].ToString();
                string Emploioccupe = rez[index++].ToString();

                string Intitulecostcenter = rez[index++].ToString();
                string Intituledelanatureducontrat = rez[index++].ToString();
              
                string Salairedebase = rez[index++].ToString();
                string Nbjoursdepresence = rez[index++].ToString();
                string SalairedebaseJtrvaille = rez[index++].ToString();
                string Inddepresence = rez[index++].ToString();
                string InddeTransport = rez[index++].ToString();
                string PrimeComplementairedePresence = rez[index++].ToString();
                string IndemnitedeVoiture = rez[index++].ToString();
                string IndFixederepresentation = rez[index++].ToString();
                string Indemnitedexpatriation = rez[index++].ToString();
                string HS125HS0 = rez[index++].ToString();
                string ValeurHS0 = rez[index++].ToString();
                string HS1502 = rez[index++].ToString();
                string ValeurHS02 = rez[index++].ToString();
                string HS175HS06 = rez[index++].ToString();
                string ValeurHS06 = rez[index++].ToString();
                string HS200HS03 = rez[index++].ToString();
                string ValeurHS03 = rez[index++].ToString();
                string HsupplementairenuitHS04 = rez[index++].ToString();
                string valeurHsuplementairenuitHS04 = rez[index++].ToString();
                string primeLVP = rez[index++].ToString();
                string PrimeSIP = rez[index++].ToString();
                string PrimeAOS = rez[index++].ToString();
                string Bonus = rez[index++].ToString();
                string Compensationsupport = rez[index++].ToString();
                string Rappelsursalaire = rez[index++].ToString();
                string Rappelprimedetransport = rez[index++].ToString();
                string PrimeperequationRC = rez[index++].ToString();
                string PrimeperequationAV = rez[index++].ToString();
                string PrimeAid = rez[index++].ToString();
                string Primemariage = rez[index++].ToString();
                string Primedenaissance = rez[index++].ToString();
                string Congespayes = rez[index++].ToString();
                string Indemnitedepreavis = rez[index++].ToString();
                string Gratificationfindeservice = rez[index++].ToString();
                string Indemnitedelicenciement = rez[index++].ToString();
                string Avticketsrestaurant = rez[index++].ToString();
                string AvAssurancemaladie = rez[index++].ToString();
                string Avassuranceretraitecomplementaire = rez[index++].ToString();
                string Avcarburant = rez[index++].ToString();
                string Avvoiture = rez[index++].ToString();
                string Avscolariteenfants = rez[index++].ToString();
                string Avutiliteexpat = rez[index++].ToString();
                string Avlogement = rez[index++].ToString();
                string SalaireBrut = rez[index++].ToString();
                string BrutEricsson = rez[index++].ToString();
                string CNSSEmploye = rez[index++].ToString();
                string AssuranceretraitecomplemCNRPSemploye = rez[index++].ToString();
                string AssuranceretraiteComplementaire = rez[index++].ToString();
                string Salaireimposable = rez[index++].ToString();
                string IRPP = rez[index++].ToString();
                string Redevance1 = rez[index++].ToString();
                string PrêtCNSS = rez[index++].ToString();
                string AutresRetenuesursalaire = rez[index++].ToString();
                string DeductionAvTR = rez[index++].ToString();
                string DeductionAvassurancemaladie = rez[index++].ToString();
                string DeductionAvassurancecomplementaire = rez[index++].ToString();
                string DeductionAvVoiture = rez[index++].ToString();
                string DeductionAvCarburant = rez[index++].ToString();
                string DeductionAvloyer = rez[index++].ToString();
                string DeductionAvIndExpat = rez[index++].ToString();
                string Deductionutiliteexpat = rez[index++].ToString();
                string NetàPayer = rez[index++].ToString();
                string CNSSEmployeur = rez[index++].ToString();
                string Accidentdetravail = rez[index++].ToString();
                string AssurancemaladieEmployeur = rez[index++].ToString();
                string AssuranceComplementaireEmployeur = rez[index++].ToString();
                string AssuranceretraitcompleCNRPSemplye = rez[index++].ToString();
                string TFP = rez[index++].ToString();
                string FOPROLOS = rez[index++].ToString();
                string TotalContributionEmployeur = rez[index++].ToString();
                string Moisdepaie = rez[index++].ToString();



                worksheet = EcrireSGNA021(cpt, worksheet, matricule,
             nom,
             prenom,
                 Datedembauchesociete,
                         Datededepartsociete,
                         Intituledepartement,
                         CodeCostcenter,
                         Intitulecostcenter,
                         Intituledelanatureducontrat,
                         Emploioccupe,
                         Salairedebase,
                         Nbjoursdepresence,
                         SalairedebaseJtrvaille,
                         Inddepresence,
                         InddeTransport,
                         PrimeComplementairedePresence,
                         IndemnitedeVoiture,
                         IndFixederepresentation,
                         Indemnitedexpatriation,
                        HS125HS0,
                        ValeurHS0,
                        HS1502,
                        ValeurHS02,
                        HS175HS06,
                        ValeurHS06,
                        HS200HS03,
                        ValeurHS03,
                        HsupplementairenuitHS04,
                        valeurHsuplementairenuitHS04,
                        primeLVP,
                        PrimeSIP,
                        PrimeAOS,
                        Bonus,
                        Compensationsupport,
                        Rappelsursalaire,
                        Rappelprimedetransport,
                        PrimeperequationRC,
                        PrimeperequationAV,
                        PrimeAid,
                        Primemariage,
                        Primedenaissance,
                        Congespayes,
                        Indemnitedepreavis,
                        Gratificationfindeservice,
                        Indemnitedelicenciement,
                        Avticketsrestaurant,
                        AvAssurancemaladie,
                        Avassuranceretraitecomplementaire,
                        Avcarburant,
                        Avvoiture,
                        Avscolariteenfants,
                        Avutiliteexpat,
                        Avlogement,
                        SalaireBrut,
                        BrutEricsson,
                        CNSSEmploye,
                        AssuranceretraitecomplemCNRPSemploye,
                        AssuranceretraiteComplementaire,
                        Salaireimposable,
                        IRPP,
                        Redevance1,
                        PrêtCNSS,
                        AutresRetenuesursalaire,
                        DeductionAvTR,
                        DeductionAvassurancemaladie,
                        DeductionAvassurancecomplementaire,
                        DeductionAvVoiture,
                        DeductionAvCarburant,
                        DeductionAvloyer,
                        DeductionAvIndExpat,
                        Deductionutiliteexpat,
                        NetàPayer,
                        CNSSEmployeur,
                        Accidentdetravail,
                        AssurancemaladieEmployeur,
                        AssuranceComplementaireEmployeur,
                        AssuranceretraitcompleCNRPSemplye,
                        TFP,
                        FOPROLOS,
                        TotalContributionEmployeur,
                        Moisdepaie
);


                cpt = cpt + 1;
            }

            return worksheet;
        }


        public Worksheet EnteteSVRLigne12A020(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA020(refSoc, MoisPe, AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string IdOracle = rez["IdOracle"].ToString();
                string Matricule = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string CodeDept = rez["CodeDept"].ToString();
                string IntituleDept = rez["IntituleDept"].ToString();
                string ExpatLocal = rez["ExpatLocal"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string CentreCout = rez["CentreCout"].ToString();
                string A1000 = rez["A1100"].ToString();
                string A1910 = rez["A1910"].ToString();
                string A1920 = rez["A1920"].ToString();
                string A2311 = rez["A2311"].ToString();
                string A3317 = rez["A3317"].ToString();
                string A3510 = rez["A3510"].ToString();
                string A3711 = rez["A3711"].ToString();
                string HS01 = rez["HS01"].ToString();
                string A4110 = rez["A4110"].ToString();
                string HS02 = rez["HS02"].ToString();
                string A4111 = rez["A4111"].ToString();
                string HS03 = rez["HS03"].ToString();
                string A4113 = rez["A4113"].ToString();
                string A4115 = rez["A4115"].ToString();
                string A4100 = rez["A4100"].ToString();
                string A4159 = rez["A4159"].ToString();
                string A4175 = rez["A4175"].ToString();
                string A4224 = rez["A4224"].ToString();
                string A4160 = rez["A4160"].ToString();
                string A4222 = rez["A4222"].ToString();
                string A4340 = rez["A4340"].ToString();
                string A4360 = rez["A4360"].ToString();
                string A4361 = rez["A4361"].ToString();
                string A4380 = rez["A4380"].ToString();
                string A4381 = rez["A4381"].ToString();
                string A7117 = rez["A7117"].ToString();
                string A4410 = rez["A4410"].ToString();
                string A4816 = rez["A4816"].ToString();
                string A4817 = rez["A4817"].ToString();
                string A4412 = rez["A4412"].ToString();
                string A4450 = rez["A4450"].ToString();
                string A4460 = rez["A4460"].ToString();
                string A4550 = rez["A4550"].ToString();
                string A4711 = rez["A4711"].ToString();
                string A4712 = rez["A4712"].ToString();
                string A4715 = rez["A4715"].ToString();
                string A4716 = rez["A4716"].ToString();
                string A4717 = rez["A4717"].ToString();
                string A4790 = rez["A4790"].ToString();
                string A4811 = rez["A4811"].ToString();
                string A4812 = rez["A4812"].ToString();
                string A4813 = rez["A4813"].ToString();
                string IFRT = rez["IFRT"].ToString();
                string IndemRetAnt = rez["IndemRetAnt"].ToString();
                string LiquidIfrt = rez["LiquidIfrt"].ToString();
                string A5100 = rez["A5100"].ToString();
                string A5420 = rez["A5420"].ToString();
                string A5200 = rez["A5200"].ToString();
                string A5400 = rez["A5400"].ToString();
                string PrimTrim = rez["PrimTrim"].ToString();
                string COPRIMOI = rez["COPRIMOI"].ToString();
                string A4171 = rez["A4171"].ToString();
                string HA04 = rez["HA04"].ToString();
                string A4170 = rez["A4170"].ToString();
                string A7114 = rez["A7114"].ToString();
                string A7191 = rez["A7191"].ToString();
                string A7193 = rez["A7193"].ToString();
                string A7196 = rez["A7196"].ToString();
                string A7692 = rez["A7692"].ToString();
                string A7115 = rez["A7115"].ToString();
                string A7118 = rez["A7118"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string A9112 = rez["A9112"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A6120 = rez["A6120"].ToString();
                string A6201 = rez["A6201"].ToString();
                string A6221 = rez["A6221"].ToString();
                string A9121 = rez["A9121"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A6840 = rez["A6840"].ToString();
                string A6841 = rez["A6841"].ToString();
                string A6835 = rez["A6835"].ToString();
                string A6831 = rez["A6831"].ToString();
                string A44509 = rez["A44509"].ToString();
                string A6832 = rez["A6832"].ToString();
                string A6834 = rez["A6834"].ToString();
                string A6850 = rez["A6850"].ToString();
                string A6410 = rez["A6410"].ToString();
                string A6430 = rez["A6430"].ToString();
                string A6470 = rez["A6470"].ToString();
                string A6491 = rez["A6491"].ToString();
                string A6520 = rez["A6520"].ToString();
                string A6510 = rez["A6510"].ToString();
                string A6521 = rez["A6521"].ToString();
                string A6591 = rez["A6591"].ToString();
                string A6614 = rez["A6614"].ToString();
                string A8201 = rez["A8201"].ToString();
                string A8215 = rez["A8215"].ToString();
                string A8217 = rez["A8217"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A61109 = rez["A61109"].ToString();
                string A61209 = rez["A61209"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A71159 = rez["A71159"].ToString();
                string A62019 = rez["A62019"].ToString();
                string RUB_COTEPY = rez["RUB_COTEPY"].ToString();
                string MoisPaie = rez["MoisPe"].ToString();


                worksheet = EcrireSGNA020(cpt, worksheet, IdOracle,  Matricule,  Nom,  Prenom,  DateEmbauche,  DateDepart,  CodeDept,  IntituleDept,
        ExpatLocal,  Emploi,  CentreCout,  A1000,  A1910,  A1920,  A2311,  A3317,  A3510,  A3711,  HS01,  A4110,
        HS02,  A4111,  HS03,  A4113,  A4115,  A4100,  A4159,  A4175,  A4224,  A4160,  A4222,  A4340,  A4360,
        A4361, A4380, A4381, A7117, A4410, A4816,A4817, A4412, A4450, A4460, A4550, A4711, A4712,
        A4715, A4716, A4717, A4790, A4811, A4812, A4813, IFRT, IndemRetAnt, LiquidIfrt, A5100,
        A5420,A5200,A5400, PrimTrim, COPRIMOI, A4171, HA04, A4170, A7114, A7191, A7193, A7196, A7692, A7115,
        A7118, BRUT, A9112, A6110, A6120, A6201, A6221, A9121, A6310, A6311, A6840, A6841, A6835,
        A6831, A44509, A6832, A6834, A6850, A6410, A6430, A6470, A6491, A6520, A6510, A6521,
        A6591, A6614, A8201, A8215, A8217, RUB_DEDSA, NETPAIE, A61109, A61209, A6150, A71159, A62019,
        RUB_COTEPY, MoisPaie);
                cpt = cpt + 1;
            }

            return worksheet;
        }

        public Worksheet EnteteSVRLigne12A014(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA014(refSoc, MoisPe, AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string IdOracle = rez["IdOracle"].ToString();
                string Matricule = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string LocalExpat = rez["LocalExpat"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string IntituleDept = rez["IntituleDept"].ToString();
                string CentreCout = rez["CentreCout"].ToString();
                string A1000 = rez["A1000"].ToString();
                string A1910 = rez["A1910"].ToString();
                string A1920 = rez["A1920"].ToString();
                string A3310 = rez["A3310"].ToString();
                string A3317 = rez["A3317"].ToString();
                string A3711 = rez["A3711"].ToString();
                string HS01 = rez["HS01"].ToString();
                string A4110 = rez["A4110"].ToString();
                string HS02 = rez["HS02"].ToString();
                string A4111 = rez["A4111"].ToString();
                string HS03 = rez["HS03"].ToString();
                string A4113 = rez["A4113"].ToString();
                string A4115 = rez["A4115"].ToString();
                string A4100 = rez["A4100"].ToString();
                string A4170 = rez["A4170"].ToString();
                string A4159 = rez["A4159"].ToString();
                string A4175 = rez["A4175"].ToString();
                string A4224 = rez["A4224"].ToString();
                string A4411 = rez["A4411"].ToString();
                string A4816 = rez["A4816"].ToString();
                string A4817 = rez["A4817"].ToString();
                string A4222 = rez["A4222"].ToString();
                string A4340 = rez["A4340"].ToString();
                string A4360 = rez["A4360"].ToString();
                string A4361 = rez["A4361"].ToString();
                string A4380 = rez["A4380"].ToString();
                string A4381 = rez["A4381"].ToString();
                string A4410 = rez["A4410"].ToString();
                string A4412 = rez["A4412"].ToString();
                string A4450 = rez["A4450"].ToString();
                string A4460 = rez["A4460"].ToString();
                string A4550 = rez["A4550"].ToString();
                string A4711 = rez["A4711"].ToString();
                string A4712 = rez["A4712"].ToString();
                string A4715 = rez["A4715"].ToString();
                string A4716 = rez["A4716"].ToString();
                string A4717 = rez["A4717"].ToString();
                string A4790 = rez["A4790"].ToString();
                string A4811 = rez["A4811"].ToString();
                string A4610 = rez["A4610"].ToString();
                string A4812 = rez["A4812"].ToString();
                string A4813 = rez["A4813"].ToString();
                string IFRT = rez["IFRT"].ToString();
                string IndemRet = rez["IndemRet"].ToString();
                string Rappel = rez["A5100"].ToString();
                string A5420 = rez["A5420"].ToString();
                string A5200 = rez["A5200"].ToString();
                string A5400 = rez["A5400"].ToString();
                string PrimeTrim = rez["PrimeTrim"].ToString();
                string CL03 = rez["CL03"].ToString();
                string A4171 = rez["A4171"].ToString();
                string A7193 = rez["A7193"].ToString();
                string A7114 = rez["A7114"].ToString();
                string A7117 = rez["A7117"].ToString();
                string A7191 = rez["A7191"].ToString();
                string A7192 = rez["A7192"].ToString();
                string A7197 = rez["A7197"].ToString();
                string A7196 = rez["A7196"].ToString();
                string A7198 = rez["A7198"].ToString();
                string A7199 = rez["A7199"].ToString();
                string A7692 = rez["A7692"].ToString();
                string A7115 = rez["A7115"].ToString();
                string A7118 = rez["A7118"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string A9112 = rez["A9112"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A6120 = rez["A6120"].ToString();
                string A6201 = rez["A6201"].ToString();
                string A6221 = rez["A6221"].ToString();
                string A6222 = rez["A6222"].ToString();
                string A9121 = rez["A9121"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A7115A = rez["A6840"].ToString();
                string A7118A = rez["A6841"].ToString();
                string A7193A = rez["A6835"].ToString();
                string A7114A = rez["A6831"].ToString();
                string A6837 = rez["A6837"].ToString();
                string A6832 = rez["A6832"].ToString();
                string A6834 = rez["A6834"].ToString();
                string A6842 = rez["A6842"].ToString();
                string A6843 = rez["A6843"].ToString();
                string A6844 = rez["A6844"].ToString();
                string A6850 = rez["A6850"].ToString();
                string A6410 = rez["A6410"].ToString();
                string A6430 = rez["A6430"].ToString();
                string A6470 = rez["A6470"].ToString();
                string A6491 = rez["A6491"].ToString();
                string A6614 = rez["A6614"].ToString();
                string A6530 = rez["A6530"].ToString();
                string A6510 = rez["A6510"].ToString();
                string A6521 = rez["A6521"].ToString();
                string A6520 = rez["A6520"].ToString();
                string A8201 = rez["A8201"].ToString();
                string A8215 = rez["A8215"].ToString();
                string A8217 = rez["A8217"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A6110A = rez["A61109"].ToString();
                string A6120A = rez["A61209"].ToString();
                string A6150 = rez["A6150"].ToString();
                string A7115AA = rez["A711599"].ToString();
                string A7118AA = rez["A711899"].ToString();
                string RUB_COTEPY = rez["RUB_COTEPY"].ToString();
                string MoisP = rez["MoisPe"].ToString();

                worksheet = EcrireSGNA014(cpt, worksheet, IdOracle, Matricule, Nom, Prenom, DateEmbauche, DateDepart, LocalExpat, Emploi,
          IntituleDept, CentreCout, A1000, A1910, A1920, A3310, A3317, A3711, HS01, A4110, HS02, A4111,
           HS03, A4113, A4115, A4100, A4170, A4159, A4175, A4224, A4411, A4816,A4817, A4222, A4340, A4360,
           A4361, A4380, A4381, A4410, A4412, A4450, A4460, A4550, A4711, A4712, A4715, A4716, A4717,
           A4790, A4811,A4610, A4812, A4813, IFRT, IndemRet, Rappel, A5420,A5200,A5400, PrimeTrim, CL03, A4171, A7193,
           A7114, A7117, A7191, A7192, A7197, A7196, A7198, A7199, A7692, A7115, A7118, BRUT, A9112,
           A6110, A6120, A6201, A6221,A6222, A9121, A6310, A6311, A7115A, A7118A, A7193A, A7114A, A6837, A6832,
           A6834, A6842, A6843, A6844, A6850, A6410, A6430, A6470, A6491, A6614, A6530, A6510, A6521,
           A6520, A8201, A8215, A8217, RUB_DEDSA, NETPAIE, A6110A, A6120A, A6150, A7115AA, A7118AA, RUB_COTEPY,MoisP);
                cpt = cpt + 1;
            }

            return worksheet;
        }

        public Worksheet EnteteSVRLigne12A015(Worksheet worksheet, int refSoc, int MoisPe, int AnnePe)
        {
            int cpt = 0;
            DataTable dt = service.RecupererDATASGNA015(refSoc, MoisPe, AnnePe);
            foreach (System.Data.DataRow rez in dt.Rows)
            {
                string Matricule = rez["Matricule"].ToString();
                string Nom = rez["Nom"].ToString();
                string Prenom = rez["Prenom"].ToString();
                string DateEmbauche = rez["DateEmbauche"].ToString();
                string DateDepart = rez["DateDepart"].ToString();
                string CodeDept = rez["CodeDept"].ToString();
                string IntituleDept = rez["IntituleDept"].ToString();
                string NatContrat = rez["NatContrat"].ToString();
                string Emploi = rez["Emploi"].ToString();
                string SalBaseAnn = rez["SalBaseAnn"].ToString();
                string SALBASE_M = rez["A1100"].ToString();
                string CL01 = rez["CL01"].ToString();
                string A1200 = rez["A1200"].ToString();
                string A3410 = rez["A3410"].ToString();
                string A3110 = rez["A3110"].ToString();
                string A3711 = rez["A3711"].ToString();
                string A4222 = rez["A4222"].ToString();
                string A4362 = rez["A4362"].ToString();
                string A4311 = rez["A4311"].ToString();
                string A4380 = rez["A4380"].ToString();
                string A4711 = rez["A4711"].ToString();
                string A4717 = rez["A4717"].ToString();
                string A4718 = rez["A4718"].ToString();
                string A4364 = rez["A4364"].ToString();
                string A4171 = rez["A4171"].ToString();
                string A5100 = rez["A5100"].ToString();
                string A7111 = rez["A7111"].ToString();
                string A7112 = rez["A7112"].ToString();
                string A7118 = rez["A7118"].ToString();
                string BRUT = rez["BRUT"].ToString();
                string A6110 = rez["A6110"].ToString();
                string A6120 = rez["A6120"].ToString();
                string A9121 = rez["A9121"].ToString();
                string A6310 = rez["A6310"].ToString();
                string A6311 = rez["A6311"].ToString();
                string A8000 = rez["A8000"].ToString();
                string A6850 = rez["A6850"].ToString();
                string A6220 = rez["A6220"].ToString();
                string RUB_DEDSA = rez["RUB_DEDSA"].ToString();
                string NETPAIE = rez["NETPAIE"].ToString();
                string A6110A = rez["A61109"].ToString();
                string A6120A = rez["A61209"].ToString();
                string A6150 = rez["A6150"].ToString();
                string RUB_EPYC = rez["RUB_COTEPY"].ToString();
                string MoisP = rez["MoisPe"].ToString();


                worksheet = EcrireSGNA015(cpt, worksheet, Matricule, Nom, Prenom, DateEmbauche, DateDepart, CodeDept,
           IntituleDept, NatContrat, Emploi, SalBaseAnn, SALBASE_M, CL01, A1200, A3410,
           A3110, A3711, A4222, A4362, A4311, A4380, A4711, A4717, A4718, A4364,
           A4171,A5100, A7111, A7112, A7118, BRUT, A6110, A6120, A9121, A6310, A6311,A8000,
           A6850, A6220, RUB_DEDSA, NETPAIE, A6110A, A6120A, A6150, RUB_EPYC, MoisP);
                cpt = cpt + 1;
            }

            return worksheet;
        }


        public Worksheet EcrireSGNA002(int cpt,Worksheet worksheet,string Mat,string Nom,string Prenom,string DateEmbauche,string DateDepart,
            string CodeSce,string IntituleSce,string NatureContrat,string Emploi,string A1100,string PRESC1J,string PRESC2H,string A1000,
            string A2200,string A2100,string A4170,string A4173,string A4174,string A3710,string A4171,string A4110,string A4111,string A4113,
            string A4130,string A4140,string A4217,string A4218,string A4220,string A4221,string A4222,string A4340,string A4711,string A4712,
            string A4715,string A4717,string CL25,string CL24,string A7113,string A7114,string A7115,string A4811,string A4812,string A4813,string A4815,
            string BRUT,string A9112,string A6110,string A6210,string A9121,string A6310,string A6311,string A6614,string A6410,string A6850,
            string A6520,string A6430,string A6490,string A6830,string A6831,string A6840,string A8800,string RUB_DEDSA,string NETPAIE,
            string A6110A,string A6150,string A7115A,string A7116,string RUBCOTEPY,string MOIS_PAIE)
        {
            worksheet.Cells[cpt + 11, 0] = new Cell(Mat);
            worksheet.Cells[cpt + 11, 1] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 2] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 11, 5] = new Cell(CodeSce);
            worksheet.Cells[cpt + 11, 6] = new Cell(IntituleSce);
            worksheet.Cells[cpt + 11, 7] = new Cell(NatureContrat);
            worksheet.Cells[cpt + 11, 8] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 9] = new Cell(A1100);
            worksheet.Cells[cpt + 11, 10] = new Cell(PRESC1J);
            worksheet.Cells[cpt + 11, 11] = new Cell(PRESC2H);
            worksheet.Cells[cpt + 11, 12] = new Cell(A1000);
            worksheet.Cells[cpt + 11, 13] = new Cell(A2200);
            worksheet.Cells[cpt + 11, 14] = new Cell(A2100);
            worksheet.Cells[cpt + 11, 15] = new Cell(A4170);
            worksheet.Cells[cpt + 11, 16] = new Cell(A4173);
            worksheet.Cells[cpt + 11, 17] = new Cell(A4174);
            worksheet.Cells[cpt + 11, 18] = new Cell(A3710);
            worksheet.Cells[cpt + 11, 19] = new Cell(A4171);
            worksheet.Cells[cpt + 11, 20] = new Cell(A4110);
            worksheet.Cells[cpt + 11, 21] = new Cell(A4111);
            worksheet.Cells[cpt + 11, 22] = new Cell(A4113);
            worksheet.Cells[cpt + 11, 23] = new Cell(A4130);
            worksheet.Cells[cpt + 11, 24] = new Cell(A4140);
            worksheet.Cells[cpt + 11, 25] = new Cell(A4217);
            worksheet.Cells[cpt + 11, 26] = new Cell(A4218);
            worksheet.Cells[cpt + 11, 27] = new Cell(A4220);
            worksheet.Cells[cpt + 11, 28] = new Cell(A4221);
            worksheet.Cells[cpt + 11, 29] = new Cell(A4222);
            worksheet.Cells[cpt + 11, 30] = new Cell(A4340);
            worksheet.Cells[cpt + 11, 31] = new Cell(A4711);
            worksheet.Cells[cpt + 11, 32] = new Cell(A4712);
            worksheet.Cells[cpt + 11, 33] = new Cell(A4715);
            worksheet.Cells[cpt + 11, 34] = new Cell(A4717);
            worksheet.Cells[cpt + 11, 35] = new Cell(CL25);
            worksheet.Cells[cpt + 11, 36] = new Cell(CL24);
            worksheet.Cells[cpt + 11, 37] = new Cell(A7113);
            worksheet.Cells[cpt + 11, 38] = new Cell(A7114);
            worksheet.Cells[cpt + 11, 39] = new Cell(A7115);
            worksheet.Cells[cpt + 11, 40] = new Cell(A7116);
            worksheet.Cells[cpt + 11, 41] = new Cell(A4811);
            worksheet.Cells[cpt + 11, 42] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 43] = new Cell(A4813);
            worksheet.Cells[cpt + 11, 44] = new Cell(A4815);
            worksheet.Cells[cpt + 11, 45] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 46] = new Cell(A9112);
            worksheet.Cells[cpt + 11, 47] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 48] = new Cell(A6210);
            worksheet.Cells[cpt + 11, 49] = new Cell(A9121);
            worksheet.Cells[cpt + 11, 50] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 51] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 52] = new Cell(A6614);
            worksheet.Cells[cpt + 11, 53] = new Cell(A6410);
            worksheet.Cells[cpt + 11, 54] = new Cell(A6520);
            worksheet.Cells[cpt + 11, 55] = new Cell(A6430);
            worksheet.Cells[cpt + 11, 56] = new Cell(A6490);
            worksheet.Cells[cpt + 11, 57] = new Cell(A6830);
            worksheet.Cells[cpt + 11, 58] = new Cell(A6831);
            worksheet.Cells[cpt + 11, 59] = new Cell(A6840);
            worksheet.Cells[cpt + 11, 60] = new Cell(A6850);
            worksheet.Cells[cpt + 11, 61] = new Cell(A8800);
            worksheet.Cells[cpt + 11, 62] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 63] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 64] = new Cell(A6110A);
            worksheet.Cells[cpt + 11, 65] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 66] = new Cell(A7115A);
            worksheet.Cells[cpt + 11, 67] = new Cell(RUBCOTEPY);
            worksheet.Cells[cpt + 11, 68] = new Cell(MOIS_PAIE);

            return worksheet;
        }

        public Worksheet EcrireSGNA003(int cpt, Worksheet worksheet, string Mat, string Nom, string Prenom, string DateEmbauche, string DateDepart,
           string CodeSce, string IntituleSce, string NatureContrat, string Emploi,string SALBASE, string CL01,string A1200,string A2200,
            string A2100,string A2360,string A2330,string A3412,string A3317,string A3712,string A3710,string A3711,string A3713,string A4112,
            string A4113,string A4120,string A4222,string A4216,string A4361,string A4215,string A4320,string A4370,string A4430,string A4812,string A5100,string A7113,
            string A4171,string BRUT,string AjusCot,string A9112,string A6110,string A6120,string A9310,string A6310,string A6311,
            string A6340,string A6210,string A6410,string A6430,string A6520,string A6510,string A6530,string A6830,string A6490,string RUB_DEDSA,
            string NETPAIE,string A6110A,string A6120A,string A6150,string A6210A,string RUB_COTEPY, string MoisP,string A4380)
        {
            worksheet.Cells[cpt + 11, 0] = new Cell(Mat);
            worksheet.Cells[cpt + 11, 1] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 2] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 11, 5] = new Cell(CodeSce);
            worksheet.Cells[cpt + 11, 6] = new Cell(IntituleSce);
            worksheet.Cells[cpt + 11, 7] = new Cell(NatureContrat);
            worksheet.Cells[cpt + 11, 8] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 9] = new Cell(SALBASE);
            worksheet.Cells[cpt + 11, 10] = new Cell(CL01);
            worksheet.Cells[cpt + 11, 11] = new Cell(A1200);
            worksheet.Cells[cpt + 11, 12] = new Cell(A2200);
            worksheet.Cells[cpt + 11, 13] = new Cell(A2100);
            worksheet.Cells[cpt + 11, 14] = new Cell(A2360);
            worksheet.Cells[cpt + 11, 15] = new Cell(A2330);
            worksheet.Cells[cpt + 11, 16] = new Cell(A3412);
            worksheet.Cells[cpt + 11, 17] = new Cell(A3317);
            worksheet.Cells[cpt + 11, 18] = new Cell(A3712);
            worksheet.Cells[cpt + 11, 19] = new Cell(A3710);
            worksheet.Cells[cpt + 11, 20] = new Cell(A3711);
            worksheet.Cells[cpt + 11, 21] = new Cell(A4380);

            worksheet.Cells[cpt + 11, 22] = new Cell(A3713);
            worksheet.Cells[cpt + 11, 23] = new Cell(A4112);
            worksheet.Cells[cpt + 11, 24] = new Cell(A4113);
            worksheet.Cells[cpt + 11, 25] = new Cell(A4120);
            worksheet.Cells[cpt + 11, 26] = new Cell(A4222);
            worksheet.Cells[cpt + 11, 27] = new Cell(A4216);
            worksheet.Cells[cpt + 11, 28] = new Cell(A4361);
            worksheet.Cells[cpt + 11, 29] = new Cell(A4215);
            worksheet.Cells[cpt + 11,30] = new Cell(A4320);
            worksheet.Cells[cpt + 11, 31] = new Cell(A4370);
            worksheet.Cells[cpt + 11, 32] = new Cell(A4430);
            worksheet.Cells[cpt + 11, 33] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 34] = new Cell(A5100);
            worksheet.Cells[cpt + 11, 35] = new Cell(A7113);
            worksheet.Cells[cpt + 11, 36] = new Cell(A4171);
            worksheet.Cells[cpt + 11, 37] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 38] = new Cell(AjusCot);
            worksheet.Cells[cpt + 11, 39] = new Cell(A9112);
            worksheet.Cells[cpt + 11, 40] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 41] = new Cell(A6120);
            worksheet.Cells[cpt + 11, 42] = new Cell(A9310);
            worksheet.Cells[cpt + 11, 43] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 44] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 45] = new Cell(A6340);
            worksheet.Cells[cpt + 11, 46] = new Cell(A6210);
            worksheet.Cells[cpt + 11, 47] = new Cell(A6410);
            worksheet.Cells[cpt + 11, 48] = new Cell(A6430);
            worksheet.Cells[cpt + 11, 49] = new Cell(A6520);
            worksheet.Cells[cpt + 11, 50] = new Cell(A6510);
            worksheet.Cells[cpt + 11, 51] = new Cell(A6530);
            worksheet.Cells[cpt + 11, 52] = new Cell(A6830);
            worksheet.Cells[cpt + 11, 53] = new Cell(A6490);
            worksheet.Cells[cpt + 11, 54] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 55] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 56] = new Cell(A6110A);
            worksheet.Cells[cpt + 11, 57] = new Cell(A6120A);
            worksheet.Cells[cpt + 11, 58] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 59] = new Cell(A6210A);
            worksheet.Cells[cpt + 11, 60] = new Cell(RUB_COTEPY);
            worksheet.Cells[cpt + 11, 61] = new Cell(MoisP);

            return worksheet;
        }

        public Worksheet EcrireSGNA004(int cpt, Worksheet worksheet,string Mat,string Nom,string Prenom, string DateEmbauche,string DateDepart,
            string CodeSce,string IntituleSce,string NatureContrat,string Emploi,string SALBASE,string CL01,string A1100,string A2200,string A2100,
            string A3411,string A3410,string A3412,string A3317,string A3712,string A3710,string A3711,string A3713,string A4112,string A4113,
            string A4120,string A4216,string A4222,string A4361,string A4215,string A4320,string A4370,string A4430,string A4171,string A4380,
            string A4811, string A7113, string BRUT, string AjusCot, string A9112, string A6110, string A6210, string A9310, string A6310,
            string A6311, string A6340, string A6210A, string A6410, string A6430, string A6520, string A6510, string A6530, string A6830, string A6490,
            string RUB_DEDSA,string NETPAIE,string A6110A,string A6120,string A6150,
            string A6210AA, string RUB_COTEPY, string MoisP,string A4812)
        {
            worksheet.Cells[cpt + 11, 0] = new Cell(Mat);
            worksheet.Cells[cpt + 11, 1] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 2] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(DateEmbauche);
            }

            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateDepart);
            }

            worksheet.Cells[cpt + 11, 5] = new Cell(CodeSce);
            worksheet.Cells[cpt + 11, 6] = new Cell(IntituleSce);
            worksheet.Cells[cpt + 11, 7] = new Cell(NatureContrat);
            worksheet.Cells[cpt + 11, 8] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 9] = new Cell(SALBASE);
            worksheet.Cells[cpt + 11, 10] = new Cell(CL01);
            worksheet.Cells[cpt + 11, 11] = new Cell(A1100);
            worksheet.Cells[cpt + 11, 12] = new Cell(A2200);
            worksheet.Cells[cpt + 11, 13] = new Cell(A2100);
            worksheet.Cells[cpt + 11, 14] = new Cell(A3411);
            worksheet.Cells[cpt + 11, 15] = new Cell(A3410);
            worksheet.Cells[cpt + 11, 16] = new Cell(A3412);
            worksheet.Cells[cpt + 11, 17] = new Cell(A3317);
            worksheet.Cells[cpt + 11, 18] = new Cell(A3712);
            worksheet.Cells[cpt + 11, 19] = new Cell(A3710);
            worksheet.Cells[cpt + 11, 20] = new Cell(A3711);
            worksheet.Cells[cpt + 11, 21] = new Cell(A3713);
            worksheet.Cells[cpt + 11, 22] = new Cell(A4112);
            worksheet.Cells[cpt + 11, 23] = new Cell(A4113);
            worksheet.Cells[cpt + 11, 24] = new Cell(A4120);
            worksheet.Cells[cpt + 11, 25] = new Cell(A4216);
            worksheet.Cells[cpt + 11, 26] = new Cell(A4222);
            worksheet.Cells[cpt + 11, 27] = new Cell(A4361);
            worksheet.Cells[cpt + 11, 28] = new Cell(A4215);
            worksheet.Cells[cpt + 11, 29] = new Cell(A4320);
            worksheet.Cells[cpt + 11, 30] = new Cell(A4370);
            worksheet.Cells[cpt + 11, 31] = new Cell(A4430);
            worksheet.Cells[cpt + 11, 32] = new Cell(A4171);
            worksheet.Cells[cpt + 11, 33] = new Cell(A4380);
            worksheet.Cells[cpt + 11, 34] = new Cell(A4811);
            worksheet.Cells[cpt + 11, 35] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 36] = new Cell(A7113);
            worksheet.Cells[cpt + 11, 37] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 38] = new Cell(AjusCot);
            worksheet.Cells[cpt + 11, 39] = new Cell(A9112);
            worksheet.Cells[cpt + 11, 40] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 41] = new Cell(A6210);
            worksheet.Cells[cpt + 11, 42] = new Cell(A9310);
            worksheet.Cells[cpt + 11, 43] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 44] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 45] = new Cell(A6340);
            worksheet.Cells[cpt + 11, 46] = new Cell(A6210A);
            worksheet.Cells[cpt + 11, 47] = new Cell(A6410);
            worksheet.Cells[cpt + 11, 48] = new Cell(A6430);
            worksheet.Cells[cpt + 11, 49] = new Cell(A6520);
            worksheet.Cells[cpt + 11, 50] = new Cell(A6510);
            worksheet.Cells[cpt + 11, 51] = new Cell(A6530);
            worksheet.Cells[cpt + 11, 52] = new Cell(A6830);
            worksheet.Cells[cpt + 11, 53] = new Cell(A6490);
            worksheet.Cells[cpt + 11, 54] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 55] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 56] = new Cell(A6110A);
            worksheet.Cells[cpt + 11, 57] = new Cell(A6120);
            worksheet.Cells[cpt + 11, 58] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 59] = new Cell(A6210AA);
            worksheet.Cells[cpt + 11, 60] = new Cell(RUB_COTEPY);
            worksheet.Cells[cpt + 11, 61] = new Cell(MoisP);

            return worksheet;
        }


   



        public Worksheet EcrireSGNA010(int cpt, Worksheet worksheet,string IdOracle,string Matricule,string Nom,string Prenom,string DateEmbauche,
            string DateDepart,string Emploi,string NatureContrat,string CodeDept,string Dept,string HorBaseSal,string BullMod,string Categorie,
            string A1100,string A1300,string A2100,string A2200,string A3310,string ABS_CONGNR,string AbsNonRem,string RUB_ARR,string A4373,string HS01,
                  string A4110,string HS02,string A4111,string HS03,string A4113,string HS06,string A4112,string HS04,string A4120,string HS07,string A4372,
            string HS05,string A4114,string HA04,string HA06,string RUB_JCHOM,string HA07,string RUB_TRP,string HA09,string A5500,
                   string A4222,string A4371,string A4130,string A4135,string A2361,string A4816,string A4810,string HA01,string A4811,string A4812,
            string A5100,string A5400,string A5900,string A5200,string A6510,string A3309,string A4817,string A4818,string A4313,string A4312,
                   string TESTSTC,string A4800,string PRES8,string PRES2,string A7115,string BRUT,string A6110,string A9310,string A6310,string A6311,string A6490,
            string A6491,string A6492,string A6840,string A6610,string A6611,string A6613,string A6612,string A6410,string A6511,string A6614,
            string RUB_DEDSA,string NETPAIE,string A6110A,string A6150,string A7115A,string RUB_COTEPY,string ModPe,string MoisP,string A4374)
        {
            worksheet.Cells[cpt + 11, 0] = new Cell(IdOracle);
            worksheet.Cells[cpt + 11, 1] = new Cell(Matricule);
            worksheet.Cells[cpt + 11, 2] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 3] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 5] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 5] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 11, 6] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 7] = new Cell(NatureContrat);
            worksheet.Cells[cpt + 11, 8] = new Cell(CodeDept);
            worksheet.Cells[cpt + 11, 9] = new Cell(Dept);
            worksheet.Cells[cpt + 11, 10] = new Cell(HorBaseSal);
            worksheet.Cells[cpt + 11, 11] = new Cell(BullMod);
            worksheet.Cells[cpt + 11, 12] = new Cell(Categorie);
            worksheet.Cells[cpt + 11, 13] = new Cell(A1100);
            worksheet.Cells[cpt + 11, 14] = new Cell(A1300);
            worksheet.Cells[cpt + 11, 15] = new Cell(A2100);
            worksheet.Cells[cpt + 11, 16] = new Cell(A2200);
            worksheet.Cells[cpt + 11, 17] = new Cell(A3310);
            worksheet.Cells[cpt + 11, 18] = new Cell(ABS_CONGNR);
            worksheet.Cells[cpt + 11, 19] = new Cell(AbsNonRem);
            worksheet.Cells[cpt + 11, 20] = new Cell(RUB_ARR);
            worksheet.Cells[cpt + 11, 21] = new Cell(A4373);
            worksheet.Cells[cpt + 11, 22] = new Cell(HS01);
            worksheet.Cells[cpt + 11, 23] = new Cell(A4110);
            worksheet.Cells[cpt + 11, 24] = new Cell(HS02);
            worksheet.Cells[cpt + 11, 25] = new Cell(A4111);
            worksheet.Cells[cpt + 11, 26] = new Cell(HS03);
            worksheet.Cells[cpt + 11, 27] = new Cell(A4113);
            worksheet.Cells[cpt + 11, 28] = new Cell(HS06);
            worksheet.Cells[cpt + 11, 29] = new Cell(A4112);
            worksheet.Cells[cpt + 11, 30] = new Cell(HS04);
            worksheet.Cells[cpt + 11, 31] = new Cell(A4120);
            worksheet.Cells[cpt + 11, 32] = new Cell(HS07);
            worksheet.Cells[cpt + 11, 33] = new Cell(A4372);
            worksheet.Cells[cpt + 11, 34] = new Cell(HS05);
            worksheet.Cells[cpt + 11, 35] = new Cell(A4374);
            worksheet.Cells[cpt + 11, 36] = new Cell(A4114);
            worksheet.Cells[cpt + 11, 37] = new Cell(HA04);
            worksheet.Cells[cpt + 11, 38] = new Cell(HA06);
            worksheet.Cells[cpt + 11, 39] = new Cell(RUB_JCHOM);
            worksheet.Cells[cpt + 11, 40] = new Cell(HA07);
            worksheet.Cells[cpt + 11, 41] = new Cell(RUB_TRP);
            worksheet.Cells[cpt + 11, 42] = new Cell(HA09);
            worksheet.Cells[cpt + 11, 43] = new Cell(A5500);
            worksheet.Cells[cpt + 11, 44] = new Cell(A4222);
            worksheet.Cells[cpt + 11, 45] = new Cell(A4371);
            worksheet.Cells[cpt + 11, 46] = new Cell(A4130);
            worksheet.Cells[cpt + 11, 47] = new Cell(A4135);
            worksheet.Cells[cpt + 11, 48] = new Cell(A2361);
            worksheet.Cells[cpt + 11, 49] = new Cell(A4816);
            worksheet.Cells[cpt + 11, 50] = new Cell(A4810);
            worksheet.Cells[cpt + 11, 51] = new Cell(HA01);
            worksheet.Cells[cpt + 11, 52] = new Cell(A4811);
            worksheet.Cells[cpt + 11, 53] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 54] = new Cell(A5100);
            worksheet.Cells[cpt + 11, 55] = new Cell(A5400);
            worksheet.Cells[cpt + 11, 56] = new Cell(A5900);
            worksheet.Cells[cpt + 11, 57] = new Cell(A5200);
            worksheet.Cells[cpt + 11, 58] = new Cell(A6510);
            worksheet.Cells[cpt + 11, 59] = new Cell(A3309);
            worksheet.Cells[cpt + 11, 60] = new Cell(A4817);
            worksheet.Cells[cpt + 11, 61] = new Cell(A4818);
            worksheet.Cells[cpt + 11, 62] = new Cell(A4313);
            worksheet.Cells[cpt + 11, 63] = new Cell(A4312);
            worksheet.Cells[cpt + 11, 64] = new Cell(TESTSTC);
            worksheet.Cells[cpt + 11, 65] = new Cell(A4800);
            worksheet.Cells[cpt + 11, 66] = new Cell(PRES8);
            worksheet.Cells[cpt + 11, 67] = new Cell(PRES2);
            worksheet.Cells[cpt + 11, 68] = new Cell(A7115);
            worksheet.Cells[cpt + 11, 69] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 70] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 71] = new Cell(A9310);
            worksheet.Cells[cpt + 11, 72] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 73] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 74] = new Cell(A6490);
            worksheet.Cells[cpt + 11, 75] = new Cell(A6491);
            worksheet.Cells[cpt + 11, 76] = new Cell(A6492);
            worksheet.Cells[cpt + 11, 77] = new Cell(A6840);
            worksheet.Cells[cpt + 11, 78] = new Cell(A6610);
            worksheet.Cells[cpt + 11, 79] = new Cell(A6611);
            worksheet.Cells[cpt + 11, 80] = new Cell(A6613);
            worksheet.Cells[cpt + 11, 81] = new Cell(A6612);
            worksheet.Cells[cpt + 11, 82] = new Cell(A6410);
            worksheet.Cells[cpt + 11, 83] = new Cell(A6511);
            worksheet.Cells[cpt + 11, 84] = new Cell(A6614);
            worksheet.Cells[cpt + 11, 85] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 86] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 87] = new Cell(A6110A);
            worksheet.Cells[cpt + 11, 88] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 89] = new Cell(A7115A);
            worksheet.Cells[cpt + 11, 90] = new Cell(RUB_COTEPY);
            worksheet.Cells[cpt + 11, 91] = new Cell(ModPe);
            worksheet.Cells[cpt + 11, 92] = new Cell(MoisP);

            return worksheet;
        }

        public Worksheet EcrireSGNA011(int cpt, Worksheet worksheet, string Matricule,string Nom,string Prenom,string DateEmbauche,string DateDepart,string IntituleDept,
            string IntituleSce,string IntituleCat,string NatContrat,string Emploi,string SalBaseSal,string CL01,string A1200,string A2100,string A2200,string A2311,
            string A3111,string A3413,string A3410,string A4351,string A4220,string A4211,string A4369,string A4360,string A4380,string A4717,string A4716,string A4110,string A4111,
            string A4113, string A4114, string A4812, string A4171, string BRUT, string A6110, string A9121, string A6310, string A6311, string TotDecEmp, string NETPAIE, string A6110A, string A6150,
                string A6320, string A6330, string RUB_EPYC, string TotCostLab, string MoisP)
        {

            worksheet.Cells[cpt + 11, 0] = new Cell(Matricule);
            worksheet.Cells[cpt + 11, 1] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 2] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 11, 5] = new Cell(IntituleDept);
            worksheet.Cells[cpt + 11, 6] = new Cell(IntituleSce);
            worksheet.Cells[cpt + 11, 7] = new Cell(IntituleCat);
            worksheet.Cells[cpt + 11, 8] = new Cell(NatContrat);
            worksheet.Cells[cpt + 11, 9] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 10] = new Cell(SalBaseSal);
            worksheet.Cells[cpt + 11, 11] = new Cell(CL01);
            worksheet.Cells[cpt + 11, 12] = new Cell(A1200);
            worksheet.Cells[cpt + 11, 13] = new Cell(A2100);
            worksheet.Cells[cpt + 11, 14] = new Cell(A2200);
            worksheet.Cells[cpt + 11, 15] = new Cell(A2311);
            worksheet.Cells[cpt + 11, 16] = new Cell(A3111);
            worksheet.Cells[cpt + 11, 17] = new Cell(A3413);
            worksheet.Cells[cpt + 11, 18] = new Cell(A3410);
            worksheet.Cells[cpt + 11, 19] = new Cell(A4351);
            worksheet.Cells[cpt + 11, 20] = new Cell(A4220);
            worksheet.Cells[cpt + 11, 21] = new Cell(A4211);
            worksheet.Cells[cpt + 11, 22] = new Cell(A4369);
            worksheet.Cells[cpt + 11, 23] = new Cell(A4360);
            worksheet.Cells[cpt + 11, 24] = new Cell(A4380);
            worksheet.Cells[cpt + 11, 25] = new Cell(A4717);
            worksheet.Cells[cpt + 11, 26] = new Cell(A4716);
            worksheet.Cells[cpt + 11, 27] = new Cell(A4110);
            worksheet.Cells[cpt + 11, 28] = new Cell(A4111);
            worksheet.Cells[cpt + 11, 29] = new Cell(A4113);
            worksheet.Cells[cpt + 11, 30] = new Cell(A4114);
            worksheet.Cells[cpt + 11, 31] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 32] = new Cell(A4171);
            worksheet.Cells[cpt + 11, 33] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 34] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 35] = new Cell(A9121);
            worksheet.Cells[cpt + 11, 36] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 37] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 38] = new Cell(TotDecEmp);
            worksheet.Cells[cpt + 11, 39] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 40] = new Cell(A6110A);
            worksheet.Cells[cpt + 11, 41] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 42] = new Cell(A6320);
            worksheet.Cells[cpt + 11, 43] = new Cell(A6330);
            worksheet.Cells[cpt + 11, 44] = new Cell(RUB_EPYC);
            worksheet.Cells[cpt + 11, 45] = new Cell(TotCostLab);
            worksheet.Cells[cpt + 11, 46] = new Cell(MoisP);

            return worksheet;
        }


        public Worksheet EcrireSGNA013(int cpt, Worksheet worksheet, string IdOracle, string Matricule, string Nom, string Prenom, string DateEmbauche,
                       string DateDepart, string Emploi, string NatureContrat, string CodeDept, string Dept, string HorBaseSal, string BullMod,
                       string Categorie, string ModPe, string A1100, string A1300, string A1310, string A2100, string A2200, string A3310, string A3110,
                       string ABS_CONGNR, string A4174, string HS01, string A4110, string HS02, string A4111, string HS03, string A4113, string HS04, string A4120,
                       string HS07, string A4372,string A4374, string HS05, string A4114, string HA04, string HA06, string A4175, string HA07, string A4176, string A4130, string A4135,
                       string A2361, string A4222, string A3309, string A4373, string A4371, string A4817, string A4818, string A4313,
                       string A4311, string HA09, string A5500, string A5100, string A5400, string A5200, string A5900, string HA01, string A4816, string A4810, string A4811,
                       string A4812, string TESTSTC, string A4800, string PRES8, string A7113, string CL23, string A7116, string A7115,
                       string BRUT, string A6110, string A9310, string A6310, string A6311, string A6614, string A6410, string A6490, string A6491, string A6492, string A6511,
                       string A6430, string A6431, string A6840, string RUB_DEDSA, string NETPAIE, string A61109, string A6150, string A71159,
                       string RUB_COTEPY, string MoisPe)
        {
            worksheet.Cells[cpt + 11, 0] = new Cell(IdOracle);
            worksheet.Cells[cpt + 11, 1] = new Cell(Matricule);
            worksheet.Cells[cpt + 11, 2] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 3] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 5] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 5] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 11, 6] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 7] = new Cell(NatureContrat);
            worksheet.Cells[cpt + 11, 8] = new Cell(CodeDept);
            worksheet.Cells[cpt + 11, 9] = new Cell(Dept);
            worksheet.Cells[cpt + 11, 10] = new Cell(HorBaseSal);
            worksheet.Cells[cpt + 11, 11] = new Cell(BullMod);
            worksheet.Cells[cpt + 11, 12] = new Cell(Categorie);
            worksheet.Cells[cpt + 11, 13] = new Cell(ModPe);
            worksheet.Cells[cpt + 11, 14] = new Cell(A1100);
            worksheet.Cells[cpt + 11, 15] = new Cell(A1300);
            worksheet.Cells[cpt + 11, 16] = new Cell(A1310);
            worksheet.Cells[cpt + 11, 17] = new Cell(A2100);
            worksheet.Cells[cpt + 11, 18] = new Cell(A2200);
            worksheet.Cells[cpt + 11, 19] = new Cell(A3310);
            worksheet.Cells[cpt + 11, 20] = new Cell(A3110);
            worksheet.Cells[cpt + 11, 21] = new Cell(ABS_CONGNR);
            worksheet.Cells[cpt + 11, 22] = new Cell(A4174);
            worksheet.Cells[cpt + 11, 23] = new Cell(HS01);
            worksheet.Cells[cpt + 11, 24] = new Cell(A4110);
            worksheet.Cells[cpt + 11, 25] = new Cell(HS02);
            worksheet.Cells[cpt + 11, 26] = new Cell(A4111);
            worksheet.Cells[cpt + 11, 27] = new Cell(HS03);
            worksheet.Cells[cpt + 11, 28] = new Cell(A4113);
            worksheet.Cells[cpt + 11, 29] = new Cell(HS04);
            worksheet.Cells[cpt + 11, 30] = new Cell(A4120);
            worksheet.Cells[cpt + 11, 31] = new Cell(HS07);
            worksheet.Cells[cpt + 11, 32] = new Cell(A4372);
            worksheet.Cells[cpt + 11, 33] = new Cell(A4374);
            worksheet.Cells[cpt + 11, 34] = new Cell(HS05);
            worksheet.Cells[cpt + 11, 35] = new Cell(A4114);
            worksheet.Cells[cpt + 11, 36] = new Cell(HA04);
            worksheet.Cells[cpt + 11, 37] = new Cell(HA06);
            worksheet.Cells[cpt + 11, 38] = new Cell(A4175);
            worksheet.Cells[cpt + 11, 39] = new Cell(HA07);
            worksheet.Cells[cpt + 11, 40] = new Cell(A4176);
            worksheet.Cells[cpt + 11, 41] = new Cell(A4130);
            worksheet.Cells[cpt + 11, 42] = new Cell(A4135);
            worksheet.Cells[cpt + 11, 43] = new Cell(A2361);
            worksheet.Cells[cpt + 11, 44] = new Cell(A4222);
            worksheet.Cells[cpt + 11, 45] = new Cell(A3309);
            worksheet.Cells[cpt + 11, 46] = new Cell(A4373);
            worksheet.Cells[cpt + 11, 47] = new Cell(A4371);
            worksheet.Cells[cpt + 11, 48] = new Cell(A4817);
            worksheet.Cells[cpt + 11, 49] = new Cell(A4818);
            worksheet.Cells[cpt + 11, 50] = new Cell(A4313);
            worksheet.Cells[cpt + 11, 51] = new Cell(A4311);
            worksheet.Cells[cpt + 11, 52] = new Cell(HA09);
            worksheet.Cells[cpt + 11, 53] = new Cell(A5500);
            worksheet.Cells[cpt + 11, 54] = new Cell(A5100);
            worksheet.Cells[cpt + 11, 55] = new Cell(A5400);
            worksheet.Cells[cpt + 11, 56] = new Cell(A5200);
            worksheet.Cells[cpt + 11, 57] = new Cell(A5900);
            worksheet.Cells[cpt + 11, 58] = new Cell(HA01);
            worksheet.Cells[cpt + 11, 59] = new Cell(A4816);
            worksheet.Cells[cpt + 11, 60] = new Cell(A4810);
            worksheet.Cells[cpt + 11, 61] = new Cell(A4811);
            worksheet.Cells[cpt + 11, 62] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 63] = new Cell(TESTSTC);
            worksheet.Cells[cpt + 11, 64] = new Cell(A4800);
            worksheet.Cells[cpt + 11, 65] = new Cell(PRES8);
            worksheet.Cells[cpt + 11, 66] = new Cell(A7113);
            worksheet.Cells[cpt + 11, 67] = new Cell(CL23);
            worksheet.Cells[cpt + 11, 68] = new Cell(A7116);
            worksheet.Cells[cpt + 11, 69] = new Cell(A7115);
            worksheet.Cells[cpt + 11, 70] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 71] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 72] = new Cell(A9310);
            worksheet.Cells[cpt + 11, 73] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 74] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 75] = new Cell(A6614);
            worksheet.Cells[cpt + 11, 76] = new Cell(A6410);
            worksheet.Cells[cpt + 11, 77] = new Cell(A6490);
            worksheet.Cells[cpt + 11, 78] = new Cell(A6491);
            worksheet.Cells[cpt + 11, 79] = new Cell(A6492);
            worksheet.Cells[cpt + 11, 80] = new Cell(A6511);
            worksheet.Cells[cpt + 11, 81] = new Cell(A6430);
            worksheet.Cells[cpt + 11, 82] = new Cell(A6431);
            worksheet.Cells[cpt + 11, 83] = new Cell(A6840);
            worksheet.Cells[cpt + 11, 84] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 85] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 86] = new Cell(A61109);
            worksheet.Cells[cpt + 11, 87] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 88] = new Cell(A71159);
            worksheet.Cells[cpt + 11, 89] = new Cell(RUB_COTEPY);
            worksheet.Cells[cpt + 11, 90] = new Cell(MoisPe);

            return worksheet;
        }


        public Worksheet EcrireSGNA020(int cpt, Worksheet worksheet, string IdOracle, string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeDept, string IntituleDept,
       string ExpatLocal, string Emploi, string CentreCout, string A1000, string A1910, string A1920, string A2311, string A3317, string A3510, string A3711, string HS01, string A4110,
       string HS02, string A4111, string HS03, string A4113, string A4115, string A4100, string A4159, string A4175, string A4224, string A4160, string A4222, string A4340, string A4360,
           string A4361, string A4380, string A4381, string A7117, string A4410, string A4816,string A4817, string A4412, string A4450, string A4460, string A4550, string A4711, string A4712,
           string A4715, string A4716, string A4717, string A4790, string A4811, string A4812, string A4813, string IFRT, string IndemRetAnt, string LiquidIfrt, string A5100,
           string A5420,string A5200,string A5400, string PrimTrim, string CongePris, string A4171, string HA04, string A4170, string A7114, string A7191, string A7193, string A7196, string A7692, string A7115,
           string A7118, string BRUT, string A9112, string A6110, string A6120, string A6201, string A6221, string A9121, string A6310, string A6311, string A6840, string A6841, string A6835,
           string A6831, string A44509, string A6832, string A6834, string A6850, string A6410, string A6430, string A6470, string A6491, string A6520, string A6510, string A6521,
           string A6591, string A6614, string A8201, string A8215, string A8217, string RUB_DEDSA, string NETPAIE, string A61109, string A61209, string A6150, string A71159, string A62019,
           string RUB_COTEPY, string MoisPe)
        {
            worksheet.Cells[cpt + 11, 0] = new Cell(IdOracle);
            worksheet.Cells[cpt + 11, 1] = new Cell(Matricule);
            worksheet.Cells[cpt + 11, 2] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 3] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 5] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 5] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 11, 6] = new Cell(CodeDept);
            worksheet.Cells[cpt + 11, 7] = new Cell(IntituleDept);
            worksheet.Cells[cpt + 11, 8] = new Cell(ExpatLocal);
            worksheet.Cells[cpt + 11, 9] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 10] = new Cell(CentreCout);
            worksheet.Cells[cpt + 11, 11] = new Cell(A1000);
            worksheet.Cells[cpt + 11, 12] = new Cell(A1910);
            worksheet.Cells[cpt + 11, 13] = new Cell(A1920);
            worksheet.Cells[cpt + 11, 14] = new Cell(A2311);
            worksheet.Cells[cpt + 11, 15] = new Cell(A3317);
            worksheet.Cells[cpt + 11, 16] = new Cell(A3510);
            worksheet.Cells[cpt + 11, 17] = new Cell(A3711);
            worksheet.Cells[cpt + 11, 18] = new Cell(HS01);
            worksheet.Cells[cpt + 11, 19] = new Cell(A4110);
            worksheet.Cells[cpt + 11, 20] = new Cell(HS02);
            worksheet.Cells[cpt + 11, 21] = new Cell(A4111);
            worksheet.Cells[cpt + 11, 22] = new Cell(HS03);
            worksheet.Cells[cpt + 11, 23] = new Cell(A4113);
            worksheet.Cells[cpt + 11, 24] = new Cell(A4115);
            worksheet.Cells[cpt + 11, 25] = new Cell(A4100);
            worksheet.Cells[cpt + 11, 26] = new Cell(A4159);
            worksheet.Cells[cpt + 11, 27] = new Cell(A4175);
            worksheet.Cells[cpt + 11, 28] = new Cell(A4224);
            worksheet.Cells[cpt + 11, 29] = new Cell(A4160);
            worksheet.Cells[cpt + 11, 30] = new Cell(A4222);
            worksheet.Cells[cpt + 11, 31] = new Cell(A4340);
            worksheet.Cells[cpt + 11, 32] = new Cell(A4360);
            worksheet.Cells[cpt + 11, 33] = new Cell(A4361);
            worksheet.Cells[cpt + 11, 34] = new Cell(A4380);
            worksheet.Cells[cpt + 11, 35] = new Cell(A4381);
            worksheet.Cells[cpt + 11, 36] = new Cell(A7117);
            worksheet.Cells[cpt + 11, 37] = new Cell(A4410);
            worksheet.Cells[cpt + 11, 38] = new Cell(A4816);
            worksheet.Cells[cpt + 11, 39] = new Cell(A4817);
            worksheet.Cells[cpt + 11, 40] = new Cell(A4412);
            worksheet.Cells[cpt + 11, 41] = new Cell(A4450);
            worksheet.Cells[cpt + 11, 42] = new Cell(A4460);
            worksheet.Cells[cpt + 11, 43] = new Cell(A4550);
            worksheet.Cells[cpt + 11, 44] = new Cell(A4711);
            worksheet.Cells[cpt + 11, 45] = new Cell(A4712);
            worksheet.Cells[cpt + 11, 46] = new Cell(A4715);
            worksheet.Cells[cpt + 11, 47] = new Cell(A4716);
            worksheet.Cells[cpt + 11, 48] = new Cell(A4717);
            worksheet.Cells[cpt + 11, 49] = new Cell(A4790);
            worksheet.Cells[cpt + 11, 50] = new Cell(A4811);
            worksheet.Cells[cpt + 11, 51] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 52] = new Cell(A4813);
            worksheet.Cells[cpt + 11, 53] = new Cell(IFRT);
            worksheet.Cells[cpt + 11, 54] = new Cell(IndemRetAnt);
            worksheet.Cells[cpt + 11, 55] = new Cell(LiquidIfrt);
            worksheet.Cells[cpt + 11, 56] = new Cell(A5100);
            worksheet.Cells[cpt + 11, 57] = new Cell(A5420);
            worksheet.Cells[cpt + 11, 58] = new Cell(A5200);
            worksheet.Cells[cpt + 11, 59] = new Cell(A5400);
            worksheet.Cells[cpt + 11, 60] = new Cell(PrimTrim);
            worksheet.Cells[cpt + 11, 61] = new Cell(CongePris);
            worksheet.Cells[cpt + 11, 62] = new Cell(A4171);
            worksheet.Cells[cpt + 11, 63] = new Cell(HA04);
            worksheet.Cells[cpt + 11, 64] = new Cell(A4170);
            worksheet.Cells[cpt + 11, 65] = new Cell(A7114);
            worksheet.Cells[cpt + 11, 66] = new Cell(A7191);
            worksheet.Cells[cpt + 11, 67] = new Cell(A7193);
            worksheet.Cells[cpt + 11, 68] = new Cell(A7196);
            worksheet.Cells[cpt + 11, 69] = new Cell(A7692);
            worksheet.Cells[cpt + 11, 70] = new Cell(A7115);
            worksheet.Cells[cpt + 11, 71] = new Cell(A7118);
            worksheet.Cells[cpt + 11, 72] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 73] = new Cell(A9112);
            worksheet.Cells[cpt + 11, 74] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 75] = new Cell(A6120);
            worksheet.Cells[cpt + 11, 76] = new Cell(A6201);
            worksheet.Cells[cpt + 11, 77] = new Cell(A6221);
            worksheet.Cells[cpt + 11, 78] = new Cell(A9121);
            worksheet.Cells[cpt + 11, 79] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 80] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 81] = new Cell(A6840);
            worksheet.Cells[cpt + 11, 82] = new Cell(A6841);
            worksheet.Cells[cpt + 11, 83] = new Cell(A6835);
            worksheet.Cells[cpt + 11, 84] = new Cell(A6831);
            worksheet.Cells[cpt + 11, 85] = new Cell(A44509);
            worksheet.Cells[cpt + 11, 86] = new Cell(A6832);
            worksheet.Cells[cpt + 11, 87] = new Cell(A6834);
            worksheet.Cells[cpt + 11, 88] = new Cell(A6850);
            worksheet.Cells[cpt + 11, 89] = new Cell(A6410);
            worksheet.Cells[cpt + 11, 90] = new Cell(A6430);
            worksheet.Cells[cpt + 11, 91] = new Cell(A6470);
            worksheet.Cells[cpt + 11, 92] = new Cell(A6491);
            worksheet.Cells[cpt + 11, 93] = new Cell(A6520);
            worksheet.Cells[cpt + 11, 94] = new Cell(A6510);
            worksheet.Cells[cpt + 11, 95] = new Cell(A6521);
            worksheet.Cells[cpt + 11, 96] = new Cell(A6591);
            worksheet.Cells[cpt + 11, 97] = new Cell(A6614);
            worksheet.Cells[cpt + 11, 98] = new Cell(A8201);
            worksheet.Cells[cpt + 11, 99] = new Cell(A8215);
            worksheet.Cells[cpt + 11, 100] = new Cell(A8217);
            worksheet.Cells[cpt + 11, 101] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 102] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 103] = new Cell(A61109);
            worksheet.Cells[cpt + 11, 104] = new Cell(A61209);
            worksheet.Cells[cpt + 11, 105] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 106] = new Cell(A71159);
            worksheet.Cells[cpt + 11, 107] = new Cell(A62019);
            worksheet.Cells[cpt + 11, 108] = new Cell(RUB_COTEPY);
            worksheet.Cells[cpt + 11, 109] = new Cell(MoisPe);

            return worksheet;
        }


        public Worksheet EcrireSGNA021(int cpt, Worksheet worksheet, string matricule,
            string nom,
            string prenom,
                string Datedembauchesociete,
                        string Datededepartsociete,
                        string Intituledepartement,
                        string CodeCostcenter,
                        string Intitulecostcenter,
                        string Intituledelanatureducontrat,
                        string Emploioccupe,
                        string Salairedebase,
                        string Nbjoursdepresence,
                        string SalairedebaseJtrvaille,
                        string Inddepresence,
                        string InddeTransport,
                        string PrimeComplementairedePresence,
                        string IndemnitedeVoiture,
                        string IndFixederepresentation,
                        string Indemnitedexpatriation,
                      string HS125HS0,
                      string ValeurHS0,
                      string HS1502,
                       string ValeurHS02,
                       string HS175HS06,
                      string ValeurHS06,
                      string HS200HS03,
                       string ValeurHS03,
                       string HsupplementairenuitHS04,
                      string valeurHsuplementairenuitHS04,
                      string primeLVP,
                      string PrimeSIP,
                      string PrimeAOS,
                       string Bonus,
                     string Compensationsupport,
                      string Rappelsursalaire,
                      string Rappelprimedetransport,
                      string PrimeperequationRC,
                      string PrimeperequationAV,
                       string PrimeAid,
                      string Primemariage,
                      string Primedenaissance,
                      string Congespayes,
                      string Indemnitedepreavis,
                      string Gratificationfindeservice,
                      string Indemnitedelicenciement,
                      string Avticketsrestaurant,
                      string AvAssurancemaladie,
                      string Avassuranceretraitecomplementaire,
                      string Avcarburant,
                      string Avvoiture,
                      string Avscolariteenfants,
                     string Avutiliteexpat,
                     string Avlogement,
                       string SalaireBrut,
                       string BrutEricsson,
                      string CNSSEmploye,
                      string AssuranceretraitecomplemCNRPSemploye,
                      string AssuranceretraiteComplementaire,
                     string Salaireimposable,
                      string IRPP,
                      string Redevance1,
                      string PrêtCNSS,
                      string AutresRetenuesursalaire,
                      string DeductionAvTR,
                      string DeductionAvassurancemaladie,
                      string DeductionAvassurancecomplementaire,
                      string DeductionAvVoiture,
                      string DeductionAvCarburant,
                      string DeductionAvloyer,
                       string DeductionAvIndExpat,
                       string Deductionutiliteexpat,
                      string NetàPayer,
                      string CNSSEmployeur,
                      string Accidentdetravail,
                      string AssurancemaladieEmployeur,
                      string AssuranceComplementaireEmployeur,
                      string AssuranceretraitcompleCNRPSemplye,
                      string TFP,
                      string FOPROLOS,
                       string TotalContributionEmployeur,
                       string Moisdepaie
)
        {
   
            worksheet.Cells[cpt + 11, 0] = new Cell(matricule);
            worksheet.Cells[cpt + 11, 1] = new Cell(nom);
            worksheet.Cells[cpt + 11, 2] = new Cell(prenom);

            if (Datedembauchesociete != "")
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(form.DateAmeric(Datedembauchesociete));
            }
            else
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(Datedembauchesociete);
            }



            if (Datededepartsociete != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(Datededepartsociete));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(Datededepartsociete);
            }
           
            worksheet.Cells[cpt + 11, 5] = new Cell(Intituledepartement);
            worksheet.Cells[cpt + 11, 6] = new Cell( CodeCostcenter);
            worksheet.Cells[cpt + 11, 7] = new Cell( Intitulecostcenter);
            worksheet.Cells[cpt + 11, 8] = new Cell( Intituledelanatureducontrat);
            worksheet.Cells[cpt + 11, 9] = new Cell( Emploioccupe);
            worksheet.Cells[cpt + 11, 10] = new Cell( Salairedebase);
            worksheet.Cells[cpt + 11, 11] = new Cell( Nbjoursdepresence);
            worksheet.Cells[cpt + 11, 12] = new Cell( SalairedebaseJtrvaille);
            worksheet.Cells[cpt + 11, 13] = new Cell( Inddepresence);
            worksheet.Cells[cpt + 11, 14] = new Cell( InddeTransport);
            worksheet.Cells[cpt + 11, 15] = new Cell( PrimeComplementairedePresence);
            worksheet.Cells[cpt + 11, 16] = new Cell( IndemnitedeVoiture);
            worksheet.Cells[cpt + 11, 17] = new Cell( IndFixederepresentation);
            worksheet.Cells[cpt + 11, 18] = new Cell( Indemnitedexpatriation);
            worksheet.Cells[cpt + 11, 19] = new Cell(HS125HS0);
            worksheet.Cells[cpt + 11, 20] = new Cell(ValeurHS0);
            worksheet.Cells[cpt + 11, 21] = new Cell(HS1502);
            worksheet.Cells[cpt + 11, 22] = new Cell(ValeurHS02);
            worksheet.Cells[cpt + 11, 23] = new Cell(HS175HS06);
            worksheet.Cells[cpt + 11, 24] = new Cell (ValeurHS06);
            worksheet.Cells[cpt + 11, 25] = new Cell (HS200HS03);
            worksheet.Cells[cpt + 11, 26] = new Cell (ValeurHS03);
            worksheet.Cells[cpt + 11, 27] = new Cell (HsupplementairenuitHS04);
            worksheet.Cells[cpt + 11, 28] = new Cell (valeurHsuplementairenuitHS04);
            worksheet.Cells[cpt + 11, 29] = new Cell (primeLVP);
            worksheet.Cells[cpt + 11, 30] = new Cell (PrimeSIP);
            worksheet.Cells[cpt + 11, 31] = new Cell (PrimeAOS);
            worksheet.Cells[cpt + 11, 32] = new Cell(Bonus);
            worksheet.Cells[cpt + 11, 33] = new Cell (Compensationsupport);
            worksheet.Cells[cpt + 11, 34] = new Cell (Rappelsursalaire);
            worksheet.Cells[cpt + 11, 35] = new Cell (Rappelprimedetransport);
            worksheet.Cells[cpt + 11, 36] = new Cell (PrimeperequationRC);
            worksheet.Cells[cpt + 11, 37] = new Cell (PrimeperequationAV);
            worksheet.Cells[cpt + 11, 38] = new Cell(PrimeAid);
            worksheet.Cells[cpt + 11, 39] = new Cell (Primemariage);
            worksheet.Cells[cpt + 11, 40] = new Cell (Primedenaissance);
            worksheet.Cells[cpt + 11, 41] = new Cell (Congespayes);
            worksheet.Cells[cpt + 11, 42] = new Cell (Indemnitedepreavis);
            worksheet.Cells[cpt + 11, 43] = new Cell (Gratificationfindeservice);
            worksheet.Cells[cpt + 11, 44] = new Cell (Indemnitedelicenciement);
            worksheet.Cells[cpt + 11, 45] = new Cell (Avticketsrestaurant);
            worksheet.Cells[cpt + 11, 46] = new Cell (AvAssurancemaladie);
            worksheet.Cells[cpt + 11, 47] = new Cell (Avassuranceretraitecomplementaire);
            worksheet.Cells[cpt + 11, 48] = new Cell (Avcarburant);
            worksheet.Cells[cpt + 11, 49] = new Cell (Avvoiture);
            worksheet.Cells[cpt + 11, 50] = new Cell (Avscolariteenfants);
            worksheet.Cells[cpt + 11, 51] = new Cell (Avutiliteexpat);
            worksheet.Cells[cpt + 11, 52] = new Cell (Avlogement);
            worksheet.Cells[cpt + 11, 53] = new Cell(SalaireBrut);
            worksheet.Cells[cpt + 11, 54] = new Cell(BrutEricsson);
            worksheet.Cells[cpt + 11, 55] = new Cell (CNSSEmploye);
            worksheet.Cells[cpt + 11, 56] = new Cell (AssuranceretraitecomplemCNRPSemploye);
            worksheet.Cells[cpt + 11, 57] = new Cell (AssuranceretraiteComplementaire);
            worksheet.Cells[cpt + 11, 58] = new Cell (Salaireimposable);
            worksheet.Cells[cpt + 11, 59] = new Cell (IRPP);
            worksheet.Cells[cpt + 11, 60] = new Cell (Redevance1);
            worksheet.Cells[cpt + 11, 61] = new Cell (PrêtCNSS);
            worksheet.Cells[cpt + 11, 62] = new Cell (AutresRetenuesursalaire);
            worksheet.Cells[cpt + 11, 63] = new Cell (DeductionAvTR);
            worksheet.Cells[cpt + 11, 64] = new Cell (DeductionAvassurancemaladie);
            worksheet.Cells[cpt + 11, 65] = new Cell (DeductionAvassurancecomplementaire);
            worksheet.Cells[cpt + 11, 66] = new Cell (DeductionAvVoiture);
            worksheet.Cells[cpt + 11, 67] = new Cell (DeductionAvCarburant);
            worksheet.Cells[cpt + 11, 68] = new Cell (DeductionAvloyer);
            worksheet.Cells[cpt + 11, 69] = new Cell(DeductionAvIndExpat);
            worksheet.Cells[cpt + 11, 70] = new Cell(Deductionutiliteexpat);
            worksheet.Cells[cpt + 11, 71] = new Cell (NetàPayer);
            worksheet.Cells[cpt + 11, 72] = new Cell (CNSSEmployeur);
            worksheet.Cells[cpt + 11, 73] = new Cell (Accidentdetravail);
            worksheet.Cells[cpt + 11, 74] = new Cell (AssurancemaladieEmployeur);
            worksheet.Cells[cpt + 11, 75] = new Cell (AssuranceComplementaireEmployeur);
            worksheet.Cells[cpt + 11, 76] = new Cell (AssuranceretraitcompleCNRPSemplye);
            worksheet.Cells[cpt + 11, 77] = new Cell (TFP);
            worksheet.Cells[cpt + 11, 78] = new Cell (FOPROLOS);
            worksheet.Cells[cpt + 11, 79] = new Cell (TotalContributionEmployeur);
            worksheet.Cells[cpt + 11, 80] = new Cell(Moisdepaie);


            return worksheet;
        }





        public Worksheet EcrireSGNA016(int cpt, Worksheet worksheet, string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeDept, string IntituleDept,
      string NatureContrat, string Emploi, string SALBASE, string HORAIRE, string A1000, string A2100, string A2200, string A3413, string A3414, string A3415, string A3418,
      string A3416, string A3317, string A3510, string A3520, string A3530, string A3417, string A5910, string A3318, string A3610, string A3620, string A3630, string A3640, string A3651,
          string A3652, string A3650, string A3660, string A3661, string A3670, string A3680, string A3681, string A3682, string A3683, string A3684, string A3685, string BRUTFX, string A3686,
          string A3690, string CL31, string A4460, string CL32, string A4461, string CL33, string A4462, string CL34, string A4463, string CL35, string A3110,
          string CL37, string A4292, string A4293, string CL39, string A4294, string HC01, string A4295, string CL40, string A4296, string A4491, string A4162, string A4391, string A4381,
          string A4840, string A4382, string A4850, string A4383, string A4860, string A4718, string A4717, string A4713, string A4716, string HS01, string A4111, string HS02, string A4112,
          string HS03, string A4113, string HS04, string A4114, string HS05, string A4115, string HS06, string A4116, string HS07, string A4117, string A4172, string A4173,
          string HC02, string A4464, string A5100, string A5700, string A5702, string A5220, string A5400, string A5800, string A5900, string A7111, string A7114, string A7112,
          string A7116, string A7117, string A7113, string A5920, string A59209, string A7115, string A7600, string A7670, string A3910, string A4170, string A4811, string A4813, string A4812,
          string A4830, string BRUT, string A9110, string A9112, string A6110, string A6120, string A9121, string A6310, string A6312, string A6311, string A6462, string A6511, string A6430,
          string A6465, string A6467, string A6840, string A6810, string A6830, string A6860, string A6463, string A6610, string A6870, string A8140, string RUB_DEDSA, string NETPAIE,
          string A61109, string A61209, string A6150, string A6210, string TFP, string RUB_COTEPY, string RD_NET, string MoisPe)
        {
            worksheet.Cells[cpt + 11, 0] = new Cell(Matricule);
            worksheet.Cells[cpt + 11, 1] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 2] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 11, 5] = new Cell(CodeDept);
            worksheet.Cells[cpt + 11, 6] = new Cell(IntituleDept);
            worksheet.Cells[cpt + 11, 7] = new Cell(NatureContrat);
            worksheet.Cells[cpt + 11, 8] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 9] = new Cell(SALBASE);
            worksheet.Cells[cpt + 11, 10] = new Cell(HORAIRE);
            worksheet.Cells[cpt + 11, 11] = new Cell(A1000);
            worksheet.Cells[cpt + 11, 12] = new Cell(A2100);
            worksheet.Cells[cpt + 11, 13] = new Cell(A2200);
            worksheet.Cells[cpt + 11, 14] = new Cell(A3413);
            worksheet.Cells[cpt + 11, 15] = new Cell(A3414);
            worksheet.Cells[cpt + 11, 16] = new Cell(A3415);
            worksheet.Cells[cpt + 11, 17] = new Cell(A3418);
            worksheet.Cells[cpt + 11, 18] = new Cell(A3416);
            worksheet.Cells[cpt + 11, 19] = new Cell(A3317);
            worksheet.Cells[cpt + 11, 20] = new Cell(A3510);
            worksheet.Cells[cpt + 11, 21] = new Cell(A3520);
            worksheet.Cells[cpt + 11, 22] = new Cell(A3530);
            worksheet.Cells[cpt + 11, 23] = new Cell(A3417);
            worksheet.Cells[cpt + 11, 24] = new Cell(A5910);
            worksheet.Cells[cpt + 11, 25] = new Cell(A3318);
            worksheet.Cells[cpt + 11, 26] = new Cell(A3610);
            worksheet.Cells[cpt + 11, 27] = new Cell(A3620);
            worksheet.Cells[cpt + 11, 28] = new Cell(A3630);
            worksheet.Cells[cpt + 11, 29] = new Cell(A3640);
            worksheet.Cells[cpt + 11, 30] = new Cell(A3651);
            worksheet.Cells[cpt + 11, 31] = new Cell(A3652);
            worksheet.Cells[cpt + 11, 32] = new Cell(A3650);
            worksheet.Cells[cpt + 11, 33] = new Cell(A3660);
            worksheet.Cells[cpt + 11, 34] = new Cell(A3661);
            worksheet.Cells[cpt + 11, 35] = new Cell(A3670);
            worksheet.Cells[cpt + 11, 36] = new Cell(A3680);
            worksheet.Cells[cpt + 11, 37] = new Cell(A3681);
            worksheet.Cells[cpt + 11, 38] = new Cell(A3682);
            worksheet.Cells[cpt + 11, 39] = new Cell(A3683);
            worksheet.Cells[cpt + 11, 40] = new Cell(A3684);
            worksheet.Cells[cpt + 11, 41] = new Cell(A3685);
            worksheet.Cells[cpt + 11, 42] = new Cell(BRUTFX);
            worksheet.Cells[cpt + 11, 43] = new Cell(A3686);
            worksheet.Cells[cpt + 11, 44] = new Cell(A3690);
            worksheet.Cells[cpt + 11, 45] = new Cell(CL31);
            worksheet.Cells[cpt + 11, 46] = new Cell(A4460);
            worksheet.Cells[cpt + 11, 47] = new Cell(CL32);
            worksheet.Cells[cpt + 11, 48] = new Cell(A4461);
            worksheet.Cells[cpt + 11, 49] = new Cell(CL33);
            worksheet.Cells[cpt + 11, 50] = new Cell(A4462);
            worksheet.Cells[cpt + 11, 51] = new Cell(CL34);
            worksheet.Cells[cpt + 11, 52] = new Cell(A4463);
            worksheet.Cells[cpt + 11, 53] = new Cell(CL35);
            worksheet.Cells[cpt + 11, 54] = new Cell(A3110);
            worksheet.Cells[cpt + 11, 55] = new Cell(CL37);
            worksheet.Cells[cpt + 11, 56] = new Cell(A4292);
            worksheet.Cells[cpt + 11, 57] = new Cell(A4293);
            worksheet.Cells[cpt + 11, 58] = new Cell(CL39);
            worksheet.Cells[cpt + 11, 59] = new Cell(A4294);
            worksheet.Cells[cpt + 11, 60] = new Cell(HC01);
            worksheet.Cells[cpt + 11, 61] = new Cell(A4295);
            worksheet.Cells[cpt + 11, 62] = new Cell(CL40);
            worksheet.Cells[cpt + 11, 63] = new Cell(A4296);
            worksheet.Cells[cpt + 11, 64] = new Cell(A4491);
            worksheet.Cells[cpt + 11, 65] = new Cell(A4162);
            worksheet.Cells[cpt + 11, 66] = new Cell(A4391);
            worksheet.Cells[cpt + 11, 67] = new Cell(A4381);
            worksheet.Cells[cpt + 11, 68] = new Cell(A4840);
            worksheet.Cells[cpt + 11, 69] = new Cell(A4382);
            worksheet.Cells[cpt + 11, 70] = new Cell(A4850);
            worksheet.Cells[cpt + 11, 71] = new Cell(A4383);
            worksheet.Cells[cpt + 11, 72] = new Cell(A4860);
            worksheet.Cells[cpt + 11, 73] = new Cell(A4718);
            worksheet.Cells[cpt + 11, 74] = new Cell(A4717);
            worksheet.Cells[cpt + 11, 75] = new Cell(A4713);
            worksheet.Cells[cpt + 11, 76] = new Cell(A4716);
            worksheet.Cells[cpt + 11, 77] = new Cell(HS01);
            worksheet.Cells[cpt + 11, 78] = new Cell(A4111);
            worksheet.Cells[cpt + 11, 79] = new Cell(HS02);
            worksheet.Cells[cpt + 11, 80] = new Cell(A4112);
            worksheet.Cells[cpt + 11, 81] = new Cell(HS03);
            worksheet.Cells[cpt + 11, 82] = new Cell(A4113);
            worksheet.Cells[cpt + 11, 83] = new Cell(HS04);
            worksheet.Cells[cpt + 11, 84] = new Cell(A4114);
            worksheet.Cells[cpt + 11, 85] = new Cell(HS05);
            worksheet.Cells[cpt + 11, 86] = new Cell(A4115);
            worksheet.Cells[cpt + 11, 87] = new Cell(HS06);
            worksheet.Cells[cpt + 11, 88] = new Cell(A4116);
            worksheet.Cells[cpt + 11, 89] = new Cell(HS07);
            worksheet.Cells[cpt + 11, 90] = new Cell(A4117);
            worksheet.Cells[cpt + 11, 91] = new Cell(A4172);
            worksheet.Cells[cpt + 11, 92] = new Cell(A4173);
            worksheet.Cells[cpt + 11, 93] = new Cell(HC02);
            worksheet.Cells[cpt + 11, 94] = new Cell(A4464);
            worksheet.Cells[cpt + 11, 95] = new Cell(A5100);
            worksheet.Cells[cpt + 11, 96] = new Cell(A5700);
            worksheet.Cells[cpt + 11, 97] = new Cell(A5702);
            worksheet.Cells[cpt + 11, 98] = new Cell(A5220);
            worksheet.Cells[cpt + 11, 99] = new Cell(A5400);
            worksheet.Cells[cpt + 11, 100] = new Cell(A5800);
            worksheet.Cells[cpt + 11, 101] = new Cell(A5900);
            worksheet.Cells[cpt + 11, 102] = new Cell(A7111);
            worksheet.Cells[cpt + 11, 103] = new Cell(A7114);
            worksheet.Cells[cpt + 11, 104] = new Cell(A7112);
            worksheet.Cells[cpt + 11, 105] = new Cell(A7116);
            worksheet.Cells[cpt + 11, 106] = new Cell(A7117);
            worksheet.Cells[cpt + 11, 107] = new Cell(A7113);
            worksheet.Cells[cpt + 11, 108] = new Cell(A5920);
            worksheet.Cells[cpt + 11, 109] = new Cell(A59209);
            worksheet.Cells[cpt + 11, 110] = new Cell(A7115);
            worksheet.Cells[cpt + 11, 111] = new Cell(A7600);
            worksheet.Cells[cpt + 11, 112] = new Cell(A7670);
            worksheet.Cells[cpt + 11, 113] = new Cell(A3910);
            worksheet.Cells[cpt + 11, 114] = new Cell(A4170);
            worksheet.Cells[cpt + 11, 115] = new Cell(A4811);
            worksheet.Cells[cpt + 11, 116] = new Cell(A4813);
            worksheet.Cells[cpt + 11, 117] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 118] = new Cell(A4830);
            worksheet.Cells[cpt + 11, 119] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 120] = new Cell(A9110);
            worksheet.Cells[cpt + 11, 121] = new Cell(A9112);
            worksheet.Cells[cpt + 11, 122] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 123] = new Cell(A6120);
            worksheet.Cells[cpt + 11, 124] = new Cell(A9121);
            worksheet.Cells[cpt + 11, 125] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 126] = new Cell(A6312);
            worksheet.Cells[cpt + 11, 127] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 128] = new Cell(A6462);
            worksheet.Cells[cpt + 11, 129] = new Cell(A6511);
            worksheet.Cells[cpt + 11, 130] = new Cell(A6430);
            worksheet.Cells[cpt + 11, 131] = new Cell(A6465);
            worksheet.Cells[cpt + 11, 132] = new Cell(A6467);
            worksheet.Cells[cpt + 11, 133] = new Cell(A6840);
            worksheet.Cells[cpt + 11, 134] = new Cell(A6810);
            worksheet.Cells[cpt + 11, 135] = new Cell(A6830);
            worksheet.Cells[cpt + 11, 136] = new Cell(A6860);
            worksheet.Cells[cpt + 11, 137] = new Cell(A6463);
            worksheet.Cells[cpt + 11, 138] = new Cell(A6610);
            worksheet.Cells[cpt + 11, 139] = new Cell(A6870);
            worksheet.Cells[cpt + 11, 140] = new Cell(A8140);
            worksheet.Cells[cpt + 11, 141] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 142] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 143] = new Cell(A61109);
            worksheet.Cells[cpt + 11, 144] = new Cell(A61209);
            worksheet.Cells[cpt + 11, 145] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 146] = new Cell(A6210);
            worksheet.Cells[cpt + 11, 147] = new Cell(TFP);
            worksheet.Cells[cpt + 11, 148] = new Cell(RUB_COTEPY);
            worksheet.Cells[cpt + 11, 149] = new Cell(RD_NET);
            worksheet.Cells[cpt + 11, 150] = new Cell(MoisPe);

            return worksheet;
        }


  

        public Worksheet EcrireSGNA014(int cpt, Worksheet worksheet, string IdOracle, string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string LocalExpat, string Emploi,
          string IntituleDept, string CentreCout, string A1000, string A1910, string A1920, string A3310, string A3317, string A3711, string HS01, string A4110, string HS02, string A4111,
           string HS03, string A4113, string A4115, string A4100, string A4170, string A4159, string A4175, string A4224, string A4411, string A4816,string A4817, string A4222, string A4340, string A4360,
           string A4361, string A4380, string A4381, string A4410, string A4412, string A4450, string A4460, string A4550, string A4711, string A4712, string A4715, string A4716, string A4717,
           string A4790, string A4811,string A4610, string A4812, string A4813, string IFRT, string IndemRet, string Rappel,string A5200,string A5400, string A5420, string PrimeTrim, string CL03, string A4171, string A7193,
           string A7114, string A7117, string A7191, string A7192, string A7197, string A7196, string A7198, string A7199, string A7692, string A7115, string A7118, string BRUT, string A9112,
           string A6110, string A6120, string A6201, string A6221,string A6222, string A9121, string A6310, string A6311, string A7115A, string A7118A, string A7193A, string A7114A, string A6837, string A6832,
           string A7197A, string A6842, string A6843, string A6844, string A6850, string A6410, string A6430, string A6470, string A6491, string A6614, string A6530, string A6510, string A6521,
           string A6520, string A8201, string A8215, string A8217, string RUB_DEDSA, string NETPAIE, string A6110A, string A6120A, string A6150, string A7115AA, string A7118AA, string RUB_COTEPY, string MoisP)
        {

            worksheet.Cells[cpt + 11, 0] = new Cell(IdOracle);
            worksheet.Cells[cpt + 11, 1] = new Cell(Matricule);
            worksheet.Cells[cpt + 11, 2] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 3] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 5] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 5] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 11, 6] = new Cell(LocalExpat);
            worksheet.Cells[cpt + 11, 7] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 8] = new Cell(IntituleDept);
            worksheet.Cells[cpt + 11, 9] = new Cell(CentreCout);
            worksheet.Cells[cpt + 11, 10] = new Cell(A1000);
            worksheet.Cells[cpt + 11, 11] = new Cell(A1910);
            worksheet.Cells[cpt + 11, 12] = new Cell(A1920);
            worksheet.Cells[cpt + 11, 13] = new Cell(A3310);
            worksheet.Cells[cpt + 11, 14] = new Cell(A3317);
            worksheet.Cells[cpt + 11, 15] = new Cell(A3711);
            worksheet.Cells[cpt + 11, 16] = new Cell(HS01);
            worksheet.Cells[cpt + 11, 17] = new Cell(A4110);
            worksheet.Cells[cpt + 11, 18] = new Cell(HS02);
            worksheet.Cells[cpt + 11, 19] = new Cell(A4111);
            worksheet.Cells[cpt + 11, 20] = new Cell(HS03);
            worksheet.Cells[cpt + 11, 21] = new Cell(A4113);
            worksheet.Cells[cpt + 11, 22] = new Cell(A4115);
            worksheet.Cells[cpt + 11, 23] = new Cell(A4100);
            worksheet.Cells[cpt + 11, 24] = new Cell(A4170);
            worksheet.Cells[cpt + 11, 25] = new Cell(A4159);
            worksheet.Cells[cpt + 11, 26] = new Cell(A4175);
            worksheet.Cells[cpt + 11, 27] = new Cell(A4224);
            worksheet.Cells[cpt + 11, 28] = new Cell(A4411);
            worksheet.Cells[cpt + 11, 29] = new Cell(A4816);
            worksheet.Cells[cpt + 11, 30] = new Cell(A4817);
            worksheet.Cells[cpt + 11, 31] = new Cell(A4222);
            worksheet.Cells[cpt + 11, 32] = new Cell(A4340);
            worksheet.Cells[cpt + 11, 33] = new Cell(A4360);
            worksheet.Cells[cpt + 11, 34] = new Cell(A4361);
            worksheet.Cells[cpt + 11, 35] = new Cell(A4380);
            worksheet.Cells[cpt + 11, 36] = new Cell(A4381);
            worksheet.Cells[cpt + 11, 37] = new Cell(A4410);
            worksheet.Cells[cpt + 11, 38] = new Cell(A4412);
            worksheet.Cells[cpt + 11, 39] = new Cell(A4450);
            worksheet.Cells[cpt + 11, 40] = new Cell(A4460);
            worksheet.Cells[cpt + 11, 41] = new Cell(A4550);
            worksheet.Cells[cpt + 11, 42] = new Cell(A4711);
            worksheet.Cells[cpt + 11, 43] = new Cell(A4712);
            worksheet.Cells[cpt + 11, 44] = new Cell(A4715);
            worksheet.Cells[cpt + 11, 45] = new Cell(A4716);
            worksheet.Cells[cpt + 11, 46] = new Cell(A4717);
            worksheet.Cells[cpt + 11, 47] = new Cell(A4790);
            worksheet.Cells[cpt + 11, 48] = new Cell(A4811);
            worksheet.Cells[cpt + 11, 49] = new Cell(A4610);
            worksheet.Cells[cpt + 11, 50] = new Cell(A4812);
            worksheet.Cells[cpt + 11, 51] = new Cell(A4813);
            worksheet.Cells[cpt + 11, 52] = new Cell(IFRT);
            worksheet.Cells[cpt + 11, 53] = new Cell(IndemRet);
            worksheet.Cells[cpt + 11, 54] = new Cell(Rappel);
            worksheet.Cells[cpt + 11, 55] = new Cell(A5420);
            worksheet.Cells[cpt + 11, 56] = new Cell(A5200);
            worksheet.Cells[cpt + 11, 57] = new Cell(PrimeTrim);
            worksheet.Cells[cpt + 11, 58] = new Cell(A5400);
            worksheet.Cells[cpt + 11, 59] = new Cell(CL03);
            worksheet.Cells[cpt + 11, 60] = new Cell(A4171);
            worksheet.Cells[cpt + 11, 61] = new Cell(A7193);
            worksheet.Cells[cpt + 11, 62] = new Cell(A7114);
            worksheet.Cells[cpt + 11, 63] = new Cell(A7117);
            worksheet.Cells[cpt + 11, 64] = new Cell(A7191);
            worksheet.Cells[cpt + 11, 65] = new Cell(A7192);
            worksheet.Cells[cpt + 11, 66] = new Cell(A7197);
            worksheet.Cells[cpt + 11, 67] = new Cell(A7196);
            worksheet.Cells[cpt + 11, 68] = new Cell(A7198);
            worksheet.Cells[cpt + 11, 69] = new Cell(A7199);
            worksheet.Cells[cpt + 11, 70] = new Cell(A7692);
            worksheet.Cells[cpt + 11, 71] = new Cell(A7115);
            worksheet.Cells[cpt + 11, 72] = new Cell(A7118);
            worksheet.Cells[cpt + 11, 73] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 74] = new Cell(A9112);
            worksheet.Cells[cpt + 11, 75] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 76] = new Cell(A6120);
            worksheet.Cells[cpt + 11, 77] = new Cell(A6201);
            worksheet.Cells[cpt + 11, 78] = new Cell(A6221);
            worksheet.Cells[cpt + 11, 79] = new Cell(A6222);
            worksheet.Cells[cpt + 11, 80] = new Cell(A9121);
            worksheet.Cells[cpt + 11, 81] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 82] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 83] = new Cell(A7115A);
            worksheet.Cells[cpt + 11, 84] = new Cell(A7118A);
            worksheet.Cells[cpt + 11, 85] = new Cell(A7193A);
            worksheet.Cells[cpt + 11, 86] = new Cell(A7114A);
            worksheet.Cells[cpt + 11, 87] = new Cell(A6837);
            worksheet.Cells[cpt + 11, 88] = new Cell(A6832);
            worksheet.Cells[cpt + 11, 89] = new Cell(A7197A);
            worksheet.Cells[cpt + 11, 90] = new Cell(A6842);
            worksheet.Cells[cpt + 11, 91] = new Cell(A6843);
            worksheet.Cells[cpt + 11, 92] = new Cell(A6844);
            worksheet.Cells[cpt + 11, 93] = new Cell(A6850);
            worksheet.Cells[cpt + 11, 94] = new Cell(A6410);
            worksheet.Cells[cpt + 11, 95] = new Cell(A6430);
            worksheet.Cells[cpt + 11, 96] = new Cell(A6470);
            worksheet.Cells[cpt + 11, 97] = new Cell(A6491);
            worksheet.Cells[cpt + 11, 98] = new Cell(A6614);
            worksheet.Cells[cpt + 11, 99] = new Cell(A6530);
            worksheet.Cells[cpt + 11, 100] = new Cell(A6510);
            worksheet.Cells[cpt + 11, 101] = new Cell(A6521);
            worksheet.Cells[cpt + 11, 102] = new Cell(A6520);
            worksheet.Cells[cpt + 11, 103] = new Cell(A8201);
            worksheet.Cells[cpt + 11, 104] = new Cell(A8215);
            worksheet.Cells[cpt + 11, 105] = new Cell(A8217);
            worksheet.Cells[cpt + 11, 106] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 107] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 108] = new Cell(A6110A);
            worksheet.Cells[cpt + 11, 109] = new Cell(A6120A);
            worksheet.Cells[cpt + 11, 110] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 111] = new Cell(A7115AA);
            worksheet.Cells[cpt + 11, 112] = new Cell(A7118AA);
            worksheet.Cells[cpt + 11, 113] = new Cell(RUB_COTEPY);
            worksheet.Cells[cpt + 11, 114] = new Cell(MoisP);


            return worksheet;
           
        }

        public Worksheet EcrireSGNA015(int cpt, Worksheet worksheet, string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeDept,
           string IntituleDept, string NatContrat, string Emploi, string SalBaseAnn, string SALBASE_M, string CL01, string A1200, string A3410,
           string A3110, string A3711, string A4222, string A4362, string A4311, string A4380, string A4711, string A4717, string A4718, string A4364,
           string A4171,string A5100, string A7111, string A7112, string A7118, string BRUT, string A6110, string A6120, string A9121, string A6310, string A6311,string A8000,
           string A6850, string A6220, string RUB_DEDSA, string NETPAIE, string A6110A, string A6120A, string A6150, string RUB_EPYC,
           string MOIS_PAIE)
        {

            worksheet.Cells[cpt + 11, 0] = new Cell(Matricule);
            worksheet.Cells[cpt + 11, 1] = new Cell(Nom);
            worksheet.Cells[cpt + 11, 2] = new Cell(Prenom);
            if (DateEmbauche != "")
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(form.DateAmeric(DateEmbauche));
            }
            else
            {
                worksheet.Cells[cpt + 11, 3] = new Cell(DateEmbauche);
            }
            if (DateDepart != "")
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(form.DateAmeric(DateDepart));
            }
            else
            {
                worksheet.Cells[cpt + 11, 4] = new Cell(DateDepart);
            }
            worksheet.Cells[cpt + 11, 5] = new Cell(CodeDept);
            worksheet.Cells[cpt + 11, 6] = new Cell(IntituleDept);
            worksheet.Cells[cpt + 11, 7] = new Cell(NatContrat);
            worksheet.Cells[cpt + 11, 8] = new Cell(Emploi);
            worksheet.Cells[cpt + 11, 9] = new Cell(SalBaseAnn);
            worksheet.Cells[cpt + 11, 10] = new Cell(SALBASE_M);
            worksheet.Cells[cpt + 11, 11] = new Cell(CL01);
            worksheet.Cells[cpt + 11, 12] = new Cell(A1200);
            worksheet.Cells[cpt + 11, 13] = new Cell(A3410);
            worksheet.Cells[cpt + 11, 14] = new Cell(A3110);
            worksheet.Cells[cpt + 11, 15] = new Cell(A3711);
            worksheet.Cells[cpt + 11, 16] = new Cell(A4222);
            worksheet.Cells[cpt + 11, 17] = new Cell(A4362);
            worksheet.Cells[cpt + 11, 18] = new Cell(A4311);
            worksheet.Cells[cpt + 11, 19] = new Cell(A4380);
            worksheet.Cells[cpt + 11, 20] = new Cell(A4711);
            worksheet.Cells[cpt + 11, 21] = new Cell(A4717);
            worksheet.Cells[cpt + 11, 22] = new Cell(A4718);
            worksheet.Cells[cpt + 11, 23] = new Cell(A4364);
            worksheet.Cells[cpt + 11, 24] = new Cell(A4171);
            worksheet.Cells[cpt + 11, 25] = new Cell(A5100);
            worksheet.Cells[cpt + 11, 26] = new Cell(A7111);
            worksheet.Cells[cpt + 11, 27] = new Cell(A7112);
            worksheet.Cells[cpt + 11, 28] = new Cell(A7118);
            worksheet.Cells[cpt + 11, 29] = new Cell(BRUT);
            worksheet.Cells[cpt + 11, 30] = new Cell(A6110);
            worksheet.Cells[cpt + 11, 31] = new Cell(A6120);
            worksheet.Cells[cpt + 11, 32] = new Cell(A9121);
            worksheet.Cells[cpt + 11, 33] = new Cell(A6310);
            worksheet.Cells[cpt + 11, 34] = new Cell(A6311);
            worksheet.Cells[cpt + 11, 35] = new Cell(A8000);
            worksheet.Cells[cpt + 11, 36] = new Cell(A6850);
            worksheet.Cells[cpt + 11, 37] = new Cell(A6220);
            worksheet.Cells[cpt + 11, 38] = new Cell(RUB_DEDSA);
            worksheet.Cells[cpt + 11, 39] = new Cell(NETPAIE);
            worksheet.Cells[cpt + 11, 40] = new Cell(A6110A);
            worksheet.Cells[cpt + 11, 41] = new Cell(A6120A);
            worksheet.Cells[cpt + 11, 42] = new Cell(A6150);
            worksheet.Cells[cpt + 11, 43] = new Cell(RUB_EPYC);
            worksheet.Cells[cpt + 11, 44] = new Cell(MOIS_PAIE);

            return worksheet;
        }


    }
}
