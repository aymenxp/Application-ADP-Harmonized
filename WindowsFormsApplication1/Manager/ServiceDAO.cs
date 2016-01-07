using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace HarmonizedApp.Manager
{
    class ServiceDAO
    {

        public OleDbDataReader AffichageListeUsers()
        {
            return ManagerDAO.GetInstance().AffichageListeUsers();
        }

        public bool insertParamGlFile(string CodeRub,string LibCodRub, string CodeCpt, string Desc, string Sens,int IdSoc)
        {
            return ManagerDAO.GetInstance().insertParamGlFile(CodeRub,LibCodRub, CodeCpt, Desc, Sens,IdSoc);
        }

        public OleDbDataReader AffichageListeRubrique(int idSoc)
        {
            return ManagerDAO.GetInstance().AffichageListeRubrique(idSoc);
        }

        public OleDbDataReader AffichageListeGLFILE(int idSoc)
        {
            return ManagerDAO.GetInstance().AffichageListeGLFILE(idSoc);
        }

        public double SUMRetenue(string codeM, string NomTab, int Mois, int Annee, string Mat)
        {
            return ManagerDAO.GetInstance().SUMRetenue(codeM, NomTab, Mois, Annee, Mat);
        }

        public DataTable RecupererListeRub(int refSoc)
        {
            return ManagerDAO.GetInstance().RecupererListeRub(refSoc);
        }

        public DataTable ListeRubSVR(int refSoc, int cpt)
        {
            return ManagerDAO.GetInstance().ListeRubSVR(refSoc, cpt);
        }

        public double RecupererSumRubrique(int refSoc, int mois, int annee, string Code, string NomTab)
        {
            return ManagerDAO.GetInstance().RecupererSumRubrique(refSoc, mois, annee, Code, NomTab);
        }

        public DataTable ListeSceSGL(string NomTab, int Mois, int Annee, string CodeSce, string IntituleSce)
        {
            return ManagerDAO.GetInstance().ListeSceSGL(NomTab,Mois,Annee, CodeSce, IntituleSce);
        }

        public DataTable ListeOPR8(int Mois, int Annee)
        {
            return ManagerDAO.GetInstance().ListeOPR8(Mois, Annee);
        }

        public DataTable Rub11()
        {
            return ManagerDAO.GetInstance().Rub11();
        }

        public string RIBBGMOISPREC(string Mat, int MoisPr, int AnneePr)
        {
            return ManagerDAO.GetInstance().RIBBGMOISPREC(Mat, MoisPr, AnneePr);
        }

        public string ValRubA016(string Mat, int MoisPr, int AnneePr, string Rub)
        {
            return ManagerDAO.GetInstance().ValRubA016(Mat, MoisPr, AnneePr, Rub);
        }

        public string NETPAIEBGMOISPREC(string Mat, int MoisPr, int AnneePr)
        {
            return ManagerDAO.GetInstance().NETPAIEBGMOISPREC(Mat, MoisPr, AnneePr);
        }

        public string SALBASEBGMOISPREC(string Mat, int MoisPr, int AnneePr)
        {
            return ManagerDAO.GetInstance().SALBASEBGMOISPREC(Mat, MoisPr, AnneePr);
        }

        public DataTable ListDeGlFileData(string COSTCENTER, string GLACCOUNT,string Sens)
        {
            return ManagerDAO.GetInstance().ListDeGlFileData(COSTCENTER, GLACCOUNT,Sens);
        }

        public DataTable ListGlAccount()
        {
            return ManagerDAO.GetInstance().ListGlAccount();
        }

        public DataTable CodeCptF(string COSTCENTER)
        {
            return ManagerDAO.GetInstance().CodeCptF(COSTCENTER);
        }

        public double ValeurFinalSens(string GLACCOUNT, string COSTCENTER, string SENS)
        {
            return ManagerDAO.GetInstance().ValeurFinalSens(GLACCOUNT, COSTCENTER, SENS);
        }

        public DataTable ListeRubCptSGL(int IdSoc)
        {
            return ManagerDAO.GetInstance().ListeRubCptSGL(IdSoc);
        }

        public double ValeurCodeGLFILE(string Valeur, string Table, int Mois, int Annee,string CodeSce,string CentreC)
        {
            return ManagerDAO.GetInstance().ValeurCodeGLFILE(Valeur, Table, Mois, Annee,CodeSce,CentreC);
        }

        public DataTable ListeEmployee(string table, int mois, int annee)
        {
            return ManagerDAO.GetInstance().ListeEmployee(table, mois, annee);
        }

        public double RecupererSVE(int refSoc, int mois, int annee, string Code, string NomTab, string mat)
        {
            return ManagerDAO.GetInstance().RecupererSVE(refSoc, mois, annee, Code, NomTab, mat);
        }

        public double RecupererSVEIdOra(int refSoc, int mois, int annee, string Code, string NomTab, string mat)
        {
            return ManagerDAO.GetInstance().RecupererSVEIdOra(refSoc, mois, annee, Code, NomTab, mat);
        }

        public double RecupererValeur(int refSoc, int mois, int annee, string Code,string NomTab)
        {
            return ManagerDAO.GetInstance().RecupererValeur(refSoc, mois, annee, Code,NomTab);
        }

        public double RecupererVF(int refSoc, int mois, int annee, string Code, string NomTab, string Mat)
        {
            return ManagerDAO.GetInstance().RecupererVF(refSoc, mois, annee, Code, NomTab, Mat);
        }

        public bool ViderTabCodeMemo()
        {
            return ManagerDAO.GetInstance().ViderTabCodeMemo();
        }

        public bool insertPaymentReport(string CODEMEMO, string VALEUR, string Matricule, string idSoc)
        {
            return ManagerDAO.GetInstance().insertPaymentReport(CODEMEMO, VALEUR, Matricule, idSoc);
        }

        public DataTable RecupererListe2330(int refSoc)
        {
            return ManagerDAO.GetInstance().RecupererListe2330(refSoc);
        }

        public int OrdreSeq(int IdSoc)
        {
            return ManagerDAO.GetInstance().OrdreSeq(IdSoc);
        }

        public DataTable RecupererDATASGN(int refSoc,int MoisPe,int AnnePe)
        {
            return ManagerDAO.GetInstance().RecupererDATASGN(refSoc,MoisPe,AnnePe);
        }

        public DataTable RecupererDATASGNA002(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA002(refSoc, MoisPe, annee);
        }

        public DataTable RecupererDATASGNA003(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA003(refSoc, MoisPe, annee);
        }

        public DataTable RecupererDATASGNA004(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA004(refSoc, MoisPe, annee);
        }

        public DataTable RecupererDATASGNA021(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA021(refSoc, MoisPe, annee);
        }


        public DataTable RecupererDATASGNA010(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA010(refSoc, MoisPe, annee);
        }

        public DataTable RecupererDATASGNA013(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA013(refSoc, MoisPe, annee);
        }

        public DataTable RecupererDATASGNA020(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA020(refSoc, MoisPe, annee);
        }

        public DataTable RecupererDATASGNA016(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA016(refSoc, MoisPe, annee);
        }

        public DataTable RecupererDATASGNA011(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA011(refSoc, MoisPe, annee);
        }

        public DataTable RecupererDATASGNA014(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA014(refSoc, MoisPe, annee);
        }

        public DataTable RecupererDATASGNA015(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASGNA015(refSoc, MoisPe, annee);
        }

        public bool insertGrossToNetA010(string IdOracle, string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string Emploi,
          string NatureContrat, string CodeDept, string Dept, string HorBaseSal, string BullMod, string Categorie, string A1100, string A1300, string A2100,
           string A2200, string A3310, string ABS_CONGNR, string AbsNonRem, string RUB_ARR, string A4373, string HS01, string A4110, string HS02, string A4111,
           string HS03, string A4113, string HS06, string A4112, string HS04, string A4120, string HS07, string A4372, string HS05, string A4114, string HA04,
           string HA06, string RUB_JCHOM, string HA07, string RUB_TRP, string HA09, string A5500, string A4222, string A4371, string A4130, string A4135,
           string A2361, string A4816, string A4810, string HA01, string A4811, string A4812, string A5100, string A5400, string A5900, string A5200,
           string A6510, string A3309, string A4817, string A4818, string A4313, string A4312, string TESTSTC, string A4800, string PRES8, string PRES2,
           string A7115, string BRUT, string A6110, string A9310, string A6310, string A6311, string A6490, string A6491, string A6492, string A6840,
           string A6610, string A6611, string A6613, string A6612, string A6410, string A6511, string A6614, string RUB_DEDSA, string NETPAIE, string A6110A,
           string A6150, string A7115A, string RUB_COTEPY, string ModPe, string MoisPe, int IdSoc, int Mois, int Annee,string A4175,string A4374)
        {
            return ManagerDAO.GetInstance().insertGrossToNetA010(IdOracle, Matricule, Nom, Prenom, DateEmbauche, DateDepart, Emploi,
          NatureContrat, CodeDept, Dept, HorBaseSal, BullMod, Categorie, A1100, A1300, A2100,
           A2200, A3310, ABS_CONGNR, AbsNonRem, RUB_ARR, A4373, HS01, A4110, HS02, A4111,
           HS03, A4113, HS06, A4112, HS04, A4120, HS07, A4372, HS05, A4114, HA04,
           HA06, RUB_JCHOM, HA07, RUB_TRP, HA09, A5500, A4222, A4371, A4130, A4135,
           A2361, A4816, A4810, HA01, A4811, A4812, A5100, A5400, A5900, A5200,
           A6510, A3309, A4817, A4818, A4313, A4312, TESTSTC, A4800, PRES8, PRES2,
           A7115, BRUT, A6110, A9310, A6310, A6311, A6490, A6491, A6492, A6840,
           A6610, A6611, A6613, A6612, A6410, A6511, A6614, RUB_DEDSA, NETPAIE, A6110A,
           A6150, A7115A, RUB_COTEPY, ModPe, MoisPe, IdSoc, Mois, Annee,A4175,A4374);
        }


        public bool insertGrossToNetA014(string IdOracle, string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string LocalExpat, string Emploi,
          string IntituleDept, string CentreCout, string A1000, string A1910, string A1920, string A3310, string A3317, string A3711, string HS01, string A4110, string HS02, string A4111,
           string HS03, string A4113, string A4115, string A4100, string A4170, string A4159, string A4175, string A4224, string A4411, string A4816,string A4817, string A4222,string A4340, string A4360,
           string A4361, string A4380, string A4381, string A4410, string A4412, string A4450, string A4460, string A4550, string A4711, string A4712, string A4715, string A4716, string A4717,
           string A4790, string A4811, string A4812, string A4813, string IFRT, string IndemRet, string Rappel, string A5420, string A5200, string PrimeTrim, string CL03, string A4171, string A7193,
           string A7114, string A7117, string A7191, string A7192, string A7197, string A7196, string A7198, string A7199, string A7692, string A7115, string A7118, string BRUT, string A9112,
           string A6110, string A6120, string A6201, string A6221,string A6222, string A9121, string A6310, string A6311, string A7115A, string A7118A, string A7193A, string A7114A, string A6837, string A6832,
           string A7197A, string A6842, string A6843, string A6844, string A6850, string A6410, string A6430, string A6470, string A6491, string A6614, string A6530, string A6510, string A6521,
           string A6520, string A8201, string A8215, string A8217, string RUB_DEDSA, string NETPAIE, string A6110A, string A6120A, string A6150, string A7115AA, string A7118AA, string RUB_COTEPY,
           string MoisPe, int IdSoc, int Mois, int Annee,string A5400,string A4610)
        {
            return ManagerDAO.GetInstance().insertGrossToNetA014(IdOracle, Matricule, Nom, Prenom, DateEmbauche, DateDepart, LocalExpat, Emploi,
          IntituleDept, CentreCout, A1000, A1910, A1920, A3310, A3317, A3711, HS01, A4110, HS02, A4111,
           HS03, A4113, A4115, A4100, A4170, A4159, A4175, A4224, A4411, A4816,A4817, A4222, A4340, A4360,
           A4361, A4380, A4381, A4410, A4412, A4450, A4460, A4550, A4711, A4712, A4715, A4716, A4717,
           A4790, A4811, A4812, A4813, IFRT, IndemRet, Rappel, A5420, A5200, PrimeTrim, CL03, A4171, A7193,
           A7114, A7117, A7191, A7192, A7197, A7196, A7198, A7199, A7692, A7115, A7118, BRUT, A9112,
           A6110, A6120, A6201, A6221,A6222, A9121, A6310, A6311, A7115A, A7118A, A7193A, A7114A, A6837, A6832,
           A7197A, A6842, A6843, A6844, A6850, A6410, A6430, A6470, A6491, A6614, A6530, A6510, A6521,
           A6520, A8201, A8215, A8217, RUB_DEDSA, NETPAIE, A6110A, A6120A, A6150, A7115AA, A7118AA, RUB_COTEPY,
           MoisPe, IdSoc, Mois, Annee,A5400,A4610);
        }


        public bool insertGrossToNetA003(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeSce, string IntituleSce
         , string NatureContrat, string Emploi, string SalBase, string CL01, string A1200, string A2200, string A2100, string A2360,
         string A2330, string A3412, string A3317, string A3712, string A3710, string A3711, string A3713, string A4112,
         string A4113, string A4120, string A4222,string A4216, string A4361, string A4215, string A4320, string A4370, string A4430, string A7113, string A4171,
         string BRUT, string AjusCot, string A9112, string A6110, string A6120, string A9310, string A6310, string A6311,
         string A6340, string A6210, string A6410, string A6430, string A6520, string A6510, string A6530, string A6830,string A6490, string RUB_DEDSA,
         string NETPAIE, string A6110A, string A6120A, string A6150, string A6210A, string RUB_COTEPY, string MOIS_PAIE, string IdSoc, int Mois, int annee,
            string A4812,string A5100,string A4380)
        {
             return ManagerDAO.GetInstance().insertGrossToNetA003( Matricule, Nom, Prenom, DateEmbauche, DateDepart, CodeSce, IntituleSce,
         NatureContrat, Emploi, SalBase, CL01, A1200, A2200, A2100, A2360, A2330, A3412, A3317, A3712, A3710, A3711, A3713, A4112,
         A4113, A4120, A4222,A4216, A4361, A4215, A4320, A4370, A4430, A7113, A4171, BRUT, AjusCot, A9112, A6110, A6120, A9310, A6310, A6311,
         A6340, A6210, A6410, A6430, A6520, A6510, A6530, A6830,A6490, RUB_DEDSA, NETPAIE, A6110A, A6120A, A6150, A6210A, RUB_COTEPY, MOIS_PAIE,
         IdSoc, Mois, annee, A4812, A5100,A4380);
        }

        public bool insertGrossToNetA004(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeSce, string IntituleSce
        , string NatureContrat, string Emploi, string SalBase, string CL01, string A1100, string A2200, string A2100, string A3411,
        string A3410,string A3412, string A3317, string A3712, string A3710, string A3711, string A3713, string A4112, string A4113, string A4120,string A4216,
        string A4222, string A4361, string A4215, string A4320, string A4370, string A4430, string A7113, string BRUT, string AjusCot,
        string A9112, string A6110, string A6210, string A9310, string A6310, string A6311, string A6340, string A6210A, string A6410,
        string A6430, string A6520, string A6510, string A6530, string A6830,string A6490, string RUB_DEDSA, string NETPAIE, string A6110A,
           string A6120, string A6150, string A6210AA, string RUB_COTEPY, string MOIS_PAIE, string IdSoc, int Mois, int annee,string A4811,string A4171,string A4380,string A4812)
        {
            return ManagerDAO.GetInstance().insertGrossToNetA004(Matricule,Nom,Prenom,DateEmbauche,DateDepart,CodeSce,IntituleSce
        , NatureContrat, Emploi, SalBase, CL01, A1100, A2200, A2100, A3411,A3410, A3412, A3317, A3712, A3710, A3711, A3713, A4112,
        A4113, A4120,A4216, A4222, A4361, A4215, A4320, A4370, A4430, A7113, BRUT, AjusCot, A9112, A6110, A6210, A9310, A6310,
         A6311, A6340, A6210A, A6410, A6430, A6520, A6510, A6530, A6830,A6490, RUB_DEDSA, NETPAIE, A6110A, A6120, A6150, A6210AA, RUB_COTEPY,
         MOIS_PAIE, IdSoc, Mois,annee,A4811,A4171,A4380,A4812);
        }


        public bool insertGrossToNetA021(string matricule,
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
            return ManagerDAO.GetInstance().insertGrossToNetA021(    matricule,
             nom,
             prenom,
                 Datedembauchesociete ,
                         Datededepartsociete ,
                         Intituledepartement ,
                         CodeCostcenter ,
                         Intitulecostcenter ,
                         Intituledelanatureducontrat ,
                         Emploioccupe ,
                         Salairedebase ,
                         Nbjoursdepresence ,
                         SalairedebaseJtrvaille ,
                         Inddepresence ,
                         InddeTransport ,
                         PrimeComplementairedePresence ,
                         IndemnitedeVoiture ,
                         IndFixederepresentation ,
                         Indemnitedexpatriation ,
                        HS125HS0,
                        ValeurHS0,
                        HS1502,
                        ValeurHS02,
                        HS175HS06	,
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
                        PrimeAid	,
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
                        DeductionAvassurancemaladie	,
                        DeductionAvassurancecomplementaire	,
                        DeductionAvVoiture,
                        DeductionAvCarburant,
                        DeductionAvloyer,
                        DeductionAvIndExpat	,
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
        }



        public bool insertGrossToNetA011(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string IntituleDept, string IntituleSce,
          string IntituleCat, string NatContrat, string Emploi, string SalBaseSal, string CL01, string A1200, string A2100, string A2200, string A2311, string A3111,
           string A3413, string A3410, string A4351, string A4220, string A4211, string A4369, string A4360, string A4380, string A4717, string A4716, string A4110, string A4111,
           string A4113, string A4114, string A4812, string A4171, string BRUT, string A6110, string A9121, string A6310, string A6311, string TotDecEmp, string NETPAIE, string A6110A,
           string A6150, string A6320, string A6330, string RUB_EPYC, string TotCostLab, string MOIS_PAIE, string IdSoc, int Mois, int annee,string CodeDept)
        {
             return ManagerDAO.GetInstance().insertGrossToNetA011(Matricule,Nom, Prenom, DateEmbauche, DateDepart, IntituleDept, IntituleSce,
          IntituleCat, NatContrat, Emploi, SalBaseSal, CL01, A1200, A2100, A2200, A2311, A3111,
           A3413,A3410, A4351, A4220, A4211, A4369, A4360, A4380, A4717, A4716, A4110, A4111,
           A4113, A4114,A4812,A4171, BRUT, A6110, A9121, A6310, A6311, TotDecEmp, NETPAIE, A6110A,
           A6150, A6320, A6330, RUB_EPYC, TotCostLab, MOIS_PAIE, IdSoc, Mois, annee,CodeDept);
        }

        public bool insertGrossToNetA015(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeDept,
           string IntituleDept, string NatContrat, string Emploi, string SalBaseAnn, string SALBASE_M, string CL01, string A1200, string A3410,
           string A3110, string A3711, string A4222, string A4362, string A4311, string A4380, string A4711, string A4717, string A4718, string A4364,
           string A4171,string A5100, string A7111, string A7112, string A7118, string BRUT, string A6110, string A6120, string A9121, string A6310, string A6311,
           string A8000,string A6850, string A6220, string RUB_DEDSA, string NETPAIE, string A6110A, string A6120A, string A6150, string RUB_EPYC,
           string MOIS_PAIE, string IdSoc, int Mois, int annee)
        {
            return ManagerDAO.GetInstance().insertGrossToNetA015( Matricule, Nom, Prenom, DateEmbauche, DateDepart, CodeDept,
           IntituleDept, NatContrat, Emploi, SalBaseAnn, SALBASE_M, CL01, A1200, A3410,
           A3110, A3711, A4222, A4362, A4311, A4380, A4711, A4717, A4718, A4364,A4171,
           A5100, A7111, A7112, A7118, BRUT, A6110, A6120, A9121, A6310, A6311,
           A8000,A6850, A6220, RUB_DEDSA, NETPAIE, A6110A, A6120A, A6150, RUB_EPYC,
           MOIS_PAIE, IdSoc, Mois, annee);
        }

        public bool insertGrossToNetA013(string IdOracle,string Matricule,string Nom,string Prenom, string DateEmbauche,
                       string DateDept,string Emploi,string NatureContrat,string CodeDept,string Dept, string HorBaseSal,string BullMod,
                       string Categorie,string ModPe,string A1100,string A1300,string A1310,string A2100,string A2200,string A3310,string A3110,
                       string ABS_CONGNR,string A4174,string HS01,string A4110,string HS02,string A4111,string HS03,string A4113,string HS04,string A4120,
                       string HS07,string A4372,string HS05,string A4114,string HA04,string HA06,string A4175,string HA07,string A4176,string A4130,string A4135,
                       string A2361,string A4222,string A3309,string A4373,string A4371,string A4817,string A4818,string A4313,
                       string A4311,string HA09,string A5500,string A5100,string A5400,string A5200,string A5900,string HA01,string A4816,string A4810,string A4811,
                       string A4812,string TESTSTC,string A4800, string PRES8,string A7113,string CL23,string A7116,string A7115,
                       string BRUT,string A6110,string A9310,string A6310,string A6311,string A6614,string A6410,string A6490,string A6491,string A6492,string A6511,
                       string A6430,string A6431,string A6840,string RUB_DEDSA,string NETPAIE,string A61109,string A6150,string A71159,
                       string RUB_COTEPY,string MoisPe,int NumSoc,int Mois,int Annee,string A4374)
        {
            return ManagerDAO.GetInstance().insertGrossToNetA013(IdOracle, Matricule, Nom, Prenom, DateEmbauche,
                       DateDept, Emploi, NatureContrat, CodeDept, Dept, HorBaseSal, BullMod,
                       Categorie, ModPe, A1100, A1300, A1310, A2100, A2200, A3310, A3110, ABS_CONGNR, A4174, HS01, A4110, HS02, A4111, HS03, A4113, HS04, A4120,
                        HS07, A4372, HS05, A4114, HA04, HA06, A4175, HA07, A4176, A4130, A4135, A2361, A4222, A3309, A4373, A4371, A4817, A4818, A4313,
                        A4311, HA09, A5500, A5100, A5400, A5200, A5900, HA01, A4816, A4810, A4811, A4812, TESTSTC, A4800, PRES8, A7113, CL23, A7116, A7115,
                        BRUT, A6110, A9310, A6310, A6311, A6614, A6410, A6490, A6491, A6492, A6511, A6430, A6431, A6840, RUB_DEDSA, NETPAIE, A61109, A6150, A71159,
                        RUB_COTEPY, MoisPe, NumSoc, Mois, Annee,A4374);
        }

        public bool insertGrossToNetA020(string IdOracle, string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeDept, string IntituleDept,
       string ExpatLocal, string Emploi, string CentreCout, string A1000, string A1910, string A1920, string A2311, string A3317, string A3510, string A3711, string HS01, string A4110,
       string HS02, string A4111, string HS03, string A4113, string A4115, string A4100, string A4159, string A4175, string A4224, string A4160, string A4222, string A4340, string A4360,
           string A4361, string A4380, string A4381, string A7117, string A4410, string A4816,string A4817, string A4412, string A4450, string A4460, string A4550, string A4711, string A4712,
           string A4715, string A4716, string A4717, string A4790, string A4811, string A4812, string A4813, string IFRT, string IndemRetAnt, string LiquidIfrt, string A5100,
           string A5420,string A5200, string PrimTrim, string CongePris, string A4171, string HA04, string A4170, string A7114, string A7191, string A7193, string A7196, string A7692, string A7115,
           string A7118, string BRUT, string A9112, string A6110, string A6120, string A6201, string A6221, string A9121, string A6310, string A6311, string A6840, string A6841, string A6835,
           string A6831, string A44509, string A6832, string A6834, string A6850, string A6410, string A6430, string A6470, string A6491, string A6520, string A6510, string A6521,
           string A6591, string A6614, string A8201, string A8215, string A8217, string RUB_DEDSA, string NETPAIE, string A61109, string A61209, string A6150, string A71159, string A62019,
           string RUB_COTEPY, string MoisPe, int IdSoc, int Mois, int Annee,string A5400)
        {
             return ManagerDAO.GetInstance().insertGrossToNetA020( IdOracle,  Matricule,  Nom,  Prenom,  DateEmbauche,  DateDepart,  CodeDept,  IntituleDept,
        ExpatLocal,  Emploi,  CentreCout,  A1000,  A1910,  A1920,  A2311,  A3317,  A3510,  A3711,  HS01,  A4110,
        HS02,  A4111,  HS03,  A4113,  A4115,  A4100,  A4159,  A4175,  A4224,  A4160,  A4222,  A4340,  A4360,
        A4361, A4380, A4381, A7117, A4410, A4816,A4817, A4412, A4450, A4460, A4550, A4711, A4712,
        A4715, A4716, A4717, A4790, A4811, A4812, A4813, IFRT, IndemRetAnt, LiquidIfrt, A5100,
        A5420,A5200, PrimTrim, CongePris, A4171, HA04, A4170, A7114, A7191, A7193, A7196, A7692, A7115,
        A7118, BRUT, A9112, A6110, A6120, A6201, A6221, A9121, A6310, A6311, A6840, A6841, A6835,
        A6831, A44509, A6832, A6834, A6850, A6410, A6430, A6470, A6491, A6520, A6510, A6521,
        A6591, A6614, A8201, A8215, A8217, RUB_DEDSA, NETPAIE, A61109, A61209, A6150, A71159, A62019,
        RUB_COTEPY, MoisPe, IdSoc, Mois, Annee,A5400);
        }

        public int NumberSalary(int mois, int annee, string NomTab)
        {
            return ManagerDAO.GetInstance().NumberSalary(mois, annee, NomTab);
        }

        public double SumNetPay(int mois, int annee, string NomTab)
        {
            return ManagerDAO.GetInstance().SumNetPay(mois, annee, NomTab);
        }

        public double SumBrutPay(int mois, int annee, string NomTab)
        {
            return ManagerDAO.GetInstance().SumBrutPay(mois, annee, NomTab);
        }

        public double SumNetPayA021(int mois, int annee, string NomTab)
        {
            return ManagerDAO.GetInstance().SumNetPayA021(mois, annee, NomTab);
        }

        public double SumBrutPayA021(int mois, int annee, string NomTab)
        {
            return ManagerDAO.GetInstance().SumBrutPayA021(mois, annee, NomTab);
        }


        public DataTable RecupererDATASEP(int refSoc, int MoisPe, int annee)
        {
            return ManagerDAO.GetInstance().RecupererDATASEP(refSoc, MoisPe, annee);
        }

        public DataTable RecupererListe7600(int refSoc)
        {
            return ManagerDAO.GetInstance().RecupererListe7600(refSoc);
        }

        public DataTable RecupererListe4813(int refSoc)
        {
            return ManagerDAO.GetInstance().RecupererListe4813(refSoc);
        }

        public bool insertIncrement(string codeSoc, int num)
        {
            return ManagerDAO.GetInstance().insertIncrement(codeSoc, num);
        }

        public OleDbDataReader AffichageListeCompany()
        {
            return ManagerDAO.GetInstance().AffichageListeCompany();
        }

        public int recapIncrement(string num)
        {
            return ManagerDAO.GetInstance().recapIncrement(num);
        }

        public List<DataRaison> GetSociteList()
        {
            return ManagerDAO.GetInstance().GetSociteList();
        }

        public String Conc(String type, int lentgh, String valeur)
        {
            return ManagerDAO.GetInstance().Conc(type, lentgh, valeur);
        }

        public String ConcChaine(String type, int lentgh, String valeur)
        {
            return ManagerDAO.GetInstance().ConcChaine(type, lentgh, valeur);
        }

        public bool modifIncrementation(string IdSoc, int NumIncrement)
        {
            return ManagerDAO.GetInstance().modifIncrementation(IdSoc, NumIncrement);
        }

        public DataTable AffichageParam(string codeSoc)
        {
            return ManagerDAO.GetInstance().AffichageParam(codeSoc);
        }


        public bool insertRaison(string rais, string codeSoc)
        {
            return ManagerDAO.GetInstance().insertRaison(rais, codeSoc);
        }

        public bool insertRubrique(string CodeLib, string Libelles, string idSoc, int NbMax)
        {
            return ManagerDAO.GetInstance().insertRubrique(CodeLib, Libelles, idSoc, NbMax);
        }

        public bool Delete_User(int mat)
        {
            return ManagerDAO.GetInstance().Delete_User(mat);
        }

        public bool Delete_PayReport(int idSoc, int mois, int annee)
        {
            return ManagerDAO.GetInstance().Delete_PayReport(idSoc, mois, annee);
        }

        public bool Delete_Rubric(int mat)
        {
            return ManagerDAO.GetInstance().Delete_Rubric(mat);
        }

        public bool Delete_GLFILE(int mat)
        {
            return ManagerDAO.GetInstance().Delete_GLFILE(mat);
        }

        public bool Delete_GrossToNet(string NomTab, int idSoc, int mois, int annee)
        {
            return ManagerDAO.GetInstance().Delete_GrossToNet(NomTab, idSoc, mois, annee);
        }

        public bool Delete_Company(int mat)
        {
            return ManagerDAO.GetInstance().Delete_Company(mat);
        }

        public bool ModifierLeUser(int RefUser, string Login, string Pass)
        {
            return ManagerDAO.GetInstance().ModifierLeUser(RefUser,Login, Pass);
        }

        public bool ModifierSociete(int Refer, string CodeRub, string LibellesRub)
        {
            return ManagerDAO.GetInstance().ModifierSociete(Refer, CodeRub, LibellesRub);
        }

        public bool ModifierRubric(int Refer, string code, string libelles)
        {
            return ManagerDAO.GetInstance().ModifierRubric(Refer, code, libelles);
        }

        public bool ModifierGlFile(int Refer, string CodeRub,string LibCptRub, string CodeCpt, string Libelles, string Sens)
        {
            return ManagerDAO.GetInstance().ModifierGlFile(Refer, CodeRub,LibCptRub, CodeCpt, Libelles, Sens);
        }

        public OleDbDataReader VerifierIdentif(string Login, string Pass)
        {
            return ManagerDAO.GetInstance().VerifierIdentif(Login, Pass);
        }

        public bool insertUserWKF(string login, string pass)
        {
            return ManagerDAO.GetInstance().insertUserWKF(login, pass);
        }

        public bool insertGrossToNetA001(string matricule,string nom,string prenom,string DateEmbauche,string DateDept,string CodeSce,string IntituleSce,
                               string NatureContrat,string Emploi,string A1100,string CL01,string A1200,string A2200,string A2100,string A2311,
                                string A2320,string A3318,string A3110,string A4311,string A4330,string A3111,string A3210,string A3317,string A3319,string A4100,
                                string A4171,string A4130,string A4223,string A4380,string A4711,string A4717,string A5100,string A7600,string A4811,string A4812,string A4813,
                                string BRUT,string AjusCot,string A9112,string A6110,string A6120,string A6210,string A6230,string A9121,string A6310,string A6311,
                                string A6614,string A6410,string A6520,string A6510,string A6530,string A6890,string A8000,string RUB_DEDSA,string NETPAIE,string A61109,string A61209,string A6150,
                                string A62109, string A62309, string A6320, string A6330, string RUB_COTEPY, string A6490, string MoisPe, string NumSoc, int Mois, int Annee,
            string A4712,string A4226)
        {
            return ManagerDAO.GetInstance().insertGrossToNetA001(matricule, nom, prenom, DateEmbauche, DateDept, CodeSce, IntituleSce,
                                 NatureContrat, Emploi, A1100, CL01, A1200, A2200, A2100, A2311,
                                 A2320, A3318, A3110, A4311, A4330, A3111, A3210, A3317, A3319, A4100,
                                 A4171,A4130, A4223, A4380, A4711, A4717, A5100, A7600, A4811, A4812, A4813, BRUT, AjusCot,
                                A9112, A6110, A6120, A6210, A6230, A9121, A6310, A6311, A6614, A6410, A6520,A6510, A6530,
                                 A6890, RUB_DEDSA, NETPAIE, A61109, A61209, A6150, A62109, A62309, A6320, A6330,A8000, RUB_COTEPY,A6490,
                                 MoisPe, NumSoc.ToString(), Mois, Annee, A4712,A4226);
        }

        public bool insertGrossToNetA002(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeSce, string IntituleSce
          , string NatureContrat, string Emploi, string SalBase, string NbJrsPres, string NbHrPres, string SalBaseJrsTrav, string IndemTrans, string IndemPres,
          string CongePayer, string CongSpecio, string CongAbcNonRem, string PrimExcepFix, string CongAPayer, string HSupp125, string HSupp150, string HSupp200,
          string PrimAstreint, string PrimDim, string PrimDiff, string PrimDiffFor, string PrimParr, string PrimFormation, string PrimException, string PrimObjectif, string PrimNaiss, string PrimMariage,
          string PrimDeces, string PrimAid, string RappelEltRec, string RappelEltNonRec, string AvantNatTR, string RappelAvNatTr, string AvanNatAssGrp, string IndemPreavis,
          string GratifFinS, string IndemLicen,string A4815, string SalBrut, string SalBrutCot, string CNSSEmp, string RtnAssGrp, string SalImposable, string IRPP, string Redev,
          string ContConj, string AvancSal, string PretSoc, string RtnTiketRes, string AutreDed, string DedAvanNat, string DedAvanNatTR, string DedAvanNatAssGr,string A8800,
          string TotDedEmp, string SalNetPay, string CNSSEmply, string AcciTrav, string AssGrEmp, string totContPat, string MOIS_PAIE, string IdSoc, int Mois, int annee,string A5100,
            string A7116,string A6850)
        {
            return ManagerDAO.GetInstance().insertGrossToNetA002(Matricule, Nom, Prenom, DateEmbauche, DateDepart, CodeSce, IntituleSce
          , NatureContrat, Emploi, SalBase, NbJrsPres, NbHrPres, SalBaseJrsTrav, IndemTrans, IndemPres,
          CongePayer, CongSpecio, CongAbcNonRem, PrimExcepFix, CongAPayer, HSupp125, HSupp150, HSupp200,
          PrimAstreint, PrimDim, PrimDiff, PrimDiffFor, PrimParr, PrimFormation, PrimException, PrimObjectif, PrimNaiss, PrimMariage,
          PrimDeces, PrimAid, RappelEltRec, RappelEltNonRec, AvantNatTR, RappelAvNatTr, AvanNatAssGrp, IndemPreavis,
          GratifFinS, IndemLicen,A4815, SalBrut, SalBrutCot, CNSSEmp, RtnAssGrp, SalImposable, IRPP, Redev,
          ContConj, AvancSal, PretSoc, RtnTiketRes, AutreDed, DedAvanNat, DedAvanNatTR, DedAvanNatAssGr,A8800,
          TotDedEmp, SalNetPay, CNSSEmply, AcciTrav, AssGrEmp, totContPat, MOIS_PAIE, IdSoc, Mois, annee,A5100,A7116,A6850);
        }

        public bool insertPaymentReport(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string ModePe, string InfoBank, string NetPaie,
           int RefSoc, int MoisPaie, int AnneePaie)
        {
            return ManagerDAO.GetInstance().insertPaymentReport(Matricule, Nom, Prenom, DateEmbauche, DateDepart, ModePe, InfoBank, NetPaie, RefSoc, MoisPaie, AnneePaie);
        }

        public bool insertDataFixe(string IDOracle, string IDLocale, string Civilite, string CiviliteAbrege, string CiviliteValeur, string Sexe, string SexeValeur,
          string Nom, string NomJeuneFille, string Prenom, string Adresse, string ComplAdresse, string Commune, string CodePostal,
          string Tel, string CodePays, string IntitulePays, string Nationalite, string IntituleNationalite, string CIN,
          string DateExpirationSejour, string CodePaysNaiss, string DateNaiss, string ComuneNaiss, string SitFam, string SitFamVal,
          string NbreEnfant, string EnfRenseignes, string NumCNSS, string CleCNSS, string DateEmbauche, string DateDepart,
          string EtabSal, string TypeEntreEtab, string DateSortieEtab, string CodeConvColSal, string DeptSal, string IntituleDept,
          string SceSal, string UniteSal, string IntituleUnite, string CategSal, string CodeFonction, string EmploiOccupee,string NatContrat, string DateDebContrat,
          string BullModeleSal, string TypeSal, string SalBaseSal, string HorBaseSal, string CongeResteAnnPrec,
          string ModePaie, string CodeBque, string CodeGuichet, string NumCpt, string CleRib, string NomGuichet,
           string NomBque, string LibCpte, string MtAssVie, string EnfACharge, string ChefFamille, string AdhesionAsGr, string Champs1,
           string champs2, string Champs3, string Champs4, string Champs5, string Champs6, string RtCalcule, string CodeSce,
           string IntituleSce, string Date1, int RefSoc)
        {
            return ManagerDAO.GetInstance().insertDataFixe(IDOracle, IDLocale, Civilite, CiviliteAbrege, CiviliteValeur, Sexe, SexeValeur, Nom, NomJeuneFille, Prenom,
                Adresse, ComplAdresse, Commune, CodePostal, Tel, CodePays, IntitulePays, Nationalite, IntituleNationalite, CIN, DateExpirationSejour, CodePaysNaiss, DateNaiss,
                ComuneNaiss, SitFam, SitFamVal, NbreEnfant, EnfRenseignes, NumCNSS, CleCNSS, DateEmbauche, DateDepart, EtabSal, TypeEntreEtab, DateSortieEtab, CodeConvColSal,
                DeptSal, IntituleDept, SceSal, UniteSal, IntituleUnite, CategSal, CodeFonction, EmploiOccupee,NatContrat, DateDebContrat, BullModeleSal, TypeSal, SalBaseSal, HorBaseSal,
                CongeResteAnnPrec, ModePaie, CodeBque, CodeGuichet, NumCpt, CleRib, NomGuichet, NomBque, LibCpte, MtAssVie, EnfACharge, ChefFamille, AdhesionAsGr, Champs1, Champs1,
                Champs3, Champs4, Champs5, Champs6, RtCalcule, CodeSce, IntituleSce, Date1, RefSoc);
        }

        public string NomDuMois(int Mois)
        {
            return ManagerDAO.GetInstance().NomDuMois(Mois);
        }

        public string NomDuMoisEnglish(int Mois)
        {
            return ManagerDAO.GetInstance().NomDuMoisEnglish(Mois);
        }

        public bool Delete_DATAFIXE()
        {
            return ManagerDAO.GetInstance().Delete_DATAFIXE();
        }

        public DataTable RecupererData(int refSoc)
        {
            return ManagerDAO.GetInstance().RecupererData(refSoc);
        }

        public double BrutGross(int refSoc, int mois1)
        {
            return ManagerDAO.GetInstance().BrutGross(refSoc, mois1);
        }

        public double Gross1100(int refSoc, int mois1)
        {
            return ManagerDAO.GetInstance().Gross1100(refSoc, mois1);
        }

        public DataTable RecupererEmployee(int RefSoc)
        {
            return ManagerDAO.GetInstance().RecupererEmployee(RefSoc);
        }

        public double DataCodeMemo(string codeM, string Mat)
        {
            return ManagerDAO.GetInstance().DataCodeMemo(codeM, Mat);
        }

        public double CnssEmplyeurF(string codeM, string NomTab, int Mois, int Annee, string Defalquetion, string CodeSce)
        {
            return ManagerDAO.GetInstance().CnssEmplyeurF(codeM, NomTab, Mois, Annee, Defalquetion, CodeSce);
        }

        public bool insertDATAGLFILE(string CodeRub,string GLACCOUNT, string GLDESCRIPTION, string COSTCENTER, string Desc, string Sens, string valeur, int IdSoc)
        {
            return ManagerDAO.GetInstance().insertDATAGLFILE(CodeRub,GLACCOUNT, GLDESCRIPTION, COSTCENTER, Desc, Sens, valeur, IdSoc);
        }

        public double TesterSensCont(string GLACCOUNT, string Sens,string Sce)
        {
            return ManagerDAO.GetInstance().TesterSensCont(GLACCOUNT, Sens,Sce);
        }

        public bool Delete_SGLTAB()
        {
            return ManagerDAO.GetInstance().Delete_SGLTAB();
        }

        public double SUMCodeMemo(string codeM)
        {
            return ManagerDAO.GetInstance().SUMCodeMemo(codeM);
        }

        public int NbSal()
        {
            return ManagerDAO.GetInstance().NbSal();
        }

        public string NatureContrat(string CodeSal)
        {
            return ManagerDAO.GetInstance().NatureContrat(CodeSal);
        }

        public OleDbDataReader RecupererEmp(string Tab, int Mois, int Annee,string NatMat)
        {
            return ManagerDAO.GetInstance().RecupererEmp(Tab, Mois, Annee,NatMat);
        }

        public double ValeurCodeMemo(string Valeur, string Table, string Mat, int Mois, int Annee)
        {
            return ManagerDAO.GetInstance().ValeurCodeMemo(Valeur, Table, Mat, Mois, Annee);
        }

        public string SalNet(string CodeSal, string NomTab, int mois, int annee)
        {
            return ManagerDAO.GetInstance().SalNet(CodeSal,NomTab,mois,annee);
        }

        public string Avance(string CodeSal, string NomTab, int mois, int annee)
        {
            return ManagerDAO.GetInstance().Avance(CodeSal, NomTab, mois, annee);
        }

        public bool insertGrossToNetA016(string Matricule, string Nom, string Prenom, string DateEmbauche, string DateDepart, string CodeDept, string IntituleDept,
      string NatureContrat, string Emploi, string SALBASE, string HORAIRE, string A1000, string A2100, string A2200, string A3413, string A3414, string A3415, string A3418,
      string A3316, string A3317, string A3510, string A3520, string A3530, string A3417, string A5910, string A3318, string A3610, string A3620, string A3630, string A3640, string A3651,
          string A3652, string A3650, string A3660, string A3661, string A3670, string A3680, string A3681, string A3682, string A3683, string A3684, string A3685, string BRUTFX, string A3686,
          string A3690, string CL31, string A4460, string CL32, string A4461, string CL33, string A4462, string CL34, string A4463, string CL35, string A3110,
          string CL37, string A4292, string A4293, string CL39, string A4294, string HC01, string A4295, string CL40, string A4296, string A4491, string A4162, string A4391, string A4381,
          string A4840, string A4382, string A4850, string A4383, string A4860, string A4718, string A4717, string A4713, string A4716, string HS01, string A4111, string HS02, string A4112,
          string HS03, string A4113, string HS04, string A4114, string HS05, string A4115, string HS06, string A4116, string HS07, string A4117, string A4172, string A4173,
          string HC02, string A4464, string A5100, string A5700, string A5702, string A5220, string A5400, string A5800, string A5900, string A7111, string A7114, string A7112,
          string A7116, string A7117, string A7113, string A5920, string A59209, string A7115, string A7600, string A7670, string A3910, string A4170, string A4811, string A4813, string A4812,
          string A4830, string BRUT, string A9110, string A9112, string A6110, string A6120, string A9121, string A6310, string A6312, string A6311, string A6462, string A6511, string A6430,
          string A6465, string A6467, string A6840, string A6810, string A6830, string A6860, string A6463, string A6610, string A6870, string A8140, string RUB_DEDSA, string NETPAIE,
          string A61109, string A61209, string A6150, string A6210, string TFP, string RUB_COTEPY, string RD_NET, string MoisPe, int IdSoc, int Mois, int Annee, string RIB, string CUMNB2)
        {
            return ManagerDAO.GetInstance().insertGrossToNetA016( Matricule, Nom, Prenom, DateEmbauche, DateDepart, CodeDept, IntituleDept,
           NatureContrat,  Emploi,  SALBASE,  HORAIRE,  A1000,  A2100,  A2200,  A3413,  A3414,  A3415,  A3418,
           A3316,  A3317,  A3510,  A3520,  A3530,  A3417,  A5910,  A3318,  A3610,  A3620,  A3630,  A3640,  A3651,
           A3652,  A3650,  A3660,  A3661,  A3670,  A3680,  A3681,  A3682,  A3683,  A3684,  A3685,  BRUTFX,  A3686,
           A3690,  CL31,  A4460,  CL32,  A4461,  CL33,  A4462,  CL34,  A4463,  CL35,  A3110,
           CL37,  A4292,  A4293,  CL39,  A4294,  HC01,  A4295,  CL40,  A4296,  A4491,  A4162,  A4391,  A4381,
           A4840,  A4382,  A4850,  A4383,  A4860,  A4718,  A4717,  A4713,  A4716,  HS01,  A4111,  HS02,  A4112,
           HS03,  A4113,  HS04,  A4114,  HS05,  A4115,  HS06,  A4116,  HS07,  A4117,  A4172,  A4173,
           HC02,  A4464,  A5100,  A5700,  A5702,  A5220,  A5400,  A5800,  A5900,  A7111,  A7114,  A7112,
           A7116,  A7117,  A7113,  A5920,  A59209,  A7115,  A7600,  A7670,  A3910,  A4170,  A4811,  A4813,  A4812,
           A4830,  BRUT,  A9110,  A9112,  A6110,  A6120,  A9121,  A6310,  A6312,  A6311,  A6462,  A6511,  A6430,
           A6465,  A6467,  A6840,  A6810,  A6830,  A6860,  A6463,  A6610,  A6870,  A8140,  RUB_DEDSA,  NETPAIE,
           A61109, A61209, A6150, A6210, TFP, RUB_COTEPY, RD_NET, MoisPe, IdSoc, Mois, Annee, RIB, CUMNB2);
        }

    }
}
