using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelLibrary.SpreadSheet;
using System.Data;
using System.Windows.Forms;

namespace HarmonizedApp.Manager
{
    class EnteteSRF
    {

        ServiceDAO service = new ServiceDAO();
        FormatDate form = new FormatDate();

        public Worksheet EnteteSRFLigne1(Worksheet worksheet)
        {
            worksheet.Cells[0, 0] = new Cell("H0");
            worksheet.Cells[0, 1] = new Cell("DOC_ENCOD");
            worksheet.Cells[0, 2] = new Cell("DOC_CREATD");
            worksheet.Cells[0, 3] = new Cell("DOC_CREATT");
            worksheet.Cells[0, 4] = new Cell("DOC_TYPE");
            worksheet.Cells[0, 5] = new Cell("DOC_MVER");
            worksheet.Cells[0, 6] = new Cell("DOC_FVER");
            worksheet.Cells[0, 7] = new Cell("PTN_NAME");
            worksheet.Cells[0, 8] = new Cell("PTN_CID");
            worksheet.Cells[0, 9] = new Cell("CLT_LID");
            worksheet.Cells[0, 10] = new Cell("CLT_NAME");
            worksheet.Cells[0, 11] = new Cell("CLT_CID");
            worksheet.Cells[0, 12] = new Cell("CLT_CURID");
            worksheet.Cells[0, 13] = new Cell("PRL_PPD");
            worksheet.Cells[0, 14] = new Cell("PRL_PSD");
            worksheet.Cells[0, 15] = new Cell("PRL_PED");
            worksheet.Cells[0, 16] = new Cell("PRL_PSEQ");
            worksheet.Cells[0, 17] = new Cell("PRL_PDN");
            worksheet.Cells[0, 18] = new Cell("PRL_WTYPE");
            worksheet.Cells[0, 19] = new Cell("PRL_PRNUM");  
            return worksheet;
        }

        public string TypedePayroll(bool PeNormal, bool Bonus)
        {
            string TypedePe = string.Empty;
            if (PeNormal)
            {
                TypedePe = "Normal wages";
            }
            else
            {
                TypedePe = "Bonus";
            }
            return TypedePe;
        }

        public Worksheet EnteteSRFLigne2(Worksheet worksheet,string NomSoc,string Date1,string Date2,string CodeSoc,bool PeNormal,bool Bonus,string DatePe)
        {
            worksheet.Cells[1, 0] = new Cell("D0");
            worksheet.Cells[1, 1] = new Cell("UTF8");
            worksheet.Cells[1, 2] = new Cell((DateTime.Now.ToString().Substring(0,10)));
            string HeureMin = ExtraireHeure(DateTime.Now.Hour.ToString(),DateTime.Now.Minute.ToString(),DateTime.Now.Second.ToString());
            worksheet.Cells[1, 3] = new Cell(HeureMin);
            worksheet.Cells[1, 4] = new Cell("SRF");
            worksheet.Cells[1, 5] = new Cell("M2.0.2.0");
            worksheet.Cells[1, 6] = new Cell("F1.0.0");
            worksheet.Cells[1, 7] = new Cell("ABC");
            worksheet.Cells[1, 8] = new Cell("TN");
            worksheet.Cells[1, 9] = new Cell(CodeSoc);
            worksheet.Cells[1, 10] = new Cell(NomSoc);
            worksheet.Cells[1, 11] = new Cell("TN");
            worksheet.Cells[1, 12] = new Cell("TND");
            worksheet.Cells[1, 13] = new Cell(DatePe);
            worksheet.Cells[1, 14] = new Cell(Date1);
            worksheet.Cells[1, 15] = new Cell(Date2);
            worksheet.Cells[1, 16] = new Cell(1);
            int leMois = int.Parse(Date1.Substring(3, 2));
            string MoisC = service.NomDuMoisEnglish(leMois);
            string Year = Date1.Substring(6,4);
            worksheet.Cells[1, 17] = new Cell("payroll "+MoisC+" "+Year);
            string TypePay = TypedePayroll(PeNormal, Bonus);
            worksheet.Cells[1, 18] = new Cell(TypePay);
            worksheet.Cells[1, 19] = new Cell(0);       
            return worksheet;
        }

        public Worksheet EnteteSRFLigne3(Worksheet worksheet)
        {
            worksheet.Cells[2, 0] = new Cell("H1");
            worksheet.Cells[2, 1] = new Cell("EMP_PMOD");
            worksheet.Cells[2, 2] = new Cell("EMP_CEEID");
            worksheet.Cells[2, 3] = new Cell("EMP_PEEID");
            worksheet.Cells[2, 4] = new Cell("EMP_FNAME");
            worksheet.Cells[2, 5] = new Cell("EMP_MNAME");
            worksheet.Cells[2, 6] = new Cell("EMP_ALIAS");
            worksheet.Cells[2, 7] = new Cell("EMP_FST");
            worksheet.Cells[2, 8] = new Cell("EMP_NPID");
            worksheet.Cells[2, 9] = new Cell("EMP_TTL");
            worksheet.Cells[2, 10] = new Cell("EMP_DOB");
            worksheet.Cells[2, 11] = new Cell("EMP_COB");
            worksheet.Cells[2, 12] = new Cell("EMP_GND");
            worksheet.Cells[2, 13] = new Cell("EMP_MS_LC");
            worksheet.Cells[2, 14] = new Cell("EMP_MS_LBL");
            worksheet.Cells[2, 15] = new Cell("EMP_MS");
            worksheet.Cells[2, 16] = new Cell("EMP_NAT");
            worksheet.Cells[2, 17] = new Cell("EMP_NOC");
            worksheet.Cells[2, 18] = new Cell("EMP_NAD");
            worksheet.Cells[2, 19] = new Cell("EMP_AD1");
            worksheet.Cells[2, 20] = new Cell("EMP_AD2");
            worksheet.Cells[2, 21] = new Cell("EMP_AD3");
            worksheet.Cells[2, 22] = new Cell("EMP_ZIP");
            worksheet.Cells[2, 23] = new Cell("EMP_CITY");
            worksheet.Cells[2, 24] = new Cell("EMP_STATE");
            worksheet.Cells[2, 25] = new Cell("EMP_CTRY");
            worksheet.Cells[2, 26] = new Cell("EBK_IBAN01");
            worksheet.Cells[2, 27] = new Cell("EBK_BIC01");
            worksheet.Cells[2, 28] = new Cell("EBK_BBAN01");
            worksheet.Cells[2, 29] = new Cell("EBK_CURID01");
            worksheet.Cells[2, 30] = new Cell("EBK_OWN01");
            worksheet.Cells[2, 31] = new Cell("EBK_NAME01");
            worksheet.Cells[2, 32] = new Cell("EBK_IBAN02");
            worksheet.Cells[2, 33] = new Cell("EBK_BIC02");
            worksheet.Cells[2, 34] = new Cell("EBK_BBAN02");
            worksheet.Cells[2, 35] = new Cell("EBK_CURID02");
            worksheet.Cells[2, 36] = new Cell("EBK_OWN02");
            worksheet.Cells[2, 37] = new Cell("EBK_NAME02");
            worksheet.Cells[2, 38] = new Cell("MMV_BK01_PERCENT");
            worksheet.Cells[2, 39] = new Cell("MMV_BK01_AMOUNT");
            worksheet.Cells[2, 40] = new Cell("MMV_BK02_PERCENT");
            worksheet.Cells[2, 41] = new Cell("MMV_BK02_AMOUNT");
            worksheet.Cells[2, 42] = new Cell("JOB_TTL");
            worksheet.Cells[2, 43] = new Cell("JOB_TOC_LC");
            worksheet.Cells[2, 44] = new Cell("JOB_TOC_LBL");
            worksheet.Cells[2, 45] = new Cell("JOB_TOC");
            worksheet.Cells[2, 46] = new Cell("JOB_EMPST_LC");
            worksheet.Cells[2, 47] = new Cell("JOB_EMPST_LBL");
            worksheet.Cells[2, 48] = new Cell("JOB_EMPST");
            worksheet.Cells[2, 49] = new Cell("JOB_HIRR_LC");
            worksheet.Cells[2, 50] = new Cell("JOB_HIRR_LBL");
            worksheet.Cells[2, 51] = new Cell("JOB_HIRR");
            worksheet.Cells[2, 52] = new Cell("JOB_CTSD");
            worksheet.Cells[2, 53] = new Cell("JOB_CTED");
            worksheet.Cells[2, 54] = new Cell("JOB_EED");
            worksheet.Cells[2, 55] = new Cell("JOB_COSD");
            worksheet.Cells[2, 56] = new Cell("JOB_TERR_LC");
            worksheet.Cells[2, 57] = new Cell("JOB_TERR_LBL");
            worksheet.Cells[2, 58] = new Cell("JOB_TERR");
            worksheet.Cells[2, 59] = new Cell("JOB_CLASS_LC");
            worksheet.Cells[2, 60] = new Cell("JOB_CLASS_LBL");
            worksheet.Cells[2, 61] = new Cell("JOB_CLASS");
            worksheet.Cells[2, 62] = new Cell("JOB_PT");
            worksheet.Cells[2, 63] = new Cell("JOB_PTP");
            worksheet.Cells[2, 64] = new Cell("ORG_LEGAL_ID");
            worksheet.Cells[2, 65] = new Cell("ORG_LEGAL");
            worksheet.Cells[2, 66] = new Cell("ORG_BRANCH_ID");
            worksheet.Cells[2, 67] = new Cell("ORG_BRANCH");
            worksheet.Cells[2, 68] = new Cell("ORG_DEPART_ID");
            worksheet.Cells[2, 69] = new Cell("ORG_DEPARTMENT");
            worksheet.Cells[2, 70] = new Cell("ORG_SERVICE_ID");
            worksheet.Cells[2, 71] = new Cell("ORG_SERVICE");
            worksheet.Cells[2, 72] = new Cell("ORG_SITE_ID");
            worksheet.Cells[2, 73] = new Cell("ORG_SITE");
            worksheet.Cells[2, 74] = new Cell("ORG_CCID01");
            worksheet.Cells[2, 75] = new Cell("ORG_CCID02");
            worksheet.Cells[2, 76] = new Cell("ORG_CCID03");
            worksheet.Cells[2, 77] = new Cell("ORG_CCID04");
            worksheet.Cells[2, 78] = new Cell("ORG_CCID05");
            worksheet.Cells[2, 79] = new Cell("ORG_CCID06");
            worksheet.Cells[2, 80] = new Cell("ORG_CCID07");
            worksheet.Cells[2, 81] = new Cell("ORG_CCID08");
            worksheet.Cells[2, 82] = new Cell("ORG_CCID09");
            worksheet.Cells[2, 83] = new Cell("ORG_CCID10");
            worksheet.Cells[2, 84] = new Cell("ORG_CCNAME01");
            worksheet.Cells[2, 85] = new Cell("ORG_CCNAME02");
            worksheet.Cells[2, 86] = new Cell("ORG_CCNAME03");
            worksheet.Cells[2, 87] = new Cell("ORG_CCNAME04");
            worksheet.Cells[2, 88] = new Cell("ORG_CCNAME05");
            worksheet.Cells[2, 89] = new Cell("ORG_CCNAME06");
            worksheet.Cells[2, 90] = new Cell("ORG_CCNAME07");
            worksheet.Cells[2, 91] = new Cell("ORG_CCNAME08");
            worksheet.Cells[2, 92] = new Cell("ORG_CCNAME09");
            worksheet.Cells[2, 93] = new Cell("ORG_CCNAME10");
            worksheet.Cells[2, 94] = new Cell("ORG_CCSK01");
            worksheet.Cells[2, 95] = new Cell("ORG_CCSK02");
            worksheet.Cells[2, 96] = new Cell("ORG_CCSK03");
            worksheet.Cells[2, 97] = new Cell("ORG_CCSK04");
            worksheet.Cells[2, 98] = new Cell("ORG_CCSK05");
            worksheet.Cells[2, 99] = new Cell("ORG_CCSK06");
            worksheet.Cells[2, 100] = new Cell("ORG_CCSK07");
            worksheet.Cells[2, 101] = new Cell("ORG_CCSK08");
            worksheet.Cells[2, 102] = new Cell("ORG_CCSK09");
            worksheet.Cells[2, 103] = new Cell("ORG_CCSK10");
            worksheet.Cells[2, 104] = new Cell("GSL_BSS");
            worksheet.Cells[2, 105] = new Cell("GSL_SMS");
            worksheet.Cells[2, 106] = new Cell("GSL_SBC");
            worksheet.Cells[2, 107] = new Cell("GSL_OTC");
            worksheet.Cells[2, 108] = new Cell("GSL_BNS");
            worksheet.Cells[2, 109] = new Cell("GSL_COM");
            worksheet.Cells[2, 110] = new Cell("GSL_OVS");
            worksheet.Cells[2, 111] = new Cell("GSL_LEAV");
            worksheet.Cells[2, 112] = new Cell("GSL_HSD");
            worksheet.Cells[2, 113] = new Cell("GSL_HSC");
            worksheet.Cells[2, 114] = new Cell("GSL_LSD");
            worksheet.Cells[2, 115] = new Cell("GSL_LSC");
            worksheet.Cells[2, 116] = new Cell("GSL_NTHC");
            worksheet.Cells[2, 117] = new Cell("GSL_CCYN");
            worksheet.Cells[2, 118] = new Cell("GSL_CCBIK");
            worksheet.Cells[2, 119] = new Cell("GSL_OBIK");
            worksheet.Cells[2, 120] = new Cell("GSL_HALW");
            worksheet.Cells[2, 121] = new Cell("GSL_MALW");
            worksheet.Cells[2, 122] = new Cell("GSL_TALW");
            worksheet.Cells[2, 123] = new Cell("GSL_CALW");
            worksheet.Cells[2, 124] = new Cell("GSL_OALW");
            worksheet.Cells[2, 125] = new Cell("GSL_FEX");
            worksheet.Cells[2, 126] = new Cell("GSL_TPP");
            worksheet.Cells[2, 127] = new Cell("TOTAL_GSL");
            worksheet.Cells[2, 128] = new Cell("EED_WTX");
            worksheet.Cells[2, 129] = new Cell("EED_CSS");
            worksheet.Cells[2, 130] = new Cell("EED_CRT");
            worksheet.Cells[2, 131] = new Cell("EED_CUN");
            worksheet.Cells[2, 132] = new Cell("EED_CMC");
            worksheet.Cells[2, 133] = new Cell("EED_COTH");
            worksheet.Cells[2, 134] = new Cell("EED_VRT");
            worksheet.Cells[2, 135] = new Cell("EED_VMC");
            worksheet.Cells[2, 136] = new Cell("EED_VOTH");
            worksheet.Cells[2, 137] = new Cell("TOTAL_EED");
            worksheet.Cells[2, 138] = new Cell("TOTAL_NSL");
            worksheet.Cells[2, 139] = new Cell("NTP_SPPD");
            worksheet.Cells[2, 140] = new Cell("NTP_EXRF");
            worksheet.Cells[2, 141] = new Cell("NTP_EXPA");
            worksheet.Cells[2, 142] = new Cell("NTP_BIK");
            worksheet.Cells[2, 143] = new Cell("NTP_ONAD");
            worksheet.Cells[2, 144] = new Cell("NTP_SAA");
            worksheet.Cells[2, 145] = new Cell("NTP_TPP");
            worksheet.Cells[2, 146] = new Cell("TOTAL_NADJ");
            worksheet.Cells[2, 147] = new Cell("TOTAL_NTP");
            worksheet.Cells[2, 148] = new Cell("ERC_WTX");
            worksheet.Cells[2, 149] = new Cell("ERC_CSS");
            worksheet.Cells[2, 150] = new Cell("ERC_CRT");
            worksheet.Cells[2, 151] = new Cell("ERC_CUN");
            worksheet.Cells[2, 152] = new Cell("ERC_CMC");
            worksheet.Cells[2, 153] = new Cell("ERC_MCC");
            worksheet.Cells[2, 154] = new Cell("ERC_AWI");
            worksheet.Cells[2, 155] = new Cell("ERC_MCI");
            worksheet.Cells[2, 156] = new Cell("ERC_MTX");
            worksheet.Cells[2, 157] = new Cell("ERC_VOTH");
            worksheet.Cells[2, 158] = new Cell("TOTAL_ERC");
            worksheet.Cells[2, 159] = new Cell("ACR_HLS");
            worksheet.Cells[2, 160] = new Cell("ACR_HLC");
            worksheet.Cells[2, 161] = new Cell("ACR_BNS");
            worksheet.Cells[2, 162] = new Cell("ACR_BNC");
            worksheet.Cells[2, 163] = new Cell("ACR_SMS");
            worksheet.Cells[2, 164] = new Cell("ACR_SMC");
            worksheet.Cells[2, 165] = new Cell("TOTAL_ACR");
            worksheet.Cells[2, 166] = new Cell("EXT_CCBIK");
            worksheet.Cells[2, 167] = new Cell("EXT_OBIK");

            return worksheet;
        }

        public Worksheet EnteteSRFLigne4(Worksheet worksheet,int RefSoc,string ModePaie,string IDLocale,string MatSal,string Nom,string NomJeuneFille,string Prenom,string CIN,string CivilAbr,
        string DateNaiss,string CodePaysNaiss,string Gender,string sexe,string CivilVal,string Adr,string ComplAdr,string Commune,string CodeP,string SitFam,string SitFamVal,string NbreEnfant,
        string LibCpte, string RIB, string NomBque, string Difference, string Emploi, string NatContrat, string DateEmb, string DateSortie,string DeptSal,string IntituleDept,
        string NomSoc,string CodeSce,double GSL_BSS,double GSL_SMS,double GSL_SBC,double GSL_OTC,double GSL_BNS,double GSL_COM,
        double GSL_OVS,double GSL_LEAV,double GSL_HSD,double GSL_HSC,double GSL_LSD,double GSL_LSC,double GSL_NTHC,double GSL_CCYN,double GSL_CCBIK,double GSL_OBIK,
        double GSL_HALW,double GSL_MALW,double GSL_TALW,double GSL_CALW,double GSL_OALW,double GSL_FEX,double GSL_TPP,double EED_WTX,
        double EED_CSS,double EED_CRT,double EED_CUN,double EED_CMC,double EED_COTH,double EED_VRT,double EED_VMC,double EED_VOTH,double NTP_SPPD,
        double NTP_EXRF,double NTP_EXPA,double NTP_BIK,double NTP_ONAD,double NTP_SAA,double NTP_TPP,double ERC_WTX,double ERC_CSS,
        double ERC_CRT,double ERC_CUN,double ERC_CMC,double ERC_MCC,double ERC_AWI,double ERC_MCI,double ERC_MTX,double ERC_VOTH,double ACR_HLS,
        double ACR_HLC, double ACR_BNS, double ACR_BNC, double ACR_SMS, double ACR_SMC, double EXT_CCBIK,int i,string CodeSoc)
        {
                                   
                         worksheet.Cells[i+2, 0] = new Cell("D1");
                         string ModeFinal = DefinirModePe(ModePaie);
                         worksheet.Cells[i+2, 1] = new Cell(ModeFinal);
                         if (CodeSoc == "A001")
                         {
                             worksheet.Cells[i + 2, 2] = new Cell(IDLocale);
                             worksheet.Cells[i + 2, 3] = new Cell(CodeSoc + "_" + IDLocale);
                         }
                         else
                         {   
                             worksheet.Cells[i + 2, 2] = new Cell(MatSal);
                             worksheet.Cells[i + 2, 3] = new Cell(CodeSoc + "_" + MatSal);
                         }                     
                         worksheet.Cells[i + 2, 4] = new Cell(Nom);
                         worksheet.Cells[i + 2, 5] = new Cell(NomJeuneFille);
                         worksheet.Cells[i+2, 6] = new Cell("");
                         worksheet.Cells[i + 2, 7] = new Cell(Prenom);
                         worksheet.Cells[i + 2, 8] = new Cell(CIN);
                         worksheet.Cells[i + 2, 9] = new Cell(DataCivil(CivilAbr));
                         worksheet.Cells[i + 2, 10] = new Cell(form.DateAmeric(DateNaiss));
                         worksheet.Cells[i + 2, 11] = new Cell(CodePaysNaiss);
                         worksheet.Cells[i + 2, 12] = new Cell(DefinirGenre(Gender));
                         worksheet.Cells[i + 2, 13] = new Cell("");
                         worksheet.Cells[i + 2, 14] = new Cell(SitFam.Replace("�","e"));
                         worksheet.Cells[i + 2, 15] = new Cell(LaValeurCivil(SitFamVal));
                         worksheet.Cells[i + 2, 16] = new Cell(CodePaysNaiss);
                         worksheet.Cells[i + 2, 17] = new Cell(NbreEnfant);
                         worksheet.Cells[i + 2, 18] = new Cell("0");
                         worksheet.Cells[i + 2, 19] = new Cell(Adr);
                         worksheet.Cells[i + 2, 20] = new Cell(ComplAdr);
                         worksheet.Cells[i + 2, 21] = new Cell(Commune);
                         worksheet.Cells[i + 2, 22] = new Cell(CodeP);
                         worksheet.Cells[i + 2, 23] = new Cell(Commune);
                         worksheet.Cells[i + 2, 24] = new Cell("");
                         worksheet.Cells[i + 2, 25] = new Cell(CodePaysNaiss);
                         worksheet.Cells[i + 2, 26] = new Cell(LibCpte);
                         worksheet.Cells[i + 2, 27] = new Cell("");
                         worksheet.Cells[i + 2, 28] = new Cell("");
                         worksheet.Cells[i + 2, 29] = new Cell("TND");
                         worksheet.Cells[i + 2, 30] = new Cell("");
                         worksheet.Cells[i + 2, 31] = new Cell(NomBque);
                         worksheet.Cells[i + 2, 32] = new Cell("0");
                         worksheet.Cells[i + 2, 33] = new Cell("");
                         worksheet.Cells[i + 2, 34] = new Cell("");
                         worksheet.Cells[i + 2, 35] = new Cell("TND");
                         worksheet.Cells[i + 2, 36] = new Cell("");
                         worksheet.Cells[i + 2, 37] = new Cell("0");
                         worksheet.Cells[i + 2, 38] = new Cell("100%");
                         worksheet.Cells[i + 2, 39] = new Cell(Difference);
                         worksheet.Cells[i + 2, 40] = new Cell("0%");
                         worksheet.Cells[i + 2, 41] = new Cell("0");
                         worksheet.Cells[i + 2, 42] = new Cell(Emploi);
                         worksheet.Cells[i + 2, 43] = new Cell(NatContrat);
                         worksheet.Cells[i + 2, 44] = new Cell("");
                         worksheet.Cells[i + 2, 45] = new Cell(ContratInEnglish(NatContrat));
                         worksheet.Cells[i + 2, 46] = new Cell("");
                         worksheet.Cells[i + 2, 47] = new Cell("");
                         worksheet.Cells[i + 2, 48] = new Cell(ValeurFWKG(DateSortie));
                         worksheet.Cells[i + 2, 49] = new Cell("");
                         worksheet.Cells[i + 2, 50] = new Cell("");
                         worksheet.Cells[i + 2, 51] = new Cell("");
                         worksheet.Cells[i + 2, 52] = new Cell(form.DateAmeric(DateEmb));
                         worksheet.Cells[i + 2, 53] = new Cell("");
                         worksheet.Cells[i + 2, 54] = new Cell(form.DateAmeric(DateSortie));
                         worksheet.Cells[i + 2, 55] = new Cell("");
                         worksheet.Cells[i + 2, 56] = new Cell("");
                         worksheet.Cells[i + 2, 57] = new Cell("");
                         worksheet.Cells[i + 2, 58] = new Cell("");
                         worksheet.Cells[i + 2, 59] = new Cell("");
                         worksheet.Cells[i + 2, 60] = new Cell("");
                         worksheet.Cells[i + 2, 61] = new Cell(NatureEmploi(Emploi));
                         worksheet.Cells[i + 2, 62] = new Cell("FULL");
                         worksheet.Cells[i + 2, 63] = new Cell("100%");
                         worksheet.Cells[i + 2, 64] = new Cell("");
                         worksheet.Cells[i + 2, 65] = new Cell(NomSoc);
                         worksheet.Cells[i + 2, 66] = new Cell("");
                         worksheet.Cells[i + 2, 67] = new Cell("");
                         worksheet.Cells[i + 2, 68] = new Cell("");
                         worksheet.Cells[i + 2, 69] = new Cell("");
                         worksheet.Cells[i + 2, 70] = new Cell("");
                         worksheet.Cells[i + 2, 71] = new Cell("");
                         worksheet.Cells[i + 2, 72] = new Cell("");
                         worksheet.Cells[i + 2, 73] = new Cell("");
                         worksheet.Cells[i + 2, 74] = new Cell(DeptSal);
                         worksheet.Cells[i + 2, 75] = new Cell("");
                         worksheet.Cells[i + 2, 76] = new Cell("");
                         worksheet.Cells[i + 2, 77] = new Cell("");
                         worksheet.Cells[i + 2, 78] = new Cell("");
                         worksheet.Cells[i + 2, 79] = new Cell("");
                         worksheet.Cells[i + 2, 80] = new Cell("");
                         worksheet.Cells[i + 2, 81] = new Cell("");
                         worksheet.Cells[i + 2, 82] = new Cell("");
                         worksheet.Cells[i + 2, 83] = new Cell("");
                         worksheet.Cells[i + 2, 84] = new Cell(IntituleDept);
                         worksheet.Cells[i + 2, 85] = new Cell("");
                         worksheet.Cells[i + 2, 86] = new Cell("");
                         worksheet.Cells[i + 2, 87] = new Cell("");
                         worksheet.Cells[i + 2, 88] = new Cell("");
                         worksheet.Cells[i + 2, 89] = new Cell("");
                         worksheet.Cells[i + 2, 90] = new Cell("");
                         worksheet.Cells[i + 2, 91] = new Cell("");
                         worksheet.Cells[i + 2, 92] = new Cell("");
                         worksheet.Cells[i + 2, 93] = new Cell("");
                         worksheet.Cells[i + 2, 94] = new Cell("100%");
                         worksheet.Cells[i + 2, 95] = new Cell("");
                         worksheet.Cells[i + 2, 96] = new Cell("");
                         worksheet.Cells[i + 2, 97] = new Cell("");
                         worksheet.Cells[i + 2, 98] = new Cell("");
                         worksheet.Cells[i + 2, 99] = new Cell("");
                         worksheet.Cells[i + 2, 100] = new Cell("");
                         worksheet.Cells[i + 2, 101] = new Cell("");
                         worksheet.Cells[i + 2, 102] = new Cell("");
                         worksheet.Cells[i + 2, 103] = new Cell("");
                         worksheet.Cells[i + 2, 104] = new Cell(RetourValNull(GSL_BSS.ToString("#.##")));
                         worksheet.Cells[i + 2, 105] = new Cell(RetourValNull(GSL_SMS.ToString("#.##")));
                         worksheet.Cells[i + 2, 106] = new Cell(RetourValNull(GSL_SBC.ToString("#.##")));
                         worksheet.Cells[i + 2, 107] = new Cell(RetourValNull(GSL_OTC.ToString("#.##")));
                         worksheet.Cells[i + 2, 108] = new Cell(RetourValNull(GSL_BNS.ToString("#.##")));
                         worksheet.Cells[i + 2, 109] = new Cell(RetourValNull(GSL_COM.ToString("#.##")));
                         worksheet.Cells[i + 2, 110] = new Cell(RetourValNull(GSL_OVS.ToString("#.##")));
                         worksheet.Cells[i + 2, 111] = new Cell(RetourValNull(GSL_LEAV.ToString("#.##")));
                         worksheet.Cells[i + 2, 112] = new Cell(RetourValNull(GSL_HSD.ToString("#.##")));
                         worksheet.Cells[i + 2, 113] = new Cell(RetourValNull(GSL_HSC.ToString("#.##")));
                        if ((RefSoc == 4)||(RefSoc == 3))
                        {
                            worksheet.Cells[i + 2, 114] = new Cell(RetourValNull((GSL_LSD * (-1)).ToString("#.##")));
                        }
                        else
                        {
                            worksheet.Cells[i + 2, 114] = new Cell(RetourValNull(GSL_LSD.ToString("#.##")));
                        }

                        worksheet.Cells[i + 2, 115] = new Cell(RetourValNull(GSL_LSC.ToString("#.##")));
                        worksheet.Cells[i + 2, 116] = new Cell(RetourValNull(GSL_NTHC.ToString("#.##")));
                        worksheet.Cells[i + 2, 117] = new Cell("N");
                        worksheet.Cells[i + 2, 118] = new Cell(RetourValNull(GSL_CCBIK.ToString("#.##")));
                        worksheet.Cells[i + 2, 119] = new Cell(RetourValNull(GSL_OBIK.ToString("#.##")));
                        worksheet.Cells[i + 2, 120] = new Cell(RetourValNull(GSL_HALW.ToString("#.##")));
                        worksheet.Cells[i + 2, 121] = new Cell(RetourValNull(GSL_MALW.ToString("#.##")));
                        worksheet.Cells[i + 2, 122] = new Cell(RetourValNull(GSL_TALW.ToString("#.##")));
                        worksheet.Cells[i + 2, 123] = new Cell(RetourValNull(GSL_CALW.ToString("#.##")));
                        worksheet.Cells[i + 2, 124] = new Cell(RetourValNull(GSL_OALW.ToString("#.##")));
                        worksheet.Cells[i + 2, 125] = new Cell(RetourValNull(GSL_FEX.ToString("#.##")));
                        worksheet.Cells[i + 2, 126] = new Cell(RetourValNull(GSL_TPP.ToString("#.##")));
                        double TotVal = GSL_BSS + GSL_SMS + GSL_SBC + GSL_OTC + GSL_BNS + GSL_COM + GSL_OVS + GSL_LEAV + GSL_HSD + GSL_HSC + GSL_LSD + GSL_LSC + GSL_NTHC +
                        GSL_CCBIK + GSL_OBIK + GSL_HALW + GSL_MALW + GSL_TALW + GSL_CALW + GSL_OALW + GSL_FEX + GSL_TPP;
                        worksheet.Cells[i + 2, 127] = new Cell(RetourValNull(TotVal.ToString("#.##")));
                        worksheet.Cells[i + 2, 128] = new Cell(RetourValNull(EED_WTX.ToString("#.##")));
                        worksheet.Cells[i + 2, 129] = new Cell(RetourValNull(EED_CSS.ToString("#.##")));
                        worksheet.Cells[i + 2, 130] = new Cell(RetourValNull(EED_CRT.ToString("#.##")));
                        worksheet.Cells[i + 2, 131] = new Cell(RetourValNull(EED_CUN.ToString("#.##")));
                        worksheet.Cells[i + 2, 132] = new Cell(RetourValNull(EED_CMC.ToString("#.##")));
                        worksheet.Cells[i + 2, 133] = new Cell(RetourValNull(EED_COTH.ToString("#.##")));
                        worksheet.Cells[i + 2, 134] = new Cell(RetourValNull(EED_VRT.ToString("#.##")));
                        worksheet.Cells[i + 2, 135] = new Cell(RetourValNull(EED_VMC.ToString("#.##")));
                        worksheet.Cells[i + 2, 136] = new Cell(RetourValNull(EED_VOTH.ToString("#.##")));
                        double TOTAL_EED = EED_WTX + EED_CSS + EED_CRT + EED_CUN + EED_CMC + EED_COTH + EED_VRT + EED_VMC + EED_VOTH;
                        worksheet.Cells[i + 2, 137] = new Cell(RetourValNull(TOTAL_EED.ToString("#.##")));
                        double TOTAL_NSL = TotVal - TOTAL_EED;
                        worksheet.Cells[i + 2, 138] = new Cell(RetourValNull(TOTAL_NSL.ToString("#.##")));
                        worksheet.Cells[i + 2, 139] = new Cell(RetourValNull(NTP_SPPD.ToString("#.##")));
                        worksheet.Cells[i + 2, 140] = new Cell(RetourValNull(((NTP_EXRF) * (-1)).ToString("#.##")));
                        worksheet.Cells[i + 2, 141] = new Cell(RetourValNull(NTP_EXPA.ToString("#.##")));
                        worksheet.Cells[i + 2, 142] = new Cell(RetourValNull(((NTP_BIK) * (-1)).ToString("#.##")));
                        worksheet.Cells[i + 2, 143] = new Cell(RetourValNull(NTP_ONAD.ToString("#.##")));
                        worksheet.Cells[i + 2, 144] = new Cell(RetourValNull(((NTP_SAA) * (-1)).ToString("#.##")));
                        worksheet.Cells[i + 2, 145] = new Cell(RetourValNull(NTP_TPP.ToString("#.##")));
                        double TOTAL_NADJ = NTP_SPPD + NTP_EXRF + NTP_EXPA + NTP_BIK + NTP_ONAD + NTP_SAA + NTP_TPP;
                        worksheet.Cells[i + 2, 146] = new Cell(RetourValNull((TOTAL_NADJ*(-1)).ToString("#.##")));
                        double TOTAL_NTP = TOTAL_NSL - TOTAL_NADJ;
                        worksheet.Cells[i + 2, 147] = new Cell(RetourValNull(TOTAL_NTP.ToString("#.##")));
                        worksheet.Cells[i + 2, 148] = new Cell(RetourValNull(ERC_WTX.ToString("#.##")));
                        worksheet.Cells[i + 2, 149] = new Cell(RetourValNull(ERC_CSS.ToString("#.##")));
                        worksheet.Cells[i + 2, 150] = new Cell(RetourValNull(ERC_CRT.ToString("#.##")));
                        worksheet.Cells[i + 2, 151] = new Cell(RetourValNull(ERC_CUN.ToString("#.##")));
                        worksheet.Cells[i + 2, 152] = new Cell(RetourValNull(ERC_CMC.ToString("#.##")));
                        worksheet.Cells[i + 2, 153] = new Cell(RetourValNull(ERC_MCC.ToString("#.##")));
                        worksheet.Cells[i + 2, 154] = new Cell(RetourValNull(ERC_AWI.ToString("#.##")));
                        worksheet.Cells[i + 2, 155] = new Cell(RetourValNull(ERC_MCI.ToString("#.##")));
                        worksheet.Cells[i + 2, 156] = new Cell(RetourValNull(ERC_MTX.ToString("#.##")));
                        worksheet.Cells[i + 2, 157] = new Cell(RetourValNull(ERC_VOTH.ToString()));
                        double TOTAL_ERC = ERC_WTX + ERC_CSS + ERC_CRT + ERC_CUN + ERC_CMC + ERC_MCC + ERC_AWI + ERC_MCI + ERC_MTX + ERC_VOTH;
                        worksheet.Cells[i + 2, 158] = new Cell(RetourValNull(TOTAL_ERC.ToString("#.##")));
                        worksheet.Cells[i + 2, 159] = new Cell(RetourValNull(ACR_HLS.ToString("#.##")));
                        worksheet.Cells[i + 2, 160] = new Cell(RetourValNull(ACR_HLC.ToString("#.##")));
                        worksheet.Cells[i + 2, 161] = new Cell(RetourValNull(ACR_BNS.ToString("#.##")));
                        worksheet.Cells[i + 2, 162] = new Cell(RetourValNull(ACR_BNC.ToString("#.##")));
                        worksheet.Cells[i + 2, 163] = new Cell(RetourValNull(ACR_SMS.ToString("#.##")));
                        worksheet.Cells[i + 2, 164] = new Cell(RetourValNull(ACR_SMC.ToString("#.##")));
                        worksheet.Cells[i + 2, 165] = new Cell("0,00");
                        worksheet.Cells[i + 2, 166] = new Cell(RetourValNull(EXT_CCBIK.ToString("#.##")));
                        worksheet.Cells[i + 2, 167] = new Cell("0,00");


                 return worksheet;
             
        }

        public string ValeurFWKG(string DateSortie)
        {
            if ((DateSortie == "")||(DateSortie=="Vide"))
            {
                return "WKG";
            }
            else
            {
                return "TER";
            }
        }

        public Worksheet EnteteSRFLigneFooter(Worksheet worksheet, double GSL_BSS, double GSL_SMS, double GSL_SBC, double GSL_OTC, double GSL_BNS, double GSL_COM,
        double GSL_OVS, double GSL_LEAV, double GSL_HSD, double GSL_HSC, double GSL_LSD, double GSL_LSC, double GSL_NTHC, double GSL_CCYN, double GSL_CCBIK, double GSL_OBIK,
        double GSL_HALW, double GSL_MALW, double GSL_TALW, double GSL_CALW, double GSL_OALW, double GSL_FEX, double GSL_TPP, double EED_WTX,
        double EED_CSS, double EED_CRT, double EED_CUN, double EED_CMC, double EED_COTH, double EED_VRT, double EED_VMC, double EED_VOTH, double NTP_SPPD,
        double NTP_EXRF, double NTP_EXPA, double NTP_BIK, double NTP_ONAD, double NTP_SAA, double NTP_TPP, double ERC_WTX, double ERC_CSS,
        double ERC_CRT, double ERC_CUN, double ERC_CMC, double ERC_MCC, double ERC_AWI, double ERC_MCI, double ERC_MTX, double ERC_VOTH, double ACR_HLS,
        double ACR_HLC, double ACR_BNS, double ACR_BNC, double ACR_SMS, double ACR_SMC, double EXT_CCBIK, int k)
        {

            worksheet.Cells[3+k, 0] = new Cell("T1");
            worksheet.Cells[3 + k, 104] = new Cell(RetourValNull(GSL_BSS.ToString("#.##")));
            worksheet.Cells[3 + k, 105] = new Cell(RetourValNull(GSL_SMS.ToString("#.##")));
            worksheet.Cells[3 + k, 106] = new Cell(RetourValNull(GSL_SBC.ToString("#.##")));
            worksheet.Cells[3 + k, 107] = new Cell(RetourValNull(GSL_OTC.ToString("#.##")));
            worksheet.Cells[3 + k, 108] = new Cell(RetourValNull(GSL_BNS.ToString("#.##")));
            worksheet.Cells[3 + k, 109] = new Cell(RetourValNull(GSL_COM.ToString("#.##")));
            worksheet.Cells[3 + k, 110] = new Cell(RetourValNull(GSL_OVS.ToString("#.##")));
            worksheet.Cells[3 + k, 111] = new Cell(RetourValNull(GSL_LEAV.ToString("#.##")));
            worksheet.Cells[3 + k, 112] = new Cell(RetourValNull(GSL_HSD.ToString("#.##")));
            worksheet.Cells[3 + k, 113] = new Cell(RetourValNull(GSL_HSC.ToString("#.##")));
            worksheet.Cells[3 + k, 114] = new Cell(RetourValNull(GSL_LSD.ToString("#.##")));
            worksheet.Cells[3 + k, 115] = new Cell(RetourValNull(GSL_LSC.ToString("#.##")));
            worksheet.Cells[3 + k, 116] = new Cell(RetourValNull(GSL_NTHC.ToString("#.##")));
            worksheet.Cells[3 + k, 117] = new Cell("0,000");
            worksheet.Cells[3 + k, 118] = new Cell(RetourValNull(GSL_CCBIK.ToString("#.##")));
            worksheet.Cells[3 + k, 119] = new Cell(RetourValNull(GSL_OBIK.ToString("#.##")));
            worksheet.Cells[3 + k, 120] = new Cell(RetourValNull(GSL_HALW.ToString("#.##")));
            worksheet.Cells[3 + k, 121] = new Cell(RetourValNull(GSL_MALW.ToString("#.##")));
            worksheet.Cells[3 + k, 122] = new Cell(RetourValNull(GSL_TALW.ToString("#.##")));
            worksheet.Cells[3 + k, 123] = new Cell(RetourValNull(GSL_CALW.ToString("#.##")));
            worksheet.Cells[3 + k, 124] = new Cell(RetourValNull(GSL_OALW.ToString("#.##")));
            worksheet.Cells[3 + k, 125] = new Cell(RetourValNull(GSL_FEX.ToString("#.##")));
            worksheet.Cells[3 + k, 126] = new Cell(RetourValNull(GSL_TPP.ToString("#.##")));
            double TotVal = GSL_BSS + GSL_SMS + GSL_SBC + GSL_OTC + GSL_BNS + GSL_COM + GSL_OVS + GSL_LEAV + GSL_HSD + GSL_HSC + GSL_LSD + GSL_LSC + GSL_NTHC +
            GSL_CCBIK + GSL_OBIK + GSL_HALW + GSL_MALW + GSL_TALW + GSL_CALW + GSL_OALW + GSL_FEX + GSL_TPP;
            worksheet.Cells[3 + k, 127] = new Cell(RetourValNull(TotVal.ToString("#.##")));
            worksheet.Cells[3 + k, 128] = new Cell(RetourValNull(EED_WTX.ToString("#.##")));
            worksheet.Cells[3 + k, 129] = new Cell(RetourValNull(EED_CSS.ToString("#.##")));
            worksheet.Cells[3 + k, 130] = new Cell(RetourValNull(EED_CRT.ToString("#.##")));
            worksheet.Cells[3 + k, 131] = new Cell(RetourValNull(EED_CUN.ToString("#.##")));
            worksheet.Cells[3 + k, 132] = new Cell(RetourValNull(EED_CMC.ToString("#.##")));
            worksheet.Cells[3 + k, 133] = new Cell(RetourValNull(EED_COTH.ToString("#.##")));
            worksheet.Cells[3 + k, 134] = new Cell(RetourValNull(EED_VRT.ToString("#.##")));
            worksheet.Cells[3 + k, 135] = new Cell(RetourValNull(EED_VMC.ToString("#.##")));
            worksheet.Cells[3 + k, 136] = new Cell(RetourValNull(EED_VOTH.ToString("#.##")));
            double TOTAL_EED = EED_WTX + EED_CSS + EED_CRT + EED_CUN + EED_CMC + EED_COTH + EED_VRT + EED_VMC + EED_VOTH;
            worksheet.Cells[3 + k, 137] = new Cell(RetourValNull(TOTAL_EED.ToString("#.##")));
            double TOTAL_NSL = TotVal - TOTAL_EED;
            worksheet.Cells[3 + k, 138] = new Cell(RetourValNull(TOTAL_NSL.ToString("#.##")));
            worksheet.Cells[3 + k, 139] = new Cell(RetourValNull(NTP_SPPD.ToString("#.##")));
            worksheet.Cells[3 + k, 140] = new Cell(RetourValNull((NTP_EXRF*(-1)).ToString("#.##")));
            worksheet.Cells[3 + k, 141] = new Cell(RetourValNull(NTP_EXPA.ToString("#.##")));
            worksheet.Cells[3 + k, 142] = new Cell(RetourValNull((NTP_BIK*(-1)).ToString("#.##")));
            worksheet.Cells[3 + k, 143] = new Cell(RetourValNull(NTP_ONAD.ToString("#.##")));
            worksheet.Cells[3 + k, 144] = new Cell(RetourValNull((NTP_SAA*(-1)).ToString("#.##")));
            worksheet.Cells[3 + k, 145] = new Cell(RetourValNull(NTP_TPP.ToString("#.##")));
            double TOTAL_NADJ = NTP_SPPD + NTP_EXRF + NTP_EXPA + NTP_BIK + NTP_ONAD + NTP_SAA+ NTP_TPP;
            worksheet.Cells[3 + k, 146] = new Cell(RetourValNull((TOTAL_NADJ*(-1)).ToString("#.##")));
            double TOTAL_NTP = TOTAL_NSL + ((-1)*TOTAL_NADJ);
            worksheet.Cells[3 + k, 147] = new Cell(RetourValNull(TOTAL_NTP.ToString("#.##")));
            worksheet.Cells[3 + k, 148] = new Cell(RetourValNull(ERC_WTX.ToString("#.##")));
            worksheet.Cells[3 + k, 149] = new Cell(RetourValNull(ERC_CSS.ToString("#.##")));
            worksheet.Cells[3 + k, 150] = new Cell(RetourValNull(ERC_CRT.ToString("#.##")));
            worksheet.Cells[3 + k, 151] = new Cell(RetourValNull(ERC_CUN.ToString("#.##")));
            worksheet.Cells[3 + k, 152] = new Cell(RetourValNull(ERC_CMC.ToString("#.##")));
            worksheet.Cells[3 + k, 153] = new Cell(RetourValNull(ERC_MCC.ToString("#.##")));
            worksheet.Cells[3 + k, 154] = new Cell(RetourValNull(ERC_AWI.ToString("#.##")));
            worksheet.Cells[3 + k, 155] = new Cell(RetourValNull(ERC_MCI.ToString("#.##")));
            worksheet.Cells[3 + k, 156] = new Cell(RetourValNull(ERC_MTX.ToString("#.##")));
            worksheet.Cells[3 + k, 157] = new Cell(RetourValNull(ERC_VOTH.ToString("#.##")));
            double TOTAL_ERC = ERC_WTX + ERC_CSS + ERC_CRT + ERC_CUN + ERC_CMC + ERC_MCC + ERC_AWI + ERC_MCI + ERC_MTX + ERC_VOTH;
            worksheet.Cells[3 + k, 158] = new Cell(RetourValNull(TOTAL_ERC.ToString("#.##")));
            worksheet.Cells[3 + k, 159] = new Cell(RetourValNull(ACR_HLS.ToString("#.##")));
            worksheet.Cells[3 + k, 160] = new Cell(RetourValNull(ACR_HLC.ToString("#.##")));
            worksheet.Cells[3 + k, 161] = new Cell(RetourValNull(ACR_BNS.ToString("#.##")));
            worksheet.Cells[3 + k, 162] = new Cell(RetourValNull(ACR_BNC.ToString("#.##")));
            worksheet.Cells[3 + k, 163] = new Cell(RetourValNull(ACR_SMS.ToString("#.##")));
            worksheet.Cells[3 + k, 164] = new Cell(RetourValNull(ACR_SMC.ToString("#.##")));
            worksheet.Cells[3 + k, 165] = new Cell("0,00");
            worksheet.Cells[3 + k, 166] = new Cell(RetourValNull(EXT_CCBIK.ToString("#.##")));
            worksheet.Cells[3 + k, 167] = new Cell("0,00");


            return worksheet;

        }

        public string RetourValNull(string Valeur)
        {
            string ValF = string.Empty;
            if (Valeur == "")
            {
                ValF = "0,00";
            }
            else
            {
                ValF = Valeur;
            }
            return ValF;
        }

        public string NatureEmploi(string emploi)
        {
            string ValF = string.Empty;
            string emp = emploi.Substring(0, 3);
            if ((emp == "Dir") || (emp == "DIR"))
            {
                ValF="DIR";
            }
            if ((emp == "Res") || (emp == "RES"))
            {
                ValF = "SUP";
            }
            else
            {
                ValF = "EMP";
            }

            return ValF;

        }

        public string ContratInEnglish(string NatContrat)
        {
            string NatF = string.Empty;
            if (NatContrat == "CDI")
            {
                NatF = "IND";
            }
            if (NatContrat == "SIVP")
            {
                NatF = "DEF";
            }
            else
            {
                NatF = "IND";
            }
            return NatF;
        }

        public string LaValeurCivil(string CivilVal)
        {
            string Situation = string.Empty;
            if (CivilVal == "0")
            {
                Situation = "S";
            }
            if (CivilVal == "1")
            {
                Situation = "M";
            }
            if (CivilVal == "2")
            {
                Situation = "W";
            }
            if (CivilVal == "3")
            {
                Situation = "D";
            }
            return Situation;
        }

        public string DefinirGenre(string Gender)
        {
            string Ft = string.Empty;
            if (Gender == "0")
            {
                Ft = "M";
            }
            else
            {
                Ft = "F";
            }
            return Ft;
        }

        public string DateCorrecte(string CovertDate)
        {
            string taille = string.Empty;
            string YEAR = CovertDate.Substring(6, 2);
            string MONTH = CovertDate.Substring(3, 2);
            string DAY = CovertDate.Substring(0, 2);
            string DateF = "19" + YEAR + "/" + MONTH + "/" + DAY;
            return DateF;

        }

        public string DataCivil(string CivilAbr)
        {
            string Civil = string.Empty;
            if (CivilAbr == "M")
            {
                Civil = "MR";
            }
            if (CivilAbr == "Mme")
            {
                Civil = "MRS";
            }
            if (CivilAbr == "Mlle")
            {
                Civil = "MS";
            }
            return Civil;
        }

        public string DefinirModePe(string ModePe)
        {
            string ModeF = string.Empty;
            if (ModePe == "Virement")
            {
                ModeF = "BKTR";
            }
            if (ModePe == "Chèque")
            {
                ModeF = "CHQ";
            }
            if (ModePe == "Espèces")
            {
                ModeF = "CASH";
            }
            if(ModePe=="")
            {
                ModeF = "OTH";
            }
            return ModeF;

        }

        public string ExtraireHeure(string Heure, string Minute, string Seconde)
        {
            string HeureFinal = string.Empty;
            if (Heure.Length == 1)
            {
                HeureFinal = "0" + Heure;
            }
            else
            {
                HeureFinal = Heure;
            }
            if (Minute.Length == 1)
            {
                HeureFinal = HeureFinal + ":" + "0" + Minute;
            }
            else
            {
                HeureFinal = HeureFinal + ":" + Minute;
            }
            if (Seconde.Length == 1)
            {
                HeureFinal = HeureFinal + ":" + "0" + Seconde;
            }
            else
            {
                HeureFinal = HeureFinal + ":" + Seconde;
            }

            return HeureFinal;

        }

    }
}
