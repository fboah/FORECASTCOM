using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using ForecastCom.DAO;
using ForecastCom.Models;
using ForecastCom.Services;
using ForecastCom.Statistiques;
using ForecastCom.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ForecastCom
{
    public partial class MainForm : Form
    {
        private readonly CAlias daoMain = new CAlias();

        private readonly DAOForecastCom daoReport = new DAOForecastCom();

        //private readonly CRefFournisseur dFourn = new CRefFournisseur();
        //private readonly CFamCentr dFamcentr = new CFamCentr();

        private readonly FormatageForecastCom FormatMain = new FormatageForecastCom();

        private List<CAlias> ListeFilialeChoisies = new List<CAlias>();

        private List<CAlias> ListeFiliales = new List<CAlias>();

        //Listes de tous les commerciaux de toutes les bases configurées
        private List<ComClass> LCom = new List<ComClass>();

        //Listes de tous les commerciaux pour une selection donnée
        private List<ComClass> LcomSelect = new List<ComClass>();

        //Listes de tous les familles de toutes les bases configurées
        private List<ComFamille> LFam = new List<ComFamille>();

        //Listes de tous les familles centralisatrices de toutes les bases configurées
        private List<ComFamille> LFamCentr = new List<ComFamille>();

        //Listes de tous les revendeurs de toutes les bases configurées
        private List<ComRevendeur> LRev = new List<ComRevendeur>();

        //Listes de tous les familles pour une selection donnée
        private List<ComFamille> LFamSelect = new List<ComFamille>();

        //Listes de tous les familles pour une selection donnée
        private List<ComFamille> LFamCentralSelect = new List<ComFamille>();


        //Liste de tous les revendeurs pour une selection donnée

        private List<ComRevendeur> LRevSelect = new List<ComRevendeur>();

        //chaine des values des type de documents
        private string ListDocument;
        
        //Chaine pour Nom Commerciale DE

        private string ListNomLECommerciauxDE;

        //Chaine pour Nom Commerciale A

        private string ListNomLECommerciauxA;

        //Chaine pour NOM MulticheckCommerciaux

        private string ListNomMultiCommerciaux;

        //Chaine pour PRENOM MulticheckCommerciaux

        private string ListPrenomMultiCommerciaux;

        //chaine pour Le code famille
        private string ListNomMultiFamille;

        //chaine pour Le code Revendeur
        private string ListNomMultiRevendeur;

        //chaine pour Multi Famille Central
        private string ListNomMultiFamilleCentralisatrice;

        //Chaine pour Nom famille DE

        private string ListNomLEFamilleDE;

        //Chaine pour Nom famille Central DE

        private string ListNomLEFamilleCentrDE;

        //Chaine pour Nom Revendeur DE

        private string ListNomLERevendeurDE;
        
        //Chaine pour Nom famille A

        private string ListNomLEFamilleA;

        //Chaine pour Nom famille Centr A

        private string ListNomLEFamilleCentrA;

        //Chaine pour Nom Revendeur A

        private string ListNomLERevendeurA;

        //Listes pour statistiques Com
        public List<FCommercial> LcomStat = new List<FCommercial>();
        public List<FCommercial> LcomStatTypeDoc = new List<FCommercial>();
        public List<FCommercial> LcomStatNbreDoc = new List<FCommercial>();
        public List<FCommercial> LcomStatNbreTypeDocCom = new List<FCommercial>();
        public List<FCommercial> LcomTopFamMtant = new List<FCommercial>();
        public List<FCommercial> LcomTopFamQtite = new List<FCommercial>();

        public List<FCommercial> LcomTopFamCentralisatriceMtant = new List<FCommercial>();
        public List<FCommercial> LcomTopFamCentralisatriceQtite = new List<FCommercial>();

        public List<FCommercial> LcomRanking = new List<FCommercial>();

        //Listes pour statistiques Revendeur
        public List<ComRevendeur> LcomStatTypeDocRevendeur = new List<ComRevendeur>();

        public List<ComRevendeur> LcomStatComRevendeur = new List<ComRevendeur>();

        public List<ComRevendeur> LcomStatTOPRevendeur = new List<ComRevendeur>();

        public List<ComRevendeur> LcomStatBADRevendeur = new List<ComRevendeur>();



        //tester si on a que des factures ou non 
        public bool IsFactOnly = false;

        

        public MainForm()
        {
            InitializeComponent();
        }

        private void connexionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

               if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                var FenConnexion = new Connexion();

                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                FenConnexion.ShowDialog();

                ReloadApplication();

                //Rafraichir la grid des filiales

                ListeFiliales = daoMain.GetAliasConnexions();

                //Remplir les filiales à selectionner(on exclu alors abidjan)

                if (ListeFiliales != null)
                {
                    ////ListeAutreFiliales.Clear();
                    //var ListeAutreFiliales = ListeFiliales.Where(c => c.IsAbidjan != true).ToList();

                    RemplirFilialesCheckedControl(ChkListBoxControlFiliale, ListeFiliales);

                }

                //Remplir les filiales à selectionner(on exclu alors abidjan)

                
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->connexionsToolStripMenuItem_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
           
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            string chaineBABI = string.Empty;
            try
            {
               
          
                dateEditDateDeb.EditValue = DateTime.Now;

                dateEditDateFin.EditValue = DateTime.Now;

                //Remplir la liste des filiales
                ListeFiliales = daoMain.GetAlias();

                if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    chaineBABI = ChooseBABI.mAliasName;

                }   

                if(ListeFiliales!=null)
                {
                    if(ListeFiliales.Count>0)
                    {
                        //remplir Combo Commerciaux=======================================
                        LCom = daoReport.LoadComAll(chaineBABI, ListeFiliales);

                        //FillComboCommercialDE(LECommerciauxDE,LCom);
                        //FillComboCommercialDE(LECommerciauxA, LcomSelect);
                        LECommerciauxDE.Properties.DataSource = LCom;
                        LECommerciauxA.Properties.DataSource = LCom;

                        chkCmbMultCom.Properties.DataSource = LCom;
                        FillMultiCheckComboCommercial(chkCmbMultCom, LCom);

                        LECommerciauxDE.Properties.DisplayMember = "mNomCommercial";
                        LECommerciauxA.Properties.DisplayMember = "mNomCommercial";

                        //Choisir les premières valeurs
                        LECommerciauxDE.ItemIndex = 0;
                        LECommerciauxA.ItemIndex = 0;

                        //Remplir Combo Famille Centralisatrice==============================================
                        LFamCentr = daoReport.LoadFamilleCentralAll(chaineBABI, ListeFiliales);

                        LEFamilleCentralDE.Properties.DataSource = LFamCentr;
                        LEFamilleCentralA.Properties.DataSource = LFamCentr;

                        chkCmbMultFamCentral.Properties.DataSource = LFamCentr;
                        FillMultiCheckComboFamilleCentral(chkCmbMultFamCentral, LFamCentr);

                        LEFamilleCentralDE.Properties.DisplayMember = "mFa_CodeFamille";
                        LEFamilleCentralA.Properties.DisplayMember = "mFa_CodeFamille";

                        //Choisir les premières valeurs
                        LEFamilleCentralDE.ItemIndex = 0;
                        LEFamilleCentralA.ItemIndex = 0;

                        //Remplir Combo Famille==============================================
                        LFam = daoReport.LoadFamilleAll(chaineBABI, ListeFiliales);

                        LEFamilleDE.Properties.DataSource = LFam;
                        LEFamilleA.Properties.DataSource = LFam;

                        chkCmbMultFam.Properties.DataSource = LFam;
                        FillMultiCheckComboFamille(chkCmbMultFam, LFam);

                        LEFamilleDE.Properties.DisplayMember = "mFa_CodeFamille";
                        LEFamilleA.Properties.DisplayMember = "mFa_CodeFamille";

                        //Choisir les premières valeurs
                        LEFamilleDE.ItemIndex = 0;
                        LEFamilleA.ItemIndex = 0;

                        //Remplir Combo Revendeur==============================================
                        LRev = daoReport.LoadRevendeurAll(chaineBABI, ListeFiliales);

                        LERevendeurDE.Properties.DataSource = LRev;
                        LERevendeurA.Properties.DataSource = LRev;

                        chkCmbMultFam.Properties.DataSource = LRev;
                        FillMultiCheckComboRevendeur(chkCmbMultRev, LRev);

                        LERevendeurDE.Properties.DisplayMember = "mCT_Num";
                        LERevendeurA.Properties.DisplayMember = "mCT_Num";

                        //Choisir les premières valeurs
                        LERevendeurDE.ItemIndex = 0;
                        LERevendeurA.ItemIndex = 0;

                        //Reinitialiser montant indicateurs
                        ReInitialiserIndicateursMontant();
                        ////Initialiser montant PIPE
                        //lblControlTotalHT.Text = "0";
                        //lblControlTotalHT2.Text = "0";
                        //lblControlTotalHT3.Text = "0";

                        ////Initaliser facturation PIPE
                        //lblControlRealiseHT.Text = "0";

                        ////initialiser Accurency

                        //lblAccurency.Text = "0";
                    }
                }
             
                //Remplir les filiales à selectionner

                if (ListeFiliales != null)
                {
                  //  var ListeAutreFiliales = ListeFiliales.Where(c => c.IsAbidjan != true).ToList();

                    RemplirFilialesCheckedControl(ChkListBoxControlFiliale, ListeFiliales);

                }


                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->MainForm_Load -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }


        public void ReInitialiserIndicateursMontant()
        {
            try
            {
                lblControlTotalHT.Text = "0";
                lblControlTotalHT2.Text = "0";
                lblControlTotalHT3.Text = "0";

                //Initialiser fact et acc

                lblAccurency.Text = "0";
                lblControlRealiseHT.Text = "0";
                lblAccurency2.Text = "0";
                lblControlRealiseHT2.Text = "0";
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->ReInitialiserIndicateursMontant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }


        //Recharger Appli

        public void ReloadApplication()
        {
            string chaineBABI = string.Empty;
            try
            {
                if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();

                dateEditDateDeb.EditValue = DateTime.Now;

                dateEditDateFin.EditValue = DateTime.Now;

                //Remplir la liste des filiales
                ListeFiliales = daoMain.GetAlias();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    chaineBABI = ChooseBABI.mAliasName;

                }

                //remplir Combo Commerciaux=======================================
                LCom = daoReport.LoadComAll(chaineBABI, ListeFiliales);

                //FillComboCommercialDE(LECommerciauxDE,LCom);
                //FillComboCommercialDE(LECommerciauxA, LcomSelect);
                LECommerciauxDE.Properties.DataSource = LCom;
                LECommerciauxA.Properties.DataSource = LCom;

                chkCmbMultCom.Properties.DataSource = LCom;
                FillMultiCheckComboCommercial(chkCmbMultCom, LCom);

                LECommerciauxDE.Properties.DisplayMember = "mNomCommercial";
                LECommerciauxA.Properties.DisplayMember = "mNomCommercial";

                //Choisir les premières valeurs
                LECommerciauxDE.ItemIndex = 0;
                LECommerciauxA.ItemIndex = 0;

                //Remplir Combo Famille Centralisatrice==============================================
                LFamCentr = daoReport.LoadFamilleCentralAll(chaineBABI, ListeFiliales);

                LEFamilleCentralDE.Properties.DataSource = LFamCentr;
                LEFamilleCentralA.Properties.DataSource = LFamCentr;

                chkCmbMultFamCentral.Properties.DataSource = LFamCentr;
                FillMultiCheckComboFamilleCentral(chkCmbMultFamCentral, LFamCentr);

                LEFamilleCentralDE.Properties.DisplayMember = "mFa_CodeFamille";
                LEFamilleCentralA.Properties.DisplayMember = "mFa_CodeFamille";

                //Choisir les premières valeurs
                LEFamilleCentralDE.ItemIndex = 0;
                LEFamilleCentralA.ItemIndex = 0;

                //Remplir Combo Famille==============================================
                LFam = daoReport.LoadFamilleAll(chaineBABI, ListeFiliales);

                LEFamilleDE.Properties.DataSource = LFam;
                LEFamilleA.Properties.DataSource = LFam;

                chkCmbMultFam.Properties.DataSource = LFam;
                FillMultiCheckComboFamille(chkCmbMultFam, LFam);

                LEFamilleDE.Properties.DisplayMember = "mFa_CodeFamille";
                LEFamilleA.Properties.DisplayMember = "mFa_CodeFamille";

                //Choisir les premières valeurs
                LEFamilleDE.ItemIndex = 0;
                LEFamilleA.ItemIndex = 0;

                //Remplir Combo Revendeur==============================================
                LRev = daoReport.LoadRevendeurAll(chaineBABI, ListeFiliales);

                LERevendeurDE.Properties.DataSource = LRev;
                LERevendeurA.Properties.DataSource = LRev;

                chkCmbMultRev.Properties.DataSource = LRev;
                FillMultiCheckComboRevendeur(chkCmbMultRev, LRev);

                LERevendeurDE.Properties.DisplayMember = "mCT_Num";
                LERevendeurA.Properties.DisplayMember = "mCT_Num";

                //Choisir les premières valeurs
                LERevendeurDE.ItemIndex = 0;
                LERevendeurA.ItemIndex = 0;

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                ////Initialiser montant PIPE
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";


                //Remplir les filiales à selectionner

                if (ListeFiliales != null)
                {
                    //  var ListeAutreFiliales = ListeFiliales.Where(c => c.IsAbidjan != true).ToList();

                    RemplirFilialesCheckedControl(ChkListBoxControlFiliale, ListeFiliales);

                }


                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->ReloadApplication -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }



        public void FillComboCommercial(LookUpEdit cmb,List<ComClass> Lste)
        {
            try
            {
                if (Lste != null && Lste.Count > 0)
                {
                    var MySelectBases = new DataTable();

                    MySelectBases.Columns.Add("mNomCommercial");

                    MySelectBases.Columns.Add("mPrenomCommercial");

                    MySelectBases.Columns.Add("mPays");


                    foreach (var item in Lste)
                    {
                        MySelectBases.Rows.Add(item.mNomCommercial, item.mPrenomCommercial, item.mPays);

                        cmb.Properties.DataSource = MySelectBases;

                        cmb.Properties.ValueMember = "mPrenomCommercial";

                        cmb.Properties.DisplayMember = "mNomCommercial";

                        cmb.EditValue = item.mPrenomCommercial;


                    }
                    
                }
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm ->FillComboCommercial -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
          
        }

        public void FillMultiCheckComboCommercial(CheckedComboBoxEdit cmb, List<ComClass> Lste)
        {
            try
            {
                if (Lste != null && Lste.Count > 0)
                {
                    var MySelectBases = new DataTable();

                    MySelectBases.Columns.Add("mNomCommercial");

                    MySelectBases.Columns.Add("mPrenomCommercial");

                    MySelectBases.Columns.Add("mNomPrenomCommercial");

                    MySelectBases.Columns.Add("mPays");


                    foreach (var item in Lste)
                    {
                        if(item.mNomCommercial!=string.Empty && item.mPrenomCommercial!=string.Empty)
                        {
                            MySelectBases.Rows.Add(item.mNomCommercial, item.mPrenomCommercial, item.mNomCommercial + " " + item.mPrenomCommercial, item.mPays);
                            
                        }

                        if (item.mNomCommercial == string.Empty && item.mPrenomCommercial != string.Empty)
                        {
                            MySelectBases.Rows.Add(item.mNomCommercial, item.mPrenomCommercial,  item.mPrenomCommercial, item.mPays);

                        }

                        if (item.mNomCommercial != string.Empty && item.mPrenomCommercial == string.Empty)
                        {
                            MySelectBases.Rows.Add(item.mNomCommercial, item.mNomCommercial, item.mNomCommercial , item.mPays);

                        }

                        cmb.Properties.DataSource = MySelectBases;

                        cmb.Properties.ValueMember = "mPrenomCommercial";

                        cmb.Properties.DisplayMember = "mNomPrenomCommercial";

                       

                        // cmb.EditValue = item.mPrenomCommercial;


                    }



                }
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm ->FillMultiCheckComboCommercial -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }

        //decocher toutes les valeurs
         public void DecocherMultiCom(CheckedComboBoxEdit cmb)
        {
            try
            {
                foreach(CheckedListBoxItem item in cmb.Properties.Items)
                {
                    if(item.CheckState==CheckState.Checked)
                    {
                        item.CheckState = CheckState.Unchecked;
                    }

                }
               
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm ->DecocherMultiCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }


        public void FillComboFamille(LookUpEdit cmb, List<ComFamille> Lste)
        {
            try
            {
                if (Lste != null && Lste.Count > 0)
                {
                    var MySelectBases = new DataTable();

                    MySelectBases.Columns.Add("mNomCommercial");
                    MySelectBases.Columns.Add("mPays");


                    foreach (var item in Lste)
                    {
                        MySelectBases.Rows.Add(item.mFa_CodeFamille, item.mPays);

                        cmb.Properties.DataSource = MySelectBases;

                        cmb.Properties.ValueMember = "mFa_CodeFamille";

                        cmb.Properties.DisplayMember = "mFa_CodeFamille";
                        
                    }
                    

                }
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm ->FillComboFamille -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
           
        }


        public void FillMultiCheckComboFamille(CheckedComboBoxEdit cmb, List<ComFamille> Lste)
        {
            try
            {
                if (Lste != null && Lste.Count > 0)
                {
                    var MySelectBases = new DataTable();

                    MySelectBases.Columns.Add("mFa_CodeFamille");

                    foreach (var item in Lste)
                    {
                        MySelectBases.Rows.Add(item.mFa_CodeFamille);

                        cmb.Properties.DataSource = MySelectBases;

                        cmb.Properties.ValueMember = "mFa_CodeFamille";

                        cmb.Properties.DisplayMember = "mFa_CodeFamille";

                    }



                }
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm ->FillMultiCheckComboFamille -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
           
        }


        public void FillMultiCheckComboFamilleCentral(CheckedComboBoxEdit cmb, List<ComFamille> Lste)
        {
            try
            {
                if (Lste != null && Lste.Count > 0)
                {
                    var MySelectBases = new DataTable();

                    MySelectBases.Columns.Add("mFa_CodeFamille");

                    foreach (var item in Lste)
                    {
                        MySelectBases.Rows.Add(item.mFa_CodeFamille);

                        cmb.Properties.DataSource = MySelectBases;

                        cmb.Properties.ValueMember = "mFa_CodeFamille";

                        cmb.Properties.DisplayMember = "mFa_CodeFamille";

                    }



                }
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm ->FillMultiCheckComboFamilleCentral -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }


        public void FillMultiCheckComboRevendeur(CheckedComboBoxEdit cmb, List<ComRevendeur> Lste)
        {
            try
            {
                if (Lste != null && Lste.Count > 0)
                {
                    var MySelectBases = new DataTable();

                    MySelectBases.Columns.Add("mCT_Num");

                    foreach (var item in Lste)
                    {
                        MySelectBases.Rows.Add(item.mCT_Num);

                        cmb.Properties.DataSource = MySelectBases;

                        cmb.Properties.ValueMember = "mCT_Num";

                        cmb.Properties.DisplayMember = "mCT_Num";

                    }



                }
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm ->FillMultiCheckComboRevendeur -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }


        //Remplir CheckedListcontrol

        public void RemplirFilialesCheckedControl(DevExpress.XtraEditors.CheckedListBoxControl ChkListBox, List<CAlias> LAlias)
        {
            try
            {
                ChkListBox.Items.Clear();

                if (LAlias.Count > 0)
                {
                    foreach (var item in LAlias)
                    {
                        var str = new CheckedListBoxItem();

                        str.Value = item.mId;
                        str.Description = item.mVille;

                        if (str.Description != string.Empty)
                        {
                            //Checker Abidjan seulement;On choisira les filiales en cas de necessité
                            if(item.mId==1)
                            {
                                str.CheckState = CheckState.Checked;

                                //remplir combo des commerciaux pour Abidjan===============
                                
                                LcomSelect = LCom.Where(c => c.mIdPays == 1).ToList();
                                
                                LECommerciauxDE.Properties.DataSource = LcomSelect;
                                 LECommerciauxA.Properties.DataSource = LcomSelect;
                                //FillComboCommercialDE(LECommerciauxDE, LcomSelect);
                                //FillComboCommercialDE(LECommerciauxA, LcomSelect);

                                //Choisir les premières valeurs
                                LECommerciauxDE.ItemIndex = 0;
                                LECommerciauxA.ItemIndex = 0;
                          
                                chkCmbMultCom.Properties.DataSource = LcomSelect;
                              //  DecocherMultiCom(chkCmbMultCom);
                              FillMultiCheckComboCommercial(chkCmbMultCom, LcomSelect);

                                //remplir Combo famille central pour babi=========================
                               
                                LFamCentralSelect = LFamCentr.Where(c => c.mIdPays == 1).ToList();
                                LEFamilleCentralDE.Properties.DataSource = LFamCentralSelect;
                                LEFamilleCentralA.Properties.DataSource = LFamCentralSelect;

                                //Choisir les premières valeurs
                                LEFamilleCentralDE.ItemIndex = 0;
                                LEFamilleCentralA.ItemIndex = 0;
                                chkCmbMultFamCentral.Properties.DataSource = LFamCentralSelect;
                                FillMultiCheckComboFamille(chkCmbMultFamCentral, LFamCentralSelect);

                                //remplir Combo famille pour babi=========================

                                LFamSelect =LFam.Where(c => c.mIdPays == 1).ToList();
                                LEFamilleDE.Properties.DataSource = LFamSelect;
                                LEFamilleA.Properties.DataSource = LFamSelect;

                                //Choisir les premières valeurs
                                LEFamilleDE.ItemIndex = 0;
                                LEFamilleA.ItemIndex = 0;
                                chkCmbMultFam.Properties.DataSource = LFamSelect;
                                FillMultiCheckComboFamille(chkCmbMultFam, LFamSelect);

                                //remplir Combo Revendeur pour babi=========================
                                
                                LRevSelect = LRev.Where(c => c.mIdPays == 1).ToList();
                                LERevendeurDE.Properties.DataSource = LRevSelect;
                                LERevendeurA.Properties.DataSource = LRevSelect;

                                //Choisir les premières valeurs
                                LERevendeurDE.ItemIndex = 0;
                                LERevendeurA.ItemIndex = 0;
                                chkCmbMultRev.Properties.DataSource = LRevSelect;
                                FillMultiCheckComboRevendeur(chkCmbMultRev, LRevSelect);


                            }
                            else
                            {
                                str.CheckState = CheckState.Unchecked;
                            }
                         
                            ChkListBox.Items.Add(str);
                        }

                    }

                }
            }
            catch (Exception ex)
            {
               if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm -> RemplirFilialesCheckedControl -> TypeErreur: " + ex.Message; 
                CAlias.Log(msg);

            }

        }

        
        private void chkListBoxControlFiliale_ItemCheck(object sender, DevExpress.XtraEditors.Controls.ItemCheckEventArgs e)
        {
            try
            {
                //Liste utilisée pour la génération du Report
                ListeFilialeChoisies.Clear();

                List<String> Lidpays = new List<string>();
                List<ComClass> LComboCom = new List<ComClass>();
                List<ComFamille> LComboFam = new List<ComFamille>();
                List<ComFamille> LComboFamCentr = new List<ComFamille>();

                List<ComRevendeur> LComboRev = new List<ComRevendeur>();

                //Ajouter Abidjan
                //var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                //ListeFilialeChoisies.Add(ChooseBABI);

                foreach (CheckedListBoxItem item in ChkListBoxControlFiliale.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        var Choose = ListeFiliales.FirstOrDefault(c => c.mId == Convert.ToInt32(item.Value.ToString()));

                        ListeFilialeChoisies.Add(Choose);
                        Lidpays.Add(item.Value.ToString());
                    }

                }

                //Remplir le combo commerciaux d'apres les pays choisis=============================
                LComboCom = daoReport.GetListByListIdCom(Lidpays, LCom);

                //FillComboCommercialDE(LECommerciauxDE, LComboCom);
                //FillComboCommercialDE(LECommerciauxA, LComboCom);

                LECommerciauxDE.Properties.DataSource = LComboCom;
                LECommerciauxA.Properties.DataSource = LComboCom;

                chkCmbMultCom.Properties.DataSource = LComboCom;
                FillMultiCheckComboCommercial(chkCmbMultCom, LComboCom);
                
                //Actualiser le texte du Multi combo Com

                ReloadMulticomboCom(chkCmbMultCom, LComboCom);
                
                //Choisir les premières valeurs
                LECommerciauxDE.ItemIndex = 0;
                LECommerciauxA.ItemIndex = 0;

                //Choisir par défaut les premiers text au cas où on y toucherai pas aprés(DE_A)

                ListNomLECommerciauxA = LECommerciauxA.Text;
                ListNomLECommerciauxDE = LECommerciauxDE.Text;

                ListNomLEFamilleA = LEFamilleA.Text;
                ListNomLEFamilleDE = LEFamilleDE.Text;

                //Remplir Combo Famille Centralisatrice==============================================
                
                
                LComboFamCentr = daoReport.GetListByListIdFam(Lidpays, LFamCentr);

                LEFamilleCentralDE.Properties.DataSource = LComboFamCentr;
                LEFamilleCentralA.Properties.DataSource = LComboFamCentr;

                chkCmbMultFamCentral.Properties.DataSource = LComboFamCentr;
                FillMultiCheckComboFamilleCentral(chkCmbMultFamCentral, LComboFamCentr);

                LEFamilleCentralDE.Properties.DisplayMember = "mFa_CodeFamille";
                LEFamilleCentralA.Properties.DisplayMember = "mFa_CodeFamille";

                //Choisir les premières valeurs
                LEFamilleCentralDE.ItemIndex = 0;
                LEFamilleCentralA.ItemIndex = 0;

                //Remplir Combo famille===========================================================

                LComboFam = daoReport.GetListByListIdFam(Lidpays, LFam);

                LEFamilleDE.Properties.DataSource = LComboFam;
                LEFamilleA.Properties.DataSource = LComboFam;

                chkCmbMultFam.Properties.DataSource = LComboFam;
                FillMultiCheckComboFamille(chkCmbMultFam, LComboFam);

                //Actualiser le texte du Multi combo Com

                 ReloadMulticomboFamille(chkCmbMultFam, LComboFam);

                //Choisir les premières valeurs
                LEFamilleDE.ItemIndex = 0;
                LEFamilleA.ItemIndex = 0;

                //Remplir Combo Revendeur===========================================================

                LComboRev = daoReport.GetListByListIdRev(Lidpays, LRev);

                LERevendeurDE.Properties.DataSource = LComboRev;
                LERevendeurA.Properties.DataSource = LComboRev;

                chkCmbMultRev.Properties.DataSource = LComboRev;
                FillMultiCheckComboRevendeur(chkCmbMultRev, LComboRev);

                //Actualiser le texte du Multi combo rev

                ReloadMulticomboRevendeur(chkCmbMultRev, LComboRev);

                //Choisir les premières valeurs
                LERevendeurDE.ItemIndex = 0;
                LERevendeurA.ItemIndex = 0;


                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";

                CancelConfigFamille();

                CleanGrid();
                timer1.Enabled = true;

            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->chkListBoxControlFiliale_ItemCheck -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }


        private void CleanGrid()
        {
            try
            {
                //Nettoyer la grid=================================================
                List<FCommercial> Lempty = new List<FCommercial>();

                gridControlForecast.DataSource = Lempty;
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->CleanGrid -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
       
        }

        //Actualisercom 
        private void ReloadMulticomboCom(CheckedComboBoxEdit cmb,List<ComClass> LC)
        {

            try
            {
                ListNomMultiCommerciaux = string.Empty;

                ListPrenomMultiCommerciaux = string.Empty;

                foreach (CheckedListBoxItem item in cmb.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        var testNom = LC.FirstOrDefault(n => n.mPrenomCommercial.Equals(item.Value.ToString()));

                        if(testNom!=null)
                        {
                            string tmp = " " + item.Value.ToString();

                            string NomCom = item.Description.Replace(tmp, "");

                            var testPrenom = LC.FirstOrDefault(n=>n.mNomCommercial.Equals(NomCom));
                            
                            if(testPrenom!=null)
                            {
                                //le value c'est le prenom ,
                                ListPrenomMultiCommerciaux += item.Value.ToString() + ",";

                                //on peut donc retirer le nom par déduction d'avec le texte(description =Nom +" "+Prenom)

                                ListNomMultiCommerciaux += NomCom + ",";


                            }


                        }


                    }

                }

                if (ListPrenomMultiCommerciaux.Length > 0) ListPrenomMultiCommerciaux = ListPrenomMultiCommerciaux.Substring(0, ListPrenomMultiCommerciaux.Length - 1);
                if (ListNomMultiCommerciaux.Length > 0) ListNomMultiCommerciaux = ListNomMultiCommerciaux.Substring(0, ListNomMultiCommerciaux.Length - 1);

                cmb.RefreshEditValue();
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
              
                var msg = "MainForm -> ReloadMulticombo -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        

    }

        //Actualiser Famille
        private void ReloadMulticomboFamille(CheckedComboBoxEdit cmb, List<ComFamille> LC)
        {

            try
            {
                ListNomMultiFamille = string.Empty;
                
                foreach (CheckedListBoxItem item in cmb.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {

                        ListNomMultiFamille += item.Value.ToString() + ",";
                           
                    }

                }

                if (ListNomMultiFamille.Length > 0) ListNomMultiFamille = ListNomMultiFamille.Substring(0, ListNomMultiFamille.Length - 1);
             
                cmb.RefreshEditValue();
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
               
                var msg = "MainForm -> ReloadMulticomboFamille -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }


        }



        //Actualiser Revendeur
        private void ReloadMulticomboRevendeur(CheckedComboBoxEdit cmb, List<ComRevendeur> LC)
        {

            try
            {
                ListNomMultiRevendeur = string.Empty;

                foreach (CheckedListBoxItem item in cmb.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {

                        ListNomMultiRevendeur += item.Value.ToString() + ",";

                    }

                }

                if (ListNomMultiRevendeur.Length > 0) ListNomMultiRevendeur = ListNomMultiRevendeur.Substring(0, ListNomMultiRevendeur.Length - 1);

                cmb.RefreshEditValue();
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> ReloadMulticomboRevendeur -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }


        }



        private void sBtnVisualiserArticle_Click(object sender, EventArgs e)
        {
            try
            {

                if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
                {
                    MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                List<FCommercial> ListCom = new List<FCommercial>();

                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if(ChooseBABI!=null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if(ListeFilialeChoisies.Count==0)
                    {
                        //Verifier qu'on a rien choisi

                        if(ChkListBoxControlFiliale.CheckedItemsCount>0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c=>c.IsAbidjan==true).ToList() ;

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                    }

                    if(ListeFilialeChoisies.Count>0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                      //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if(ListNomLEFamilleDE==string.Empty || ListNomLEFamilleDE==null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }
                        
                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }
                        
                        
                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //cas hors facturation Pipe

                        if(chkFacture.Checked)
                        {

                            ListCom = daoReport.GetForecastBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);


                            if (ListCom != null && ListCom.Count > 0)
                            {
                                var SommeSolde = ListCom.FirstOrDefault();

                                //cas Hors Pipe

                                if (!chkSaisieDevis.Checked && !chkAccepteDevis.Checked)
                                {
                                    lblControlRealiseHT.Text= SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    lblControlRealiseHT2.Text= SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");

                                    //lblControlTotalHT.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    //lblControlTotalHT2.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    //lblControlTotalHT3.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                }
                                else
                                {
                                    lblControlTotalHT.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                    lblControlTotalHT2.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                    lblControlTotalHT3.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                }


                            }
                            else
                            {
                                lblControlRealiseHT.Text = "0";
                                lblControlRealiseHT2.Text = "0";
                                //lblControlTotalHT.Text = "0";
                                //lblControlTotalHT2.Text = "0";
                                //lblControlTotalHT3.Text = "0";
                            }



                        }
                        else
                        {
                           //ListCom = daoReport.GetForecast(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);
                            ListCom = daoReport.GetForecastMARGE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);
                            
                            if (ListCom != null && ListCom.Count > 0)
                            {
                                var SommeSolde = ListCom.FirstOrDefault();
                               
                                //cas Hors Pipe

                                if (!chkSaisieDevis.Checked && !chkAccepteDevis.Checked)
                                {
                                    lblControlTotalHT.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    lblControlTotalHT2.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    lblControlTotalHT3.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                }
                                else
                                {
                                    lblControlTotalHT.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                    lblControlTotalHT2.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                    lblControlTotalHT3.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                }

                                //  Montant Marge
                                //gg
                                lblAccurency.Text=SommeSolde.mMontantTotalMarge.ToString("n0");
                                lblAccurency2.Text=SommeSolde.mMontantTotalMarge.ToString("n0");

                            }
                            else
                            {

                                lblControlTotalHT.Text = "0";
                                lblControlTotalHT2.Text = "0";
                                lblControlTotalHT3.Text = "0";
                                lblAccurency.Text = "0";
                                lblAccurency2.Text = "0";
                            }
                        }

                        ////Accurency
                        ////  taux transformation (fact/devis)*100

                        //double MtFact= Double.Parse(lblControlRealiseHT.Text);
                        //double MtDev = Double.Parse(lblControlTotalHT.Text);
                        //double acc = 0;

                        //if(MtDev != 0)
                        //{
                        //    acc = (MtFact / MtDev) * 100;

                        //    lblAccurency.Text = acc.ToString("n2");
                        //    lblAccurency2.Text = acc.ToString("n2");
                        //}
                        //else
                        //{
                        //    lblAccurency.Text = "0";
                        //    lblAccurency2.Text = "0";
                        //}
                        

                      
                          if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                        gridControlForecast.DataSource = ListCom;

                        //Cas BUNDLE ,Masquer la colonne MARGE
                        if(chkFacture.Checked)
                        {
                            gridView1.Columns["mMargeBrute"].Visible = false;
                            gridView1.Columns["mFamilleCentral"].Visible = false;
                            gridView1.Columns["mFamille"].Visible = false;
                            
                        }
                        else
                        {
                            gridView1.Columns["mMargeBrute"].Visible = true;
                            gridView1.Columns["mFamilleCentral"].Visible = true;
                            gridView1.Columns["mFamille"].Visible = true;
                        }


                        timer1.Enabled = false;

                        sBtnVisualiserArticle.ForeColor = Color.Black;
                    }

                 
                }
               
               
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->sBtnVisualiserArticle_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

        

        public string GetDateFormat(DateTime dat)
        {
            string ret = string.Empty;
            try
            {
                int day = 0;
                int mois = 0;
                int an = 0;

                string daydat = string.Empty;
                string moisdat = string.Empty;
                string andat = string.Empty;

                day = dat.Day;
                mois = dat.Month;
                an = dat.Year;

                if (day < 10)
                {
                    daydat = "0" + day.ToString() ;
                }
                else
                {
                    daydat = day.ToString() ;
                }

                if (mois < 10)
                {
                    moisdat = "0" + mois.ToString();
                }
                else
                {
                    moisdat = mois.ToString();
                }

                ret = daydat + "_" + moisdat + "_" + an.ToString();

                return ret;
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetDateFormat -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return ret;
            }


        }
        
        private void sBtnGenererExcel_Click(object sender, EventArgs e)
        {
            string cheminFichier = string.Empty;

            // string RepDestXls = string.Empty;

            int processIdAvant = 0;

            int PidExcelproc = 0;

            bool IsOK = false;

            string chaineBABI = string.Empty;

            List<FCommercial> ListFC = new List<FCommercial>();

            try
            {
                if(dateEditDateDeb.Text==string.Empty || dateEditDateFin.Text==string.Empty)
                {
                    MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                List<FCommercial> ListCom = new List<FCommercial>();

                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                     chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant la génération!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {

                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }


                        //cas hors facturation Pipe

                        if (chkFacture.Checked)
                        {
                            ListCom = daoReport.GetForecastBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);
                            
                            if (ListCom != null && ListCom.Count > 0)
                            {
                                var SommeSolde = ListCom.FirstOrDefault();

                                //cas Hors Pipe

                                if (!chkSaisieDevis.Checked && !chkAccepteDevis.Checked)
                                {
                                    lblControlRealiseHT.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    lblControlRealiseHT2.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");

                                    //lblControlTotalHT.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    //lblControlTotalHT2.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    //lblControlTotalHT3.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                }
                                else
                                {
                                    lblControlTotalHT.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                    lblControlTotalHT2.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                    lblControlTotalHT3.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                }


                            }
                            else
                            {
                                lblControlRealiseHT.Text = "0";
                                lblControlRealiseHT2.Text = "0";
                                //lblControlTotalHT.Text = "0";
                                //lblControlTotalHT2.Text = "0";
                                //lblControlTotalHT3.Text = "0";
                            }
                            
                        }
                        else
                        {
                           //ListCom = daoReport.GetForecast(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);
                            ListCom = daoReport.GetForecastMARGE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                            if (ListCom != null && ListCom.Count > 0)
                            {
                                var SommeSolde = ListCom.FirstOrDefault();

                                //cas Hors Pipe

                                if (!chkSaisieDevis.Checked && !chkAccepteDevis.Checked)
                                {
                                    lblControlTotalHT.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    lblControlTotalHT2.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                    lblControlTotalHT3.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                                }
                                else
                                {
                                    lblControlTotalHT.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                    lblControlTotalHT2.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                    lblControlTotalHT3.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                                }


                            }
                            else
                            {

                                lblControlTotalHT.Text = "0";
                                lblControlTotalHT2.Text = "0";
                                lblControlTotalHT3.Text = "0";
                            }
                        }


                        // //ListCom = daoReport.GetForecast(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur);

                        //ListCom = daoReport.GetForecast(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);


                        //if (ListCom != null && ListCom.Count > 0)
                        //{
                        //    var SommeSolde = ListCom.FirstOrDefault();

                        //    //cas Hors Pipe

                        //    if (!chkSaisieDevis.Checked && !chkAccepteDevis.Checked)
                        //    {
                        //        lblControlTotalHT.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                        //        lblControlTotalHT2.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                        //        lblControlTotalHT3.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                        //    }
                        //    else
                        //    {
                        //        lblControlTotalHT.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                        //        lblControlTotalHT2.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                        //        lblControlTotalHT3.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                        //    }


                        //}
                        //else
                        //{

                        //    lblControlTotalHT.Text = "0";
                        //    lblControlTotalHT2.Text = "0";
                        //    lblControlTotalHT3.Text = "0";
                        //}

                        if (ListCom.Count==0)
                        {
                            //Rien à afficher
                            MessageBox.Show("Aucune données à exporter en Excel!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        //recupérer liste des bases choisies
                        
                        string NomSheet = "imputs_" + GetDateFormat(DateTime.Now);

                        string RepDestXlsSIT = string.Empty;

                        //On a des eléments ,on peut imprimer dans excel

                        if (sFDExcel.ShowDialog() == DialogResult.OK)
                        {
                            cheminFichier = sFDExcel.FileName;

                            if (cheminFichier.Length >= 203)
                            {
                                //chemin trop long

                                MessageBox.Show("Le chemin spécifié est trop long !Veuillez choisir un autre dossier!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }

                            ////Obtenir la somme des pid Excel ouvert avant qu'on initialise notre objet
                            //processIdAvant = GetPidExcelAvant();

                            var strSIT = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ressources/" + "input.xlsx");
                            
                            if (File.Exists(strSIT))//verifier que le template existe
                            {

                                RepDestXlsSIT = cheminFichier ;

                                bool isCopyOk = true;

                                if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();

                                //Copier le template
                                File.Copy(strSIT, RepDestXlsSIT, isCopyOk);

                                #region Chargement Listes
                                splashScreenManager2.SetWaitFormDescription("Veuillez patienter! Traitement en cours...40%");
                                //  ListeCPOS = daoReport.GetPOSEXCEL(ChaineConnexionBabi, datedeb, datefin, RefFourn, FamCentr, ListeFilialeChoisies);

                                if (chkFacture.Checked)
                                {
                                  //  isImputsOK = FormatMain.FormatForecastBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, xlsApp, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                                    ListFC = daoReport.GetForecastBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);
                                    
                                }
                                else
                                {
                                    ListFC = daoReport.GetForecast(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);
                                    
                                }
                                
                                //Initialiser les objets Excel

                                Microsoft.Office.Interop.Excel.Application ExcelApplication = new Microsoft.Office.Interop.Excel.Application();

                                var xclasseur = ExcelApplication.Workbooks.Open(RepDestXlsSIT);

                                Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

                                #region EcrireFichier
                                
                                if (ListFC.Count > 0 && ListFC != null)
                                {
                                    ((Microsoft.Office.Interop.Excel.Worksheet)xclasseur.Worksheets[1]).Select();
                                    worksheet = xclasseur.ActiveSheet;
                                    worksheet.Name = NomSheet;
                                    splashScreenManager2.SetWaitFormDescription("Veuillez patienter! Traitement en cours...75%");

                                    if (chkFacture.Checked)
                                    {
                                        //Bundle

                                        IsOK = FormatMain.FormatForecastBUNDLEECRIRE(ListFC, worksheet);
                                    }
                                    else
                                    {

                                        IsOK = FormatMain.FormatForecastECRIRE(ListFC, worksheet);
                                    }

                                }


                                splashScreenManager2.SetWaitFormDescription("Veuillez patienter! Traitement en cours...90%");
                                if (IsOK)
                                {
                                    ExcelApplication.ActiveWorkbook.Save();
                                    ExcelApplication.ActiveWorkbook.Close();
                                    ExcelApplication.Quit();

                                    if (ListFC.Count == 0)
                                    {
                                        if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                                        MessageBox.Show("Aucune donnée à exporter en Excel !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                        
                                    }
                                    else
                                    {
                                        if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                                        MessageBox.Show("Génération OK !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    }
                                    
                                }
                                else
                                {
                                    ExcelApplication.ActiveWorkbook.Save();
                                    ExcelApplication.ActiveWorkbook.Close();
                                    ExcelApplication.Quit();
                                    if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                                    MessageBox.Show("Une erreur est survenue lors de la génération !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }


                                #endregion



                                #endregion




                                //if(File.Exists(RepDestXlsSIT))
                                //{

                                //    Microsoft.Office.Interop.Excel.Application xlsApp;
                                //    Microsoft.Office.Interop.Excel.Workbook xlsClasseur;
                                //    Microsoft.Office.Interop.Excel.Worksheet xlsFeuille;

                                //       if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                                //    // Créer un objet Excel
                                //    xlsApp = new Microsoft.Office.Interop.Excel.Application();

                                //     splashScreenManager2.SetWaitFormDescription("Traitement en cours...20%");
                                //    //recupérer le Pid de l'objet pour le tuer à la fin ou en cas d'erreur

                                //    int pidTot = 0;

                                //    pidTot = GetPidExcelApres();

                                //    //On le tuera à la fin ou en cas d'exception
                                //    PidExcelproc = pidTot + processIdAvant;
                                //       splashScreenManager2.SetWaitFormDescription("Traitement en cours...50%");
                                //    if (xlsApp == null)
                                //    {
                                //        ////installer Microsoft Excel sur le poste
                                //         if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                                //        MessageBox.Show("Veuiller installer Excel sur votre poste au préalable!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                //        //tuer le processus créé

                                //        bool iskok = false;

                                //        iskok = TuerProcessus(PidExcelproc);

                                //        if (!iskok)
                                //        {
                                //             if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                                //            MessageBox.Show("Une erreur est survenue lors de la suppression du processus!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                //        }

                                //        return;

                                //    }

                                //    // Ne pas tenir compte des alertes
                                //    xlsApp.DisplayAlerts = false;

                                //    //// Ajout d'un classeur
                                //    xlsClasseur = xlsApp.Workbooks.Open(cheminFichier);
                                //    splashScreenManager2.SetWaitFormDescription("Traitement en cours...70%");

                                //    ((Microsoft.Office.Interop.Excel.Worksheet)xlsClasseur.Worksheets[1]).Select();

                                //    xlsFeuille = (Microsoft.Office.Interop.Excel.Worksheet)xlsClasseur.Worksheets[1];

                                //        splashScreenManager2.SetWaitFormDescription("Traitement en cours...80%");

                                //    bool isImputsOK = false;

                                //    //cas hors facturation Pipe

                                //    if (chkFacture.Checked)
                                //    {
                                //        isImputsOK = FormatMain.FormatForecastBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, xlsApp, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                                //    }
                                //    else
                                //    {
                                //        isImputsOK = FormatMain.FormatForecast(chaineBABI, datedeb, datefin, ListeFilialeChoisies, xlsApp, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                                //    }


                                //    if (!isImputsOK)
                                //    {
                                //        if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                                //        MessageBox.Show("Une erreur est survenue lors de la génération du fichier Excel!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                //    }

                                //      splashScreenManager2.SetWaitFormDescription("Traitement en cours...90%");
                                //    xlsFeuille.Name = NomSheet;

                                //    //xlsFeuille.AutoFilterMode = false;

                                //    xlsApp.ActiveWorkbook.SaveAs(cheminFichier);
                                //    xlsApp.ActiveWorkbook.Close();

                                //    bool isExcelok = false;

                                //    isExcelok = TuerProcessus(PidExcelproc);

                                //    if (!isExcelok)
                                //    {
                                //       if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                                //        MessageBox.Show("Une erreur est survenue lors de la suppression du processus!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                                //    }
                                //    splashScreenManager2.SetWaitFormDescription("Traitement en cours...100%");
                                //    //Generation Excel
                                //    if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                                //    MessageBox.Show("Génération Excel effectuée avec succès!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                //}
                                
                            }

                        }


                        }


                }

            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->sBtnGenererExcel_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }


        /// <summary>
        /// Obtenir somme des pid Excel avant 
        /// </summary>
        /// <returns></returns>
        public int GetPidExcelAvant()
        {
            int ret = 0;
            int processId = 0;
            try
            {
                /////////////////////////////////////////////////////////////////////
                //détecter les autres excel existants ,on ne les tuera pas

                Process[] localByName = Process.GetProcessesByName("EXCEL");

                // Additionner les identifiants des processus "Excel" avant l'ouverture de l'automation :

                foreach (Process item in localByName)
                {
                    processId -= item.Id;

                }

                ret = processId;

                return ret;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                var msg = "MainForm -> GetPidExcelAvant -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return ret;
            }

        }

        /// <summary>
        /// Tuer un processus en prenant son pid en paramètre
        /// </summary>
        /// <param name="Pid"></param>
        /// <returns></returns>
        public bool TuerProcessus(int processToKill)
        {
            bool test = false;
            try
            {
                //Tuer les processus qu'on a créé
                //sassurer que le pid >0 et que c'est un processus EXCEL

                Process localproc = Process.GetProcessById(processToKill);

                if (processToKill > 0 && localproc.ProcessName == "EXCEL")
                {

                    Process.GetProcessById(processToKill).Kill();

                    test = true;

                    return test;
                }

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                var msg = "MainForm -> TuerProcessus -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                test = false;
                return test;
            }
            return test;
        }

        public int GetPidExcelApres()
        {
            int ret = 0;
            int processId = 0;
            try
            {
                /////////////////////////////////////////////////////////////////////
                //détecter les autres excel existants ,on ne les tuera pas

                Process[] localByName = Process.GetProcessesByName("EXCEL");

                // Additionner les identifiants des processus "Excel" avant l'ouverture de l'automation :

                foreach (Process item in localByName)
                {
                    processId += item.Id;

                }

                ret = processId;

                return ret;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                var msg = "MainForm -> GetPidExcelApres -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return ret;
            }

        }

        private void chkTousCom_CheckStateChanged(object sender, EventArgs e)
        {
            try
            {
                if(chkTousCom.CheckState==CheckState.Checked)
                {
                    LECommerciauxA.Enabled = false;
                    LECommerciauxDE.Enabled = false;
                    //Reinitialiser indicateurs montant
                    ReInitialiserIndicateursMontant();
                    //lblControlTotalHT.Text = "0";
                    //lblControlTotalHT2.Text = "0";
                    //lblControlTotalHT3.Text = "0";

                    ////Initialiser fact et acc

                    //lblAccurency.Text = "0";

                    //lblControlRealiseHT.Text = "0";

                    CleanGrid();
                    timer1.Enabled = true;
                }
                else
                {
                    LECommerciauxA.Enabled = true;
                    LECommerciauxDE.Enabled = true;
                    //Reinitialiser indicateurs montant
                    ReInitialiserIndicateursMontant();
                    //lblControlTotalHT.Text = "0";
                    //lblControlTotalHT2.Text = "0";
                    //lblControlTotalHT3.Text = "0";

                    ////Initialiser fact et acc

                    //lblAccurency.Text = "0";

                    //lblControlRealiseHT.Text = "0";
                    CleanGrid();
                    timer1.Enabled = true;
                }

            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
               
                var msg = "MainForm -> chkTousCom_CheckStateChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
               
            }
        }

        private void chkTousTypeDoc_CheckStateChanged(object sender, EventArgs e)
        {
            try
            {
                if(chkTousTypeDoc.CheckState==CheckState.Checked)
                {
                    chkCmbTypeDoc.Enabled = false;

                    //Reinitialiser indicateurs montant
                    ReInitialiserIndicateursMontant();
                    //lblControlTotalHT.Text = "0";
                    //lblControlTotalHT2.Text = "0";
                    //lblControlTotalHT3.Text = "0";

                    ////Initialiser fact et acc

                    //lblAccurency.Text = "0";

                    //lblControlRealiseHT.Text = "0";
                    CleanGrid();
                }
                else
                {
                    chkCmbTypeDoc.Enabled = true;

                    //Reinitialiser indicateurs montant
                    ReInitialiserIndicateursMontant();
                    //lblControlTotalHT.Text = "0";
                    //lblControlTotalHT2.Text = "0";
                    //lblControlTotalHT3.Text = "0";

                    ////Initialiser fact et acc

                    //lblAccurency.Text = "0";

                    //lblControlRealiseHT.Text = "0";
                    CleanGrid();
                }
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkTousTypeDoc_CheckStateChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        
        private void chkCmbTypeDoc_CloseUp(object sender, CloseUpEventArgs e)
        {
            try
            {
                ListDocument = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbTypeDoc.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if(item.Value!=null)
                        {
                            ListDocument += item.Value.ToString() + ",";
                        }
                       
                    }

                }
                if (ListDocument.Length > 0)
                {
                    ListDocument = ListDocument.Substring(0, ListDocument.Length - 1);

                    if(ListDocument.Length==1)
                    {
                        if (ListDocument.Equals("8")) IsFactOnly = true;
                    }
                    else
                    {
                        IsFactOnly = false;
                    }
                    
                }
                
                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkCmbTypeDoc_CloseUp -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }
        

        private void chkMulCom_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulCom.CheckState == CheckState.Checked)
                {
                    chkCmbMultCom.Visible = true;

                    chkTousCom.Visible = false;

                    LECommerciauxA.Visible = false;

                    LECommerciauxDE.Visible = false;

                    lblACom.Visible = false;
                    lblDeCom.Visible = false;
                    timer1.Enabled = true;

                }
                else
                {
                    chkCmbMultCom.Visible = false;

                    chkTousCom.Visible = true;

                    LECommerciauxA.Visible = true;

                    LECommerciauxDE.Visible = true;

                    lblACom.Visible = true;
                    lblDeCom.Visible = true;
                    timer1.Enabled = true;
                }
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkMulCom_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        //recupérer le nom et le prénom à passer au DAO
       

        private void LECommerciauxDE_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomLECommerciauxDE = string.Empty;
                

                ListNomLECommerciauxDE = LECommerciauxDE.Text;

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();

                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LECommerciauxDE_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LECommerciauxA_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomLECommerciauxA = string.Empty;


                ListNomLECommerciauxA = LECommerciauxA.Text;

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LECommerciauxA_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultCom_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomMultiCommerciaux = string.Empty;

                ListPrenomMultiCommerciaux = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultCom.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        //if (item.Value.ToString() != string.Empty)
                        if (!item.Value.ToString().Equals(item.Description))
                        {
                                //le value c'est le prenom ,
                                ListPrenomMultiCommerciaux += item.Value.ToString() + ",";

                            //si la description ramenée= prénom donc on a pas de nom
                            if(item.Value.ToString().Equals(item.Description))
                            {
                                ListNomMultiCommerciaux += " "+",";
                            }
                            else
                            {

                                //on peut donc retirer le nom par déduction d'avec le texte(description =Nom +" "+Prenom)
                                string tmp = " " + item.Value.ToString();
                                ListNomMultiCommerciaux += item.Description.Replace(tmp, "") + ",";
                            }
                            

                        }
                        else
                        {
                            //On a que le nom

                            //le value c'est le prenom donc il prend vide,
                            ListPrenomMultiCommerciaux += " " + ",";
                            
                            ListNomMultiCommerciaux += item.Description + ",";
                           

                        }

                    }

                }

               if(ListPrenomMultiCommerciaux.Length>0) ListPrenomMultiCommerciaux = ListPrenomMultiCommerciaux.Substring(0, ListPrenomMultiCommerciaux.Length - 1);
                if (ListNomMultiCommerciaux.Length > 0) ListNomMultiCommerciaux = ListNomMultiCommerciaux.Substring(0, ListNomMultiCommerciaux.Length - 1);

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                timer1.Enabled = true;
                CleanGrid();
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkCmbMultCom_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkTousFamille_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousFamille.CheckState == CheckState.Checked)
                {
                    LEFamilleDE.Enabled = false;
                    LEFamilleA.Enabled = false;
                }
                else
                {
                    LEFamilleDE.Enabled = true;
                    LEFamilleA.Enabled = true;
                }
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkTousFamille_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

            }
        }

        private void chkCmbMultFam_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomMultiFamille = string.Empty;
                
                foreach (CheckedListBoxItem item in chkCmbMultFam.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if (item.Value != null)
                        {

                            ListNomMultiFamille += item.Value.ToString().Trim() + ",";
                            
                        }

                    }

                }

                if (ListNomMultiFamille.Length > 0) ListNomMultiFamille = ListNomMultiFamille.Substring(0, ListNomMultiFamille.Length - 1);

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkCmbMultFam_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

     


        private void LEFamilleDE_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomLEFamilleDE = string.Empty;


                ListNomLEFamilleDE = LEFamilleDE.Text;

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";
                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LEFamilleDE_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleA_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomLEFamilleA = string.Empty;


                ListNomLEFamilleA = LEFamilleA.Text;
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LEFamilleA_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkMulFamille_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulFamille.CheckState == CheckState.Checked)
                {
                    chkCmbMultFam.Visible = true;

                    chkTousFamille.Visible = false;

                    LEFamilleA.Visible = false;

                    LEFamilleDE.Visible = false;

                    lblFamDe.Visible = false;
                    lblFamA.Visible = false;

                }
                else
                {
                    chkCmbMultFam.Visible = false;

                    chkTousFamille.Visible = true;

                    LEFamilleA.Visible = true;

                    LEFamilleDE.Visible = true;

                    lblFamDe.Visible = true;
                    lblFamA.Visible = true;
                }
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkMulFamille_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void xtraTabControl1_Click(object sender, EventArgs e)
        {
            try
            {
                if(chkTousFamille.CheckState==CheckState.Checked)
                {
                    LEFamilleDE.ItemIndex = 0;
                    LEFamilleA.ItemIndex = 0;
                }
               
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> xtraTabControl1_Click -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkSaisieDevis_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if(chkSaisieDevis.Checked)
                {
                    chkFacture.Checked = false;

                    IsFactOnly = false;

                    panTypeDoc.Visible = false;
                }
                else
                {
                    if (!chkAccepteDevis.Checked && !chkSaisieDevis.Checked && !chkFacture.Checked)
                    {

                        panTypeDoc.Visible = true;
                    }

                }
                
                if(!chkAccepteDevis.Checked && !chkSaisieDevis.Checked && !chkFacture.Checked)
                {
                    lblControlRealiseHT.Text = "0";
                    lblControlRealiseHT2.Text = "0";
                    lblAccurency.Text = "0";
                   lblAccurency2.Text = "0";

                    lblControlTotalHT.Text = "0";
                    lblControlTotalHT2.Text = "0";
                    lblControlTotalHT3.Text = "0";
                }

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkSaisieDevis_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkAccepteDevis_CheckedChanged(object sender, EventArgs e)
        {
            try
            {

                if (chkAccepteDevis.Checked)
                {
                    chkFacture.Checked = false;

                    IsFactOnly = false;

                    panTypeDoc.Visible = false;
                }
                else
                {
                    if (!chkAccepteDevis.Checked && !chkSaisieDevis.Checked && !chkFacture.Checked)
                    {
                        panTypeDoc.Visible = true;
                    }
                }

                if (!chkAccepteDevis.Checked && !chkSaisieDevis.Checked && !chkFacture.Checked)
                {
                    lblControlRealiseHT.Text = "0";
                    lblControlRealiseHT2.Text = "0";
                    lblAccurency.Text = "0";
                    lblAccurency2.Text = "0";

                    lblControlTotalHT.Text = "0";
                    lblControlTotalHT2.Text = "0";
                    lblControlTotalHT3.Text = "0";
                }

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkAccepteDevis_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void dateEditDateDeb_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";

                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> dateEditDateDeb_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void dateEditDateFin_EditValueChanged(object sender, EventArgs e)
        {
            try
            {

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";

                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> dateEditDateFin_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void dateEditDateDeb_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";

                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> dateEditDateDeb_TextChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void dateEditDateFin_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";

                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> dateEditDateFin_TextChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                if (sBtnVisualiserArticle.ForeColor == Color.Black)
                {
                    sBtnVisualiserArticle.ForeColor = Color.Red;
                }
                else
                {
                    if (sBtnVisualiserArticle.ForeColor == Color.Red)
                    {
                        sBtnVisualiserArticle.ForeColor = Color.Black;
                    }
                }
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> timer1_Tick -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LECommerciauxDE_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";

                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LECommerciauxDE_TextChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LECommerciauxA_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";

                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LECommerciauxA_TextChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleDE_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";

                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LEFamilleDE_TextChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleA_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";

                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LEFamilleA_TextChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void BtnStat_Click(object sender, EventArgs e)
        {
            try
            {
                
                //Gestion des bundle donc on ne prend pas en comptes les familles et les marques
                if(!chkFacture.Checked)
                {
                    LcomStat = GetListStatDateByCom();

                    LcomStatTypeDoc = GetListStatTypeDocByCom();

                    LcomStatNbreDoc = GetListStatDateByNbreDocByCom();

                    LcomStatNbreTypeDocCom = GetListStatNbreTypeDocByCom();

                    //=========================================
                    LcomTopFamMtant = GetListTopFamMtant();

                    LcomTopFamCentralisatriceMtant = GetListTopFamCentralisatriceMtant();

                    LcomTopFamQtite = GetListTopFamQtite();

                    LcomTopFamCentralisatriceQtite = GetListTopFamCentralisatriceQtite();
                    //==============================================

                    LcomRanking = GetRankingCom();
                }
                else
                {
                    LcomStat = GetListStatDateByComBUNDLE();

                    LcomStatTypeDoc = GetListStatTypeDocByComBUNDLE();

                    LcomStatNbreDoc = GetListStatDateByNbreDocByCom();
         
                    //2 graphes cumulés
                     LcomStatNbreTypeDocCom = GetListStatNbreTypeDocByComBUNDLE();
                    
                    //On gere les BUNDLE
                    LcomRanking = GetRankingComBUNDLE();
                }
            

          
                

                ComFiltre Cfiltre = new ComFiltre();

                Cfiltre = GetFiltre();


                if ((LcomStat != null && LcomStat.Count > 0) || (LcomStatTypeDoc != null && LcomStatTypeDoc.Count > 0) || (LcomStatNbreDoc != null && LcomStatNbreDoc.Count > 0) || (LcomStatNbreTypeDocCom != null && LcomStatNbreTypeDocCom.Count > 0)  || (LcomTopFamMtant!=null && LcomTopFamMtant.Count>0) || (LcomTopFamCentralisatriceMtant != null && LcomTopFamCentralisatriceMtant.Count > 0) || (LcomTopFamQtite!=null && LcomTopFamQtite.Count>0) || (LcomTopFamCentralisatriceQtite != null && LcomTopFamCentralisatriceQtite.Count > 0) || (LcomRanking != null && LcomRanking.Count > 0))
                {
                    var FenStat = new FenStatistiques(LcomStat, LcomStatTypeDoc, LcomStatNbreDoc, LcomStatNbreTypeDocCom, Cfiltre, LcomTopFamMtant, LcomTopFamCentralisatriceMtant, LcomTopFamQtite, LcomTopFamCentralisatriceQtite, LcomRanking, IsFactOnly);

                    if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                    FenStat.ShowDialog();
                }
                else
                {
                    if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                    MessageBox.Show("Aucune donnée à afficher!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->BtnStat_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
             
            }
        }


        public ComFiltre GetFiltre()
        {
            ComFiltre cf = new ComFiltre();
            try
            {
                cf.mIsPipe = false;

                if(dateEditDateDeb.Text!=string.Empty && dateEditDateFin.Text!=string.Empty)
                {
                    cf.mPeriode = "Période : Du " + dateEditDateDeb.Text + " Au " + dateEditDateFin.Text;
                }

                //Commerciaux

                if(SMulCom.Checked)
                {
                    cf.mCommerciaux = "Commerciaux : " + chkCmbMultCom.Text;
                }
                else
                {
                    if(chkTousCom.Checked)
                    {
                        cf.mCommerciaux = "Commerciaux : Tous les Commerciaux";
                    }
                    else
                    {
                        //DE_A
                        cf.mCommerciaux = "Commerciaux : De " + LECommerciauxDE.Text + " A " + LECommerciauxA.Text;
                    }
                }

                //Famille

                if (SMulFamille.Checked)
                {
                    cf.mFamille = "Famille : " + chkCmbMultFam.Text;
                }
                else
                {
                    if (chkTousFamille.Checked)
                    {
                        cf.mFamille = "Famille : Toutes les Familles";
                    }
                    else
                    {
                        //DE_A
                        cf.mFamille = "Famille : De " + LEFamilleDE.Text + " A " + LEFamilleA.Text;
                    }
                }

                //Famille Centralisatrice(Marque)

                if (SMulFamilleCentral.Checked)
                {
                    cf.mFamilleCentral = "Marque : " + chkCmbMultFamCentral.Text;
                }
                else
                {
                    if (chkTousFamilleCentral.Checked)
                    {
                        cf.mFamilleCentral = "Marque : Toutes les Marques";
                    }
                    else
                    {
                        //DE_A
                        cf.mFamilleCentral = "Marque : De " + LEFamilleCentralDE.Text + " A " + LEFamilleCentralA.Text;
                    }
                }

                //Revendeur

                if (SMulRevendeur.Checked)
                {
                    cf.mRevendeur = "Revendeur : " + chkCmbMultRev.Text;
                }
                else
                {
                    if (chkTousRevendeur.Checked)
                    {
                        cf.mRevendeur = "Revendeur : Tous les Revendeurs";
                    }
                    else
                    {
                        //DE_A
                        cf.mRevendeur = "Revendeur : De " + LERevendeurDE.Text + " A " + LERevendeurA.Text;
                    }
                }

                //Pipe
                if (chkSaisieDevis.Checked && chkAccepteDevis.Checked)
                {
                    cf.mTypeDoc = "Type Document:Tous les Devis";

                    cf.mIsPipe = true;
                }

                if (!chkSaisieDevis.Checked && chkAccepteDevis.Checked)
                {
                   
                    cf.mTypeDoc = "Type Document: 75% Devis Accepté";
                    cf.mIsPipe = false;
                }

                if (chkSaisieDevis.Checked && !chkAccepteDevis.Checked)
                {
                    cf.mTypeDoc = "Type Document: 60% Devis Envoyé";
                    cf.mIsPipe = false;
                }


                if (!chkSaisieDevis.Checked && !chkAccepteDevis.Checked)
                {

                    if(chkFacture.Checked)
                    {
                        cf.mTypeDoc = "Type Document: Facture(FC et FA)";
                    }
                    else
                    {
                        if (chkTousTypeDoc.Checked)
                        {
                            cf.mTypeDoc = "Type Document: Tous";
                        }
                        else
                        {
                            cf.mTypeDoc = "Type Document: " + chkCmbTypeDoc.Text;
                        }
                    }

                  
                }


                //Site choisi

                if(ListeFilialeChoisies.Count>0)
                {

                    string listeSite = string.Empty;

                    foreach(var item in ListeFilialeChoisies)
                    {
                        listeSite = item.mVille + ",";

                    }

                    if (listeSite.Length > 1) listeSite = listeSite.Substring(0, listeSite.Length - 1);

                    cf.mSite = "Site(s) : " + listeSite;
                }

                return cf;
            }
            catch(Exception ex)
            {
              
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetFiltre -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }


        }



        public List<FCommercial> GetListStatDateByCom()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListCom = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }



                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        ListCom = daoReport.GetForecastSTATDateByCom(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);


                        //if (ListCom != null && ListCom.Count > 0)
                        //{
                        //    var SommeSolde = ListCom.FirstOrDefault();

                        //    //cas Hors Pipe

                        //    if (!chkSaisieDevis.Checked && !chkAccepteDevis.Checked)
                        //    {
                        //        lblControlTotalHT.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                        //    }
                        //    else
                        //    {
                        //        lblControlTotalHT.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                        //    }


                        //}
                        //else
                        //{

                        //    lblControlTotalHT.Text = "0";
                        //}


                        //if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                        //gridControlForecast.DataSource = ListCom;

                        //timer1.Enabled = false;

                        //sBtnVisualiserArticle.ForeColor = Color.Black;
                    }


                }

                return ListCom;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListStatDateByCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }

        public List<FCommercial> GetListStatDateByComBUNDLE()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListCom = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }



                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        ListCom = daoReport.GetForecastSTATDateByComBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);


                        //if (ListCom != null && ListCom.Count > 0)
                        //{
                        //    var SommeSolde = ListCom.FirstOrDefault();

                        //    //cas Hors Pipe

                        //    if (!chkSaisieDevis.Checked && !chkAccepteDevis.Checked)
                        //    {
                        //        lblControlTotalHT.Text = SommeSolde.mMontantTotalHorsCasPipe.ToString("n0");
                        //    }
                        //    else
                        //    {
                        //        lblControlTotalHT.Text = SommeSolde.mMontantTotalPipeDevis.ToString("n0");
                        //    }


                        //}
                        //else
                        //{

                        //    lblControlTotalHT.Text = "0";
                        //}


                        //if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                        //gridControlForecast.DataSource = ListCom;

                        //timer1.Enabled = false;

                        //sBtnVisualiserArticle.ForeColor = Color.Black;
                    }


                }

                return ListCom;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListStatDateByComBUNDLE-> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }



        public List<FCommercial> GetListStatDateByNbreDocByCom()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> Liste = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }


                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }


                        Liste = daoReport.GetForecastSTATDateByNbreDocCom(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);


                    }


                }

                return Liste;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListStatDateByNbreDocByCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }


        public List<FCommercial> GetListStatTypeDocByCom()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();
            
            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }



                        ListComTypeDoc = daoReport.GetForecastSTATTypeDocByCom(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);

                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListStatTypeDocByCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }


        public List<FCommercial> GetListStatTypeDocByComBUNDLE()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }



                        ListComTypeDoc = daoReport.GetForecastSTATTypeDocByComBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListStatTypeDocByComBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }



        public List<FCommercial> GetListStatNbreTypeDocByCom()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastSTATNbreTypeDocByCom(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);

                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListStatNbreTypeDocByCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }



        public List<FCommercial> GetListStatNbreTypeDocByComBUNDLE()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastSTATNbreTypeDocByComBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListStatNbreTypeDocByComBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }


        public List<FCommercial> GetListTopFamMtant()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }


                        ListComTypeDoc = daoReport.GetForecastSTATTopFamMtant(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);


                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListTopFamMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }

        public List<FCommercial> GetListTopFamCentralisatriceMtant()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }


                        ListComTypeDoc = daoReport.GetForecastSTATTopFamilleCentrMtant(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);


                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListTopFamMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }



        public List<FCommercial> GetListTopFamQtite()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastSTATTopFamQtite(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);


                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListTopFamQtite -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }

        public List<FCommercial> GetListTopFamCentralisatriceQtite()
        {
            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastSTATTopFamilleCentrQtite(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);


                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListTopFamCentralisatriceQtite -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }




        public List<FCommercial> GetRankingCom()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastRankingCom(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);


                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetRankingCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }

        public List<FCommercial> GetRankingComBUNDLE()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<FCommercial> ListComTypeDoc = new List<FCommercial>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastRankingComBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);


                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetRankingComBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }



        //gestion Target
        private void gestionTargetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //Donner la liste des commerciaux et des familles à la fenêtre

                List<ComClass> LComboCom = new List<ComClass>();
                List<ComFamille> LFamCom = new List<ComFamille>();
                List<CAlias> LBase = new List<CAlias>();

                if (ListeFilialeChoisies.Count==0)
                {
                    //Abidjan
                    LComboCom = LCom.Where(c => c.mIdPays == 1).ToList();
                    LFamCom = LFam.Where(c => c.mIdPays == 1).ToList();

                    LBase= ListeFiliales.Where(c => c.mId == 1).ToList();

                    var FenGestionTarget = new FenGestTarget(LComboCom, LFamCom, LBase);
                    FenGestionTarget.ShowDialog();
                }
                else
                {
                    List<string> LidPays = new List<string>();

                    foreach(var item in ListeFilialeChoisies)
                    {
                        LidPays.Add(item.mId.ToString());

                    }
                    
                    //Remplir le combo commerciaux d'apres les pays choisis=============================
                    LComboCom = daoReport.GetListByListIdCom(LidPays, LCom);

                    //Remplir le combo famille d'apres les pays choisis=============================
                    LFamCom = daoReport.GetListByListIdFam(LidPays, LFam);

                    var FenGestionTarget = new FenGestTarget(LComboCom, LFamCom, ListeFilialeChoisies);
                    FenGestionTarget.ShowDialog();
                }

               
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->gestionTargetToolStripMenuItem_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);

            }
        }

        private void chkTousRevendeur_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (chkTousRevendeur.CheckState == CheckState.Checked)
                {
                    LERevendeurDE.Enabled = false;
                    LERevendeurA.Enabled = false;
                }
                else
                {
                    LERevendeurDE.Enabled = true;
                    LERevendeurA.Enabled = true;
                }
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkTousRevendeur_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

            }
        }

        private void SMulRevendeur_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulRevendeur.CheckState == CheckState.Checked)
                {
                    chkCmbMultRev.Visible = true;

                    chkTousRevendeur.Visible = false;

                    LERevendeurA.Visible = false;

                    LERevendeurDE.Visible = false;

                    lblRevDe.Visible = false;
                    lblRevA.Visible = false;

                }
                else
                {
                    chkCmbMultRev.Visible = false;

                    chkTousRevendeur.Visible = true;

                    LERevendeurA.Visible = true;

                    LERevendeurDE.Visible = true;

                    lblRevDe.Visible = true;
                    lblRevA.Visible = true;
                }
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> SMulRevendeur_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultRev_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomMultiRevendeur = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultRev.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if (item.Value != null)
                        {

                            ListNomMultiRevendeur += item.Value.ToString().Trim() + ",";

                        }

                    }

                }

                if (ListNomMultiRevendeur.Length > 0) ListNomMultiRevendeur = ListNomMultiRevendeur.Substring(0, ListNomMultiRevendeur.Length - 1);

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkCmbMultRev_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void sBtnStatRev_Click(object sender, EventArgs e)
        {
            try
            {
                //Gestion des bundle donc on ne prend pas en comptes les familles et les marques
                if (!chkFacture.Checked)
                {
                    LcomStatTypeDocRevendeur = GetListStatTypeDocByRevendeur();

                    //Transaction des commerciaux pour les revendeurs (en % Montant)

                    LcomStatComRevendeur = GetForecastSTATTopComRevendeurMtant();

                    LcomStatTOPRevendeur = GetForecastSTATTopBESTRevendeurMtant();

                    LcomStatBADRevendeur = GetForecastSTATBADRevendeurMtant();
                }
                else
                {
                    LcomStatTypeDocRevendeur = GetListStatTypeDocByRevendeurBUNDLE();

                    //Transaction des commerciaux pour les revendeurs (en % Montant)

                    LcomStatComRevendeur = GetForecastSTATTopComRevendeurMtantBUNDLE();

                    LcomStatTOPRevendeur = GetForecastSTATTopBESTRevendeurMtantBUNDLE();

                    LcomStatBADRevendeur = GetForecastSTATBADRevendeurMtantBUNDLE();
                }

                 

                ComFiltre Cfiltre = new ComFiltre();

                Cfiltre = GetFiltre();


                if ((LcomStatTypeDocRevendeur != null && LcomStatTypeDocRevendeur.Count > 0 || LcomStatComRevendeur.Count > 0 || LcomStatTOPRevendeur.Count>0|| LcomStatBADRevendeur.Count>0) )
                {
                    var FenStat = new FenStatRevendeur(LcomStatTypeDocRevendeur, Cfiltre, LcomStatComRevendeur, LcomStatTOPRevendeur, LcomStatBADRevendeur, IsFactOnly);

                    if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                    FenStat.ShowDialog();
                }
                else
                {
                    if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                    MessageBox.Show("Aucune donnée à afficher!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> sBtnStatRev_Click -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        public List<ComRevendeur> GetListStatTypeDocByRevendeur()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<ComRevendeur> ListComTypeDoc = new List<ComRevendeur>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastSTATTypeDocByRevendeur(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);

                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListStatTypeDocByRevendeur -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }

        public List<ComRevendeur> GetListStatTypeDocByRevendeurBUNDLE()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<ComRevendeur> ListComTypeDoc = new List<ComRevendeur>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }
                        
                        ListComTypeDoc = daoReport.GetForecastSTATTypeDocByRevendeurBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                    }


                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetListStatTypeDocByRevendeurBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }

        public List<ComRevendeur> GetForecastSTATTopComRevendeurMtant()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<ComRevendeur> ListComTypeDoc = new List<ComRevendeur>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }


                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastSTATTopComRevendeurMtant(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);

                    }

                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetForecastSTATTopComRevendeurMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }

        public List<ComRevendeur> GetForecastSTATTopComRevendeurMtantBUNDLE()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<ComRevendeur> ListComTypeDoc = new List<ComRevendeur>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }


                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }
                        
                        ListComTypeDoc = daoReport.GetForecastSTATTopComRevendeurMtantBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                    }

                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetForecastSTATTopComRevendeurMtantBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }



        public List<ComRevendeur> GetForecastSTATTopBESTRevendeurMtant()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<ComRevendeur> ListComTypeDoc = new List<ComRevendeur>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastSTATTopBESTRevendeurMtant(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);

                    }

                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetForecastSTATTopBESTRevendeurMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }


        public List<ComRevendeur> GetForecastSTATTopBESTRevendeurMtantBUNDLE()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<ComRevendeur> ListComTypeDoc = new List<ComRevendeur>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                        ListComTypeDoc = daoReport.GetForecastSTATTopBESTRevendeurMtantBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                    }

                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetForecastSTATTopBESTRevendeurMtantBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }

        public List<ComRevendeur> GetForecastSTATBADRevendeurMtant()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<ComRevendeur> ListComTypeDoc = new List<ComRevendeur>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }


                        ListComTypeDoc = daoReport.GetForecastSTATBADRevendeurMtant(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice,chkFacture.Checked);

                    }

                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetForecastSTATBADRevendeurMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }


        public List<ComRevendeur> GetForecastSTATBADRevendeurMtantBUNDLE()
        {


            if (dateEditDateDeb.Text == string.Empty || dateEditDateFin.Text == string.Empty)
            {
                MessageBox.Show("Veuillez renseigner correctement les dates!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            List<ComRevendeur> ListComTypeDoc = new List<ComRevendeur>();

            try
            {
                string datedeb = DateTime.Parse(dateEditDateDeb.EditValue.ToString()).ToShortDateString();

                string datefin = DateTime.Parse(dateEditDateFin.EditValue.ToString()).ToShortDateString();

                //recupérer la chaine de Abidjan

                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                if (ChooseBABI != null)
                {
                    string chaineBABI = ChooseBABI.mAliasName;

                    if (ListeFilialeChoisies.Count == 0)
                    {
                        //Verifier qu'on a rien choisi

                        if (ChkListBoxControlFiliale.CheckedItemsCount > 0)
                        {
                            //On a choisi Abidjan par défaut

                            ListeFilialeChoisies = ListeFiliales.Where(c => c.IsAbidjan == true).ToList();

                        }
                        else
                        {
                            //Choisir au moins une société
                            MessageBox.Show("Veuillez choisir au moins une société avant l'aperçu!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return null;
                        }

                    }

                    if (ListeFilialeChoisies.Count > 0)
                    {
                        if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                        //Cas où on a pas choisi des famille ou commerciaux par fermeture des combo(de_A)

                        //Famille
                        if (ListNomLEFamilleDE == string.Empty || ListNomLEFamilleDE == null)
                        {
                            ListNomLEFamilleDE = LEFamilleDE.Text;
                        }

                        if (ListNomLEFamilleA == string.Empty || ListNomLEFamilleA == null)
                        {
                            ListNomLEFamilleA = LEFamilleA.Text;
                        }

                        //Commerciaux
                        if (ListNomLECommerciauxDE == string.Empty || ListNomLECommerciauxDE == null)
                        {
                            ListNomLECommerciauxDE = LECommerciauxDE.Text;
                        }

                        if (ListNomLECommerciauxA == string.Empty || ListNomLECommerciauxA == null)
                        {
                            ListNomLECommerciauxA = LECommerciauxA.Text;
                        }

                        //Revendeur
                        if (ListNomLERevendeurDE == string.Empty || ListNomLERevendeurDE == null)
                        {
                            ListNomLERevendeurDE = LERevendeurDE.Text;
                        }

                        if (ListNomLERevendeurA == string.Empty || ListNomLERevendeurA == null)
                        {
                            ListNomLERevendeurA = LERevendeurA.Text;
                        }

                        //Famille Central
                        if (ListNomLEFamilleCentrDE == string.Empty || ListNomLEFamilleCentrDE == null)
                        {
                            ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                        }

                        if (ListNomLEFamilleCentrA == string.Empty || ListNomLEFamilleCentrA == null)
                        {
                            ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                        }

                       
                        ListComTypeDoc = daoReport.GetForecastSTATBADRevendeurMtantBUNDLE(chaineBABI, datedeb, datefin, ListeFilialeChoisies, SMulCom.Checked, chkTousCom.Checked, ListNomLECommerciauxDE, ListNomLECommerciauxA, ListNomMultiCommerciaux, ListPrenomMultiCommerciaux, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDE, ListNomLEFamilleA, ListNomMultiFamille, chkTousTypeDoc.Checked, ListDocument, chkSaisieDevis.Checked, chkAccepteDevis.Checked, SMulRevendeur.Checked, chkTousRevendeur.Checked, ListNomLERevendeurDE, ListNomLERevendeurA, ListNomMultiRevendeur, chkConfigFamilles.Checked, SMulFamilleCentral.Checked, chkTousFamilleCentral.Checked, ListNomLEFamilleCentrDE, ListNomLEFamilleCentrA, ListNomMultiFamilleCentralisatrice, chkFacture.Checked);

                    }

                }

                return ListComTypeDoc;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->GetForecastSTATBADRevendeurMtantBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }
        }



        private void LERevendeurDE_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomLERevendeurDE = string.Empty;


                ListNomLERevendeurDE = LERevendeurDE.Text;
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LERevendeurDE_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LERevendeurA_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomLERevendeurA = string.Empty;


                ListNomLERevendeurA = LERevendeurA.Text;

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LERevendeurA_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void connexionsToolStripMenuItem1_Click(object sender, EventArgs e)
        {

            try
            {

                if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                var FenConnexion = new Connexion();

                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                FenConnexion.ShowDialog();

                ReloadApplication();

                //Rafraichir la grid des filiales

                ListeFiliales = daoMain.GetAliasConnexions();

                //Remplir les filiales à selectionner(on exclu alors abidjan)

                if (ListeFiliales != null)
                {
                    ////ListeAutreFiliales.Clear();
                    //var ListeAutreFiliales = ListeFiliales.Where(c => c.IsAbidjan != true).ToList();

                    RemplirFilialesCheckedControl(ChkListBoxControlFiliale, ListeFiliales);

                }

                //Remplir les filiales à selectionner(on exclu alors abidjan)


            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->connexionsToolStripMenuItem_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

       

        private void gestionTargetToolStripMenuItem2_Click(object sender, EventArgs e)
        {

            try
            {
                //Donner la liste des commerciaux et des familles à la fenêtre

                List<ComClass> LComboCom = new List<ComClass>();
                List<ComFamille> LFamCom = new List<ComFamille>();
                List<CAlias> LBase = new List<CAlias>();

                if (ListeFilialeChoisies.Count == 0)
                {
                    //Abidjan
                    LComboCom = LCom.Where(c => c.mIdPays == 1).ToList();
                    LFamCom = LFam.Where(c => c.mIdPays == 1).ToList();

                    LBase = ListeFiliales.Where(c => c.mId == 1).ToList();

                    var FenGestionTarget = new FenGestTarget(LComboCom, LFamCom, LBase);
                    FenGestionTarget.ShowDialog();
                }
                else
                {
                    List<string> LidPays = new List<string>();

                    foreach (var item in ListeFilialeChoisies)
                    {
                        LidPays.Add(item.mId.ToString());

                    }

                    //Remplir le combo commerciaux d'apres les pays choisis=============================
                    LComboCom = daoReport.GetListByListIdCom(LidPays, LCom);

                    //Remplir le combo famille d'apres les pays choisis=============================
                    LFamCom = daoReport.GetListByListIdFam(LidPays, LFam);

                    var FenGestionTarget = new FenGestTarget(LComboCom, LFamCom, ListeFilialeChoisies);
                    FenGestionTarget.ShowDialog();
                }


            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->gestionTargetToolStripMenuItem_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);

            }
        }

        private void connexionToolStripMenuItem_Click(object sender, EventArgs e)
        {

            try
            {
                //if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                ConAdmin FenConnexionTarget = new ConAdmin();

                //  if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                FenConnexionTarget.ShowDialog();
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->connexionToolStripMenuItem_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            try
            {

                List<String> Lidpays = new List<string>();
                List<ComClass> LComboCom = new List<ComClass>();
                List<ComFamille> LComboFam = new List<ComFamille>();
                List<ComFamille> LComboMarque = new List<ComFamille>();

                List<ComRevendeur> LComboRev = new List<ComRevendeur>();

                ListeFilialeChoisies.Clear();

                foreach (CheckedListBoxItem item in ChkListBoxControlFiliale.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        var Choose = ListeFiliales.FirstOrDefault(c => c.mId == Convert.ToInt32(item.Value.ToString()));

                        ListeFilialeChoisies.Add(Choose);
                        Lidpays.Add(item.Value.ToString());
                    }

                }

                
                // Abidjan
                var ChooseBABI = ListeFiliales.FirstOrDefault(c => c.IsAbidjan == true);

                string chaineBABI = string.Empty;

                if (ChooseBABI != null)
                {
                     chaineBABI = ChooseBABI.mAliasName;
                }


                //Remplir le combo commerciaux d'apres les pays choisis=============================
                LComboCom = daoReport.GetListByListIdCom(Lidpays, LCom);
                
                //Remplir Combo famille===========================================================

                LComboFam = daoReport.GetListByListIdFam(Lidpays, LFam);

                //Remplir Combo Marque===========================================================

                LComboMarque = daoReport.GetListByListIdFam(Lidpays, LFamCentr);

                //Remplir Combo Revendeur===========================================================

                LComboRev = daoReport.GetListByListIdRev(Lidpays, LRev);
                
                var fem = new FenAnalyseComp(chaineBABI,ListeFilialeChoisies, LComboCom, LComboFam, LComboRev, LComboMarque);

                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                fem.ShowDialog();
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->simpleButton1_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }


        //Masquer marque dans le cas où on veut gérer tout ce qui 
        //est lié au bundle (100% PIPE)

        private void HideMarqueBUNDLE()
        {
            try
            {
                chkConfigFamilles.Checked = false;
                chkConfigFamilles.Visible = false;
                panMarque.Visible = false;
                label7.Visible = false;

            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->HideMarqueBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }


        //Montrer marque au cas où cest caché

        private void ShowMarqueBUNDLE()
        {
            try
            {
                label7.Visible = true;
                chkConfigFamilles.Visible = true;
                panMarque.Visible = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm ->ShowMarqueBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }



        private void chkConfigFamilles_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkConfigFamilles.CheckState == CheckState.Checked)
                {
                    panFamille.Visible = true;

                    //Recharger les familles par rapport aux familles central

                    ReLoadFamilleByMarque(chkTousFamilleCentral.Checked,SMulFamilleCentral.Checked,LEFamilleCentralDE.Text,LEFamilleCentralA.Text,chkCmbMultFamCentral.Text);
                }
                else
                {
                    panFamille.Visible = false;
                }
                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkConfigFamilles_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }


        }


        private void CancelConfigFamille()
        {
            try
            {
                // SMulFamille.Checked = false;

                //chkTousFamille.Checked = true;
                
                chkConfigFamilles.Checked = false;

            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> CancelConfigFamille -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }




        private void ReLoadFamilleByMarque(bool IsTousMarque,bool IsMultiMarque,string MarqueDE,string MarqueA,string MultiMarque)
        {
            List<string> Lidpays = new List<string>();
            List<string> LFamCentral = new List<string>();
            List<ComFamille> LFamilleFill = new List<ComFamille>();
            try
            {
                foreach (CheckedListBoxItem item in ChkListBoxControlFiliale.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        var Choose = ListeFiliales.FirstOrDefault(c => c.mId == Convert.ToInt32(item.Value.ToString()));
                        
                        Lidpays.Add(item.Value.ToString());
                    }

                }
                

                if(IsMultiMarque)
                {
                    var tab = MultiMarque.Split(',');

                    for(int i=0;i<tab.Length;i++)
                    {
                        LFamCentral.Add(tab[i].Trim());
                    }
                    
                    LFamilleFill = daoReport.GetListByListIdPaysFamCentrFam(Lidpays, LFam, LFamCentral);
                    LEFamilleDE.Properties.DataSource = LFamilleFill;
                    LEFamilleA.Properties.DataSource = LFamilleFill;

                    //Vider le multifamille ET LES DE_A au cas où il y aurait eu un précédent choix
                    ListNomMultiFamille = string.Empty;
                    ListNomLEFamilleDE = string.Empty;
                    ListNomLEFamilleA = string.Empty;
                    /////////////////////////////////////////////////////////
                    chkCmbMultFam.Properties.DataSource = LFamilleFill;
                    FillMultiCheckComboFamille(chkCmbMultFam, LFamilleFill);

                    LEFamilleDE.Properties.DisplayMember = "mFa_CodeFamille";
                    LEFamilleA.Properties.DisplayMember = "mFa_CodeFamille";
                }
                else
                {
                    if (IsTousMarque)
                    {
                        if (Lidpays.Count > 0)
                        {
                            LFamilleFill = daoReport.GetListByListIdFam(Lidpays, LFam);

                            LEFamilleDE.Properties.DataSource = LFamilleFill;
                            LEFamilleA.Properties.DataSource = LFamilleFill;
                            //Vider le multifamille ET LES DE_A au cas où il y aurait eu un précédent choix
                            ListNomMultiFamille = string.Empty;
                            ListNomLEFamilleDE = string.Empty;
                            ListNomLEFamilleA = string.Empty;
                            /////////////////////////////////////////////////////////

                            chkCmbMultFam.Properties.DataSource = LFamilleFill;
                            FillMultiCheckComboFamille(chkCmbMultFam, LFamilleFill);

                            LEFamilleDE.Properties.DisplayMember = "mFa_CodeFamille";
                            LEFamilleA.Properties.DisplayMember = "mFa_CodeFamille";
                        }


                    }
                    else
                    {
                        //DE_A
                        bool trouveDeb = false;
                        bool trouveFin = false;

                        //récupérer la liste des fam central qui se trouve dans l'intervalle

                        List<string> ListFamCentralisatrice = new List<string>();
                        List<ComFamille> ListFCTMP = new List<ComFamille>();

                        foreach (var item in Lidpays)
                        {
                            List<ComFamille> ListFC = new List<ComFamille>();
                            ListFC = LFamCentr.Where(c => c.mIdPays == Convert.ToInt32(item.ToString())).ToList();

                            ListFCTMP.AddRange(ListFC);
                        }


                        if(ListFCTMP.Count>0)
                        {
                            ListFCTMP = ListFCTMP.OrderBy(x => x.mFa_CodeFamille).ToList();

                            foreach (var obj in ListFCTMP)
                            {
                                if(MarqueDE == obj.mFa_CodeFamille)
                                {
                                    trouveDeb = true;

                                    ListFamCentralisatrice.Add(obj.mFa_CodeFamille);
                                }

                                if(trouveDeb && !trouveFin)
                                {
                                    var IsExist = ListFamCentralisatrice.FirstOrDefault(c=>c.Equals(obj.mFa_CodeFamille));

                                  if(IsExist==null)  ListFamCentralisatrice.Add(obj.mFa_CodeFamille);
                                }

                                if (MarqueA==obj.mFa_CodeFamille)
                                {
                                    trouveFin = true;
                                    var IsExist = ListFamCentralisatrice.FirstOrDefault(c => c.Equals(obj.mFa_CodeFamille));

                                    if (IsExist == null) ListFamCentralisatrice.Add(obj.mFa_CodeFamille);
                                }

                                if(trouveDeb && trouveFin)
                                {
                                    break;
                                }

                            }

                            LFamilleFill = daoReport.GetListByListIdPaysFamCentrFam(Lidpays, LFam, ListFamCentralisatrice);

                            LEFamilleDE.Properties.DataSource = LFamilleFill;
                            LEFamilleA.Properties.DataSource = LFamilleFill;

                            //Vider le multifamille ET LES DE_A au cas où il y aurait eu un précédent choix
                            ListNomMultiFamille = string.Empty;
                            ListNomLEFamilleDE = string.Empty;
                            ListNomLEFamilleA = string.Empty;
                            /////////////////////////////////////////////////////////

                            chkCmbMultFam.Properties.DataSource = LFamilleFill;
                            FillMultiCheckComboFamille(chkCmbMultFam, LFamilleFill);

                            LEFamilleDE.Properties.DisplayMember = "mFa_CodeFamille";
                            LEFamilleA.Properties.DisplayMember = "mFa_CodeFamille";
                        }



                    }


                }

                

            }
            catch(Exception ex)
            {

                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "MainForm -> ReLoadFamilleByMarque -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

            }
        }




        private void chkTousFamilleCentral_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousFamilleCentral.CheckState == CheckState.Checked)
                {
                    LEFamilleCentralDE.Enabled = false;
                    LEFamilleCentralA.Enabled = false;
                }
                else
                {
                    LEFamilleCentralDE.Enabled = true;
                    LEFamilleCentralA.Enabled = true;
                }

                CancelConfigFamille();

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkTousFamilleCentral_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

            }
        }

        private void SMulFamilleCentral_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulFamilleCentral.CheckState == CheckState.Checked)
                {
                    chkCmbMultFamCentral.Visible = true;

                    chkTousFamilleCentral.Visible = false;

                    LEFamilleCentralA.Visible = false;

                    LEFamilleCentralDE.Visible = false;

                    lblFamCentralDe.Visible = false;
                    lblFamCentralA.Visible = false;

                }
                else
                {
                    chkCmbMultFamCentral.Visible = false;

                    chkTousFamilleCentral.Visible = true;

                    LEFamilleCentralA.Visible = true;

                    LEFamilleCentralDE.Visible = true;

                    lblFamCentralDe.Visible = true;
                    lblFamCentralA.Visible = true;
                }
                CancelConfigFamille();

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();

                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();

            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> SMulFamilleCentral_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleCentralDE_TextChanged(object sender, EventArgs e)
        {
            try
            {
                CancelConfigFamille();
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LEFamilleCentralDE_TextChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleCentralA_TextChanged(object sender, EventArgs e)
        {
            try
            {
                CancelConfigFamille();
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LEFamilleCentralA_TextChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultFamCentral_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomMultiFamilleCentralisatrice = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultFamCentral.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if (item.Value != null)
                        {
                            ListNomMultiFamilleCentralisatrice += item.Value.ToString().Trim() + ",";

                        }

                    }

                }

                if (ListNomMultiFamilleCentralisatrice.Length > 0) ListNomMultiFamilleCentralisatrice = ListNomMultiFamilleCentralisatrice.Substring(0, ListNomMultiFamilleCentralisatrice.Length - 1);

                //Reinitialiser indicateurs montant
                ReInitialiserIndicateursMontant();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";

                ////Initialiser fact et acc

                //lblAccurency.Text = "0";

                //lblControlRealiseHT.Text = "0";
                CleanGrid();
                CancelConfigFamille();
                timer1.Enabled = true;
                
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkCmbMultFamCentral_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleCentralDE_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                ListNomLEFamilleCentrDE = LEFamilleCentralDE.Text;
                CancelConfigFamille();
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LEFamilleCentralDE_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleCentralA_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                ListNomLEFamilleCentrA = LEFamilleCentralA.Text;
                CancelConfigFamille();
            }
            catch (Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> LEFamilleCentralA_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkFacture_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                //Ne pas que les facture soient cochées avec les devis
                if(chkFacture.Checked)
                {
                    chkAccepteDevis.Checked = false;
                    chkSaisieDevis.Checked = false;

                    HideMarqueBUNDLE();

                    panTypeDoc.Visible = false;

                    IsFactOnly = true;
                }
                else
                {
                    IsFactOnly = false;
                    ShowMarqueBUNDLE();
                    panTypeDoc.Visible = true;
                }

                if (!chkAccepteDevis.Checked && !chkSaisieDevis.Checked && !chkFacture.Checked)
                {
                    lblControlRealiseHT.Text = "0";
                    lblControlRealiseHT2.Text = "0";
                    lblAccurency.Text = "0";
                    lblAccurency2.Text = "0";
                    lblControlTotalHT.Text = "0";
                    lblControlTotalHT2.Text = "0";
                    lblControlTotalHT3.Text = "0";
                }

                lblAccurency.Text = "0";
                lblAccurency2.Text = "0";
                //Réalisé
                //lblControlRealiseHT.Text = "0";
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                //lblControlTotalHT3.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "MainForm -> chkFacture_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }
    }
}
