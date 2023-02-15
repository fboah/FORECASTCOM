using DevExpress.XtraCharts;
using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using ForecastCom.DAO;
using ForecastCom.Models;
using ForecastCom.Services;
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

namespace ForecastCom.Statistiques
{
    public partial class FenAnalyseComp : Form
    {
        private readonly FormatageForecastCom FormatMain = new FormatageForecastCom();
        private List<ComClass> myObjectCommercial;

        private List<ComFamille> myObjectFamille;

        private List<ComFamille> myObjectMarque;

        private List<ComRevendeur> myObjectRevendeur;

        private List<CAlias> myObjectFilialeChoisies;

        //Liste des années à rechercher
        private List<string>ListeAnnees;

       //Analyse Commerciale++++++++++++++++++++++++++++++++++
        //NomCommercial a transmettre DAO
        private string DAONomCommercial;

        //preNomCommercial a transmettre DAO
        private string DAOPrenomCommercial;

        //Analyse Famille+++++++++++++++++++++++++++++++++++
        private string DAOFamille;

        //Analyse Revendeur+++++++++++++++++++++++++++++++++
        private string DAORevendeur;

        //Analyse Marque+++++++++++++++++++++++++++++++++++
        private string DAOMarque;

        //Chaine pour NOM MulticheckCommerciaux

        private string ListNomMultiCommerciauxCritere;

        //Chaine pour PRENOM MulticheckCommerciaux

        private string ListPrenomMultiCommerciauxCritere;


        //Multi critere (on prend cette option pour eviter les espaces dans les nom ramenés)
        
        private string ListMultiRevendeurCritere;
        private string ListMultiMarqueCritere;


        //Multi Axe (on prend cette option pour eviter les espaces dans les nom ramenés)
        private string ListMultiFamilleAxe;
        private string ListMultiRevendeurAxe;
        private string ListMultiMarqueAxe;



        //=============Multiple Axe===================

        //Chaine pour NOM MulticheckCommerciaux

        private string ListNomMultiCommerciauxMultipleAxe;

        //Chaine pour PRENOM MulticheckCommerciaux

        private string ListPrenomMultiCommerciauxMultipleAxe;

        //===================================================

        private readonly DAOForecastCom daoReport = new DAOForecastCom();

        //Variables pour les familles d'apres configuration des marques
        private string ListMultiFamilleCritere;
                   private string ListNomLEFamilleDECritere ;
                  private string  ListNomLEFamilleACritere ;
        
        //Periode Analyse
        private bool IsMensuel;
        private bool IsTrimetre;
        private bool IsSemetre;
        private bool IsPeriodeAdefinir;
        private bool IsAnnuel;

        //Axe d'Analyse
        private bool IsCommercialAxe;
        private bool IsFamilleAxe;
        private bool IsMarqueAxe;
        private bool IsRevendeurAxe;

        //Axe d'Analyse Multiple
        private bool IsCommercialAxeMultiple;
        private bool IsFamilleAxeMultiple;
        private bool IsMarqueAxeMultiple;
        private bool IsRevendeurAxeMultiple;


        private string ChaineconnexionBABI;

        public FenAnalyseComp(string chaineBABI,List<CAlias> ListeFilialeChoisies,List<ComClass>LComboCom, List<ComFamille> LComboFam, List<ComRevendeur> LComboRev, List<ComFamille> LComboMarque)
        {
            InitializeComponent();

            this.myObjectCommercial = LComboCom.OrderBy(x=>x.mNomCommercial).ToList();
            this.myObjectMarque = LComboMarque.OrderBy(x=>x.mFa_CodeFamille).ToList();
            this.myObjectFamille = LComboFam.OrderBy(x => x.mFa_CodeFamille).ToList();
            this.myObjectRevendeur = LComboRev.OrderBy(x=>x.mCT_Num).ToList();
            this.myObjectFilialeChoisies = ListeFilialeChoisies;
            this.ChaineconnexionBABI = chaineBABI;
        }

        private void FenAnalyseComp_Load(object sender, EventArgs e)
        {
            try
            {
                if (!splashScreenManager1.IsSplashFormVisible) splashScreenManager1.ShowWaitForm();

                #region Chargement Combo

                #region Axe Analyse

                ////////////////Commercial
                LECommercial.Properties.DataSource = myObjectCommercial;

                LECommercial.Properties.DisplayMember = "mNomCommercial";
                LECommercial.Properties.ValueMember = "mPrenomCommercial";

                //Choisir les premières valeurs
                LECommercial.ItemIndex = 0;

                ////////////////Famille

                LEFamille.Properties.DataSource = myObjectFamille;

                LEFamille.Properties.DisplayMember = "mFa_CodeFamille";
                LEFamille.Properties.ValueMember = "mFa_CodeFamille";

                //Choisir les premières valeurs
                LEFamille.ItemIndex = 0;
               

                ////////////////Marque

                LEMarque.Properties.DataSource = myObjectMarque;

                LEMarque.Properties.DisplayMember = "mFa_CodeFamille";
                LEMarque.Properties.ValueMember = "mFa_CodeFamille";

                //Choisir les premières valeurs
                LEMarque.ItemIndex = 0;


                ///////////////Revendeurs

                LERevendeur.Properties.DataSource = myObjectRevendeur;

                LERevendeur.Properties.DisplayMember = "mCT_Num";
                

                //Choisir les premières valeurs
                LERevendeur.ItemIndex = 0;


                //==================MULTIPLE=================//

                #region Multi Axe

                ////////////Commercial
                LECommercialDEMultAxe.Properties.DataSource = myObjectCommercial;
                LECommercialAMultAxe.Properties.DataSource = myObjectCommercial;

                LECommercialDEMultAxe.Properties.DisplayMember = "mNomCommercial";
                LECommercialAMultAxe.Properties.DisplayMember = "mNomCommercial";

                chkCmbMultComAxe.Properties.DataSource = myObjectCommercial;
                FillMultiCheckComboCommercial(chkCmbMultComAxe, myObjectCommercial);

                //LEFamilleDE.Properties.DisplayMember = "mCT_Intitule";
                //LEFamilleA.Properties.DisplayMember = "mCT_Intitule";

                //Choisir les premières valeurs
                LECommercialDEMultAxe.ItemIndex = 0;
                LECommercialAMultAxe.ItemIndex = 0;

                //////////Marque
                LEMarqueDEMultAxe.Properties.DataSource = myObjectMarque;
                LEMarqueAMultAxe.Properties.DataSource = myObjectMarque;

                chkCmbMultMarqueAxe.Properties.DataSource = myObjectMarque;
                FillMultiCheckComboFamille(chkCmbMultMarqueAxe, myObjectMarque);

                LEMarqueDEMultAxe.Properties.DisplayMember = "mFa_CodeFamille";
                LEMarqueAMultAxe.Properties.DisplayMember = "mFa_CodeFamille";

                //Choisir les premières valeurs
                LEMarqueDEMultAxe.ItemIndex = 0;
                LEMarqueAMultAxe.ItemIndex = 0;

                //REvendeur

                LERevendeurDEMultAxe.Properties.DataSource = myObjectRevendeur;
                LERevendeurAMultAxe.Properties.DataSource = myObjectRevendeur;

                chkCmbMultRevAxe.Properties.DataSource = myObjectRevendeur;
                FillMultiCheckComboRevendeur(chkCmbMultRevAxe, myObjectRevendeur);

                LERevendeurDEMultAxe.Properties.DisplayMember = "mCT_Num";
                LERevendeurAMultAxe.Properties.DisplayMember = "mCT_Num";

                //Choisir les premières valeurs
                LERevendeurDEMultAxe.ItemIndex = 0;
                LERevendeurAMultAxe.ItemIndex = 0;

                //////////////Famille

                LEFamilleDEMultAxe.Properties.DataSource = myObjectFamille;
                LEFamilleAMultAxe.Properties.DataSource = myObjectFamille;

                chkCmbMultFamAxe.Properties.DataSource = myObjectFamille;
                FillMultiCheckComboFamille(chkCmbMultFamAxe, myObjectFamille);

                LEFamilleDEMultAxe.Properties.DisplayMember = "mFa_CodeFamille";
                LEFamilleAMultAxe.Properties.DisplayMember = "mFa_CodeFamille";

                //Choisir les premières valeurs
                LEFamilleDEMultAxe.ItemIndex = 0;
                LEFamilleAMultAxe.ItemIndex = 0;

                #endregion

                #endregion

                sNumAnneeDeb.Value = DateTime.Now.Year;
                sNumAnneeFin.Value = DateTime.Now.Year;

                //Remplir Combo Famille==============================================

                LEFamilleDE.Properties.DataSource = myObjectFamille;
                LEFamilleA.Properties.DataSource = myObjectFamille;

                chkCmbMultFam.Properties.DataSource = myObjectFamille;
                FillMultiCheckComboFamille(chkCmbMultFam, myObjectFamille);

                LEFamilleDE.Properties.DisplayMember = "mFa_CodeFamille";
                LEFamilleA.Properties.DisplayMember = "mFa_CodeFamille";

                //Choisir les premières valeurs
                LEFamilleDE.ItemIndex = 0;
                LEFamilleA.ItemIndex = 0;

                //Remplir Combo Marque==============================================

                LEMarqueDE.Properties.DataSource = myObjectMarque;
                LEMarqueA.Properties.DataSource = myObjectMarque;

                chkCmbMultMarque.Properties.DataSource = myObjectMarque;
                FillMultiCheckComboFamille(chkCmbMultMarque, myObjectMarque);

                LEMarqueDE.Properties.DisplayMember = "mFa_CodeFamille";
                LEMarqueA.Properties.DisplayMember = "mFa_CodeFamille";

                //Choisir les premières valeurs
                LEMarqueDE.ItemIndex = 0;
                LEMarqueA.ItemIndex = 0;


                //Remplir Combo Revendeur==============================================

                LERevendeurDE.Properties.DataSource = myObjectRevendeur;
                LERevendeurA.Properties.DataSource = myObjectRevendeur;

                chkCmbMultRev.Properties.DataSource = myObjectRevendeur;
                FillMultiCheckComboRevendeur(chkCmbMultRev, myObjectRevendeur);

                LERevendeurDE.Properties.DisplayMember = "mCT_Num";
                LERevendeurA.Properties.DisplayMember = "mCT_Num";

                //Choisir les premières valeurs
                LERevendeurDE.ItemIndex = 0;
                LERevendeurA.ItemIndex = 0;

                //Remplir Combo Commerciaux==============================================
                
                LECommercialDE.Properties.DataSource = myObjectCommercial;
                LECommercialA.Properties.DataSource = myObjectCommercial;

                LECommercialDE.Properties.DisplayMember = "mNomCommercial";
                LECommercialA.Properties.DisplayMember = "mNomCommercial";

                chkCmbMultCom.Properties.DataSource = myObjectCommercial;
                FillMultiCheckComboCommercial(chkCmbMultCom, myObjectCommercial);

                //LEFamilleDE.Properties.DisplayMember = "mCT_Intitule";
                //LEFamilleA.Properties.DisplayMember = "mCT_Intitule";

                //Choisir les premières valeurs
                LEFamilleDE.ItemIndex = 0;
                LEFamilleA.ItemIndex = 0;


                //Remplir les combo mois 
                FillComboMois(LMoisDeb);
                FillComboMois(LMoisFin);

                //Choisir les premières valeurs
                LMoisDeb.ItemIndex = 0;
                LMoisFin.ItemIndex = 0;

                //Remplir Combo Annee
                FillMultiCheckComboAnnee(ChkCmbMulAn);

                //Combo type de document
                FillComboTypeDoc(CmbTypeDoc);
                CmbTypeDoc.ItemIndex = 0;

                //Combo axe d'analyse
                FillAxeAnalyse(LEAxeAnalyse);
                LEAxeAnalyse.ItemIndex = 0;

                //Trimestre

                FillComboTrimestre(CmbTrimestre);
                CmbTrimestre.ItemIndex = 0;

                //Semestre
                FillComboSemestre(CmbSemestre);
                CmbSemestre.ItemIndex = 0;


                #endregion

                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAnalyseComp ->FenAnalyseComp_Load -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }


        private void CleanGrid()
        {
            try
            {
                //Nettoyer la grid=================================================
                List<ComRevendeur> Lempty = new List<ComRevendeur>();

                chartControlAnalCom.Series.Clear();

                gridControlData.DataSource = Lempty;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAnalyseComp ->CleanGrid -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }

        private void FillComboFamilleByChoixMarque()
        {
            try
            {
                //Charger la liste des famille pour critère par rapport a marque par defaut
                List<ComFamille> LfamilleDefautMarque = new List<ComFamille>();
                List<string> Lidpays = new List<string>();
                List<string> LFamCentrToFiltre = new List<string>();

                foreach (var elt in myObjectFilialeChoisies)
                {
                    Lidpays.Add(elt.mId.ToString());
                }

                LFamCentrToFiltre.Add(LEMarque.Text);

                LfamilleDefautMarque = daoReport.GetListByListIdPaysFamCentrFam(Lidpays, myObjectFamille, LFamCentrToFiltre);

                LEFamilleDE.Properties.DataSource = LfamilleDefautMarque;
                LEFamilleA.Properties.DataSource = LfamilleDefautMarque;

                chkCmbMultFam.Properties.DataSource = LfamilleDefautMarque;
                FillMultiCheckComboFamille(chkCmbMultFam, LfamilleDefautMarque);

                LEFamilleDE.Properties.DisplayMember = "mFa_CodeFamille";
                LEFamilleA.Properties.DisplayMember = "mFa_CodeFamille";
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAnalyseComp ->FillComboFamilleByChoixMarque -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }


       

        public void FillMultiCheckComboAnnee(CheckedComboBoxEdit cmb)
        {
            try
            {
               
                    var MySelectBases = new DataTable();

                    MySelectBases.Columns.Add("mAnnee");

                    for (int i=2000; i<=2100;i++)
                    {
                        MySelectBases.Rows.Add(i);

                        cmb.Properties.DataSource = MySelectBases;

                        cmb.Properties.ValueMember = "mAnnee";

                        cmb.Properties.DisplayMember = "mAnnee";

                    }

                    
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillMultiCheckComboFamille -> TypeErreur: " + ex.Message;
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
                        if (item.mNomCommercial != string.Empty && item.mPrenomCommercial != string.Empty)
                        {
                            MySelectBases.Rows.Add(item.mNomCommercial, item.mPrenomCommercial, item.mNomCommercial + " " + item.mPrenomCommercial, item.mPays);

                        }

                        if (item.mNomCommercial == string.Empty && item.mPrenomCommercial != string.Empty)
                        {
                            MySelectBases.Rows.Add(item.mNomCommercial, item.mPrenomCommercial, item.mPrenomCommercial, item.mPays);

                        }

                        if (item.mNomCommercial != string.Empty && item.mPrenomCommercial == string.Empty)
                        {
                            MySelectBases.Rows.Add(item.mNomCommercial, item.mNomCommercial, item.mNomCommercial, item.mPays);

                        }

                        cmb.Properties.DataSource = MySelectBases;

                        cmb.Properties.ValueMember = "mPrenomCommercial";

                        cmb.Properties.DisplayMember = "mNomPrenomCommercial";



                        // cmb.EditValue = item.mPrenomCommercial;


                    }



                }
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillMultiCheckComboCommercial -> TypeErreur: " + ex.Message;
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
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillMultiCheckComboFamille -> TypeErreur: " + ex.Message;
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
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillMultiCheckComboRevendeur -> TypeErreur: " + ex.Message;
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
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillMultiCheckComboFamilleCentral -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }

        

        private void chkTousCommerciaux_CheckStateChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousCommerciauxAxe.CheckState == CheckState.Checked)
                {
                    LECommercial.Enabled = false;

                }
                else
                {
                    LECommercial.Enabled = true;

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAnalyseComp ->chkTousCommerciaux_CheckStateChanged -> TypeErreur: " + ex.Message;
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
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousFamille_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

            }
        }

        private void SMulFamille_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulFamille.CheckState == CheckState.Checked)
                {
                    chkCmbMultFam.Visible = true;

                    chkTousFamille.Visible = false;

                    LEFamilleA.Visible = false;

                    LEFamilleDE.Visible = false;

                    lblFamilleDe.Visible = false;
                    lblFamilleA.Visible = false;

                }
                else
                {
                    chkCmbMultFam.Visible = false;

                    chkTousFamille.Visible = true;

                    LEFamilleA.Visible = true;

                    LEFamilleDE.Visible = true;

                    lblFamilleDe.Visible = true;
                    lblFamilleA.Visible = true;
                }
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkMulFamille_CheckedChanged -> TypeErreur: " + ex.Message; ;
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
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousRevendeur_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

            }
        }

        public void FillComboMois(LookUpEdit cmb)
        {
            try
            {
                    var MySelectBases = new DataTable();

                    MySelectBases.Columns.Add("mNomMois");
                    MySelectBases.Columns.Add("mId");

                 MySelectBases.Rows.Add("Janvier", 1);
                 MySelectBases.Rows.Add("Février", 2);
                 MySelectBases.Rows.Add("Mars", 3);
                 MySelectBases.Rows.Add("Avril", 4);
                 MySelectBases.Rows.Add("Mai", 5);
                 MySelectBases.Rows.Add("Juin", 6);
                 MySelectBases.Rows.Add("Juillet",7);
                 MySelectBases.Rows.Add("Août", 8);
                 MySelectBases.Rows.Add("Septembre", 9);
                 MySelectBases.Rows.Add("Octobre", 10);
                 MySelectBases.Rows.Add("Novembre", 11);
                MySelectBases.Rows.Add("Décembre", 12);
                
                        cmb.Properties.DataSource = MySelectBases;

                        cmb.Properties.ValueMember = "mId";

                        cmb.Properties.DisplayMember = "mNomMois";
                

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillComboFamille -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }

        public void FillComboTrimestre(LookUpEdit cmb)
        {
            try
            {
                var MySelectBases = new DataTable();

                MySelectBases.Columns.Add("mNomTrimestre");
                MySelectBases.Columns.Add("mId");

                MySelectBases.Rows.Add("T1 (Janvier-Mars)", 1);
                MySelectBases.Rows.Add("T2 (Avril-Juin)", 2);
                MySelectBases.Rows.Add("T3 (Juillet-Septembre)", 3);
                MySelectBases.Rows.Add("T4 (Octobre-Decembre)", 4);
          

                cmb.Properties.DataSource = MySelectBases;

                cmb.Properties.ValueMember = "mId";

                cmb.Properties.DisplayMember = "mNomTrimestre";


            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillComboFamille -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }

        public void FillComboSemestre(LookUpEdit cmb)
        {
            try
            {
                var MySelectBases = new DataTable();

                MySelectBases.Columns.Add("mNomSemestre");
                MySelectBases.Columns.Add("mId");

                MySelectBases.Rows.Add("S1 (Janvier-Juin)", 1);
                MySelectBases.Rows.Add("S2 (Juillet-Decembre)", 2);
             
                cmb.Properties.DataSource = MySelectBases;

                cmb.Properties.ValueMember = "mId";

                cmb.Properties.DisplayMember = "mNomSemestre";


            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillComboSemestre -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }


        public void FillComboTypeDoc(LookUpEdit cmb)
        {
            try
            {
                var MySelectBases = new DataTable();

                MySelectBases.Columns.Add("mNomDocument");
                MySelectBases.Columns.Add("mId");

                MySelectBases.Rows.Add("Facture (FA et FC)", 8);
                MySelectBases.Rows.Add("DE", 0);
                MySelectBases.Rows.Add("BC", 1);
                MySelectBases.Rows.Add("PL", 2);
                MySelectBases.Rows.Add("BL", 3);
                MySelectBases.Rows.Add("BR", 4);
                MySelectBases.Rows.Add("BA", 5);
            
              

                cmb.Properties.DataSource = MySelectBases;

                cmb.Properties.ValueMember = "mId";

                cmb.Properties.DisplayMember = "mNomDocument";


            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillComboFamille -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }

        public void FillAxeAnalyse(LookUpEdit cmb)
        {
            try
            {
                var MySelectBases = new DataTable();

                MySelectBases.Columns.Add("mNomAxe");
                MySelectBases.Columns.Add("mId");
                
                MySelectBases.Rows.Add("Commercial", 1);
                MySelectBases.Rows.Add("Marque", 2);
                MySelectBases.Rows.Add("Famille", 3);
                MySelectBases.Rows.Add("Revendeur", 4);
                //Multiples

                MySelectBases.Rows.Add("Multi Commerciaux", 5);
                MySelectBases.Rows.Add("Multi Marques", 6);
                MySelectBases.Rows.Add("Multi Familles", 7);
                MySelectBases.Rows.Add("Multi Revendeurs", 8);

                cmb.Properties.DataSource = MySelectBases;

                cmb.Properties.ValueMember = "mId";

                cmb.Properties.DisplayMember = "mNomAxe";


            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp ->FillAxeAnalyse -> TypeErreur: " + ex.Message;
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

                    lblRevendeurDe.Visible = false;
                    lblRevendeurA.Visible = false;

                }
                else
                {
                    chkCmbMultRev.Visible = false;

                    chkTousRevendeur.Visible = true;

                    LERevendeurA.Visible = true;

                    LERevendeurDE.Visible = true;

                    lblRevendeurDe.Visible = true;
                    lblRevendeurA.Visible = true;
                }
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> SMulRevendeur_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulAnnee.CheckState == CheckState.Checked)
                {
                    ChkCmbMulAn.Visible = true;
                    sNumAnneeDeb.Visible = false;
                    sNumAnneeFin.Visible = false;
                    lblAnneeA.Visible = false;
                    lblAnneeDe.Visible = false;
                    
                }
                else
                {
                    ChkCmbMulAn.Visible = false;
                    sNumAnneeDeb.Visible = true;
                    sNumAnneeFin.Visible = true;
                    lblAnneeA.Visible = true;
                    lblAnneeDe.Visible = true;

                }
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> checkEdit1_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void ChkCmbMulAn_Closed(object sender, DevExpress.XtraEditors.Controls.ClosedEventArgs e)
        {
            try
            {
                ListeAnnees = new List<string>();
                string anneesComp = string.Empty;

             if(ListeAnnees!=null)   ListeAnnees.Clear();

                anneesComp = ChkCmbMulAn.Text;

                if(anneesComp!=string.Empty)
                {
                    var tab = anneesComp.Split(',');


                    for(int i=0;i<tab.Length;i++)
                    {
                        var an = tab[i].ToString().Trim();

                            ListeAnnees.Add(an);
                    }

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> ChkCmbMulAn_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }


        public List<string> GetListeAnneesMulticheck()
        {
            List<string> LRetour = new List<string>();
            try
            {
                string anneesComp = string.Empty;
                
                anneesComp = ChkCmbMulAn.Text;

                if (anneesComp != string.Empty)
                {
                    var tab = anneesComp.Split(',');

                    for (int i = 0; i < tab.Length; i++)
                    {
                        LRetour.Add(tab[i].ToString().Trim());
                    }

                }

                return LRetour;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> GetListeAnneesMulticheck -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return null;
            }
        }

        
        private void LECommercial_Closed(object sender, DevExpress.XtraEditors.Controls.ClosedEventArgs e)
        {
            try
            {
                DAONomCommercial = LECommercial.Text;

                DAOPrenomCommercial = LECommercial.EditValue.ToString();

                CleanGrid();
                timer1.Enabled = true;

            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LECommercial_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        public List<string> GetListeAnnees(int AnDeb,int AnFin)
        {
            List<string> AnneesComp = new List<string>();
            try
            {
                for(int i=AnDeb; i<=AnFin; i++)
                {
                    AnneesComp.Add(i.ToString());
                }
                
                return AnneesComp;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> GetListeAnnees -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return null;
            }
        }

        

        private void sBtnApercu_Click(object sender, EventArgs e)
        {
            string JourMoisDeb = string.Empty;
            string JourMoisFin = string.Empty;

            bool IsFinFevrierDeb = false;
            bool IsFinFevrierFin = false;

            string JourMoisDebBIX = string.Empty;
            string JourMoisFinBIX = string.Empty;

            List<ComRevendeur> LResult = new List<ComRevendeur>();

            List<ComDataComp> LData = new List<ComDataComp>();

            try
            {
                //Nom et prenom du commercial a étudier

                if(IsCommercialAxe)
                {
                    DAONomCommercial = LECommercial.Text;

                    DAOPrenomCommercial = LECommercial.EditValue.ToString();
                }


                //Nom  du FAMILLE a étudier(On ne fera pas ça avec les marques)

                if (IsFamilleAxe)
                {
                    DAOFamille = LEFamille.Text;
                    
                }

                //Nom 

                if(IsRevendeurAxe)
                {
                    DAORevendeur = LERevendeur.Text;
                }
                


                if (ListeAnnees!=null)
                {
                    if (ListeAnnees.Count > 0) ListeAnnees.Clear();
                }


                //si date deb > date fin pas la peine

                if(!SMulAnnee.Checked)
                {

                    if (sNumAnneeDeb.Value > sNumAnneeFin.Value)
                    {
                        MessageBox.Show("L'année de début d'analyse est supérieure à l'année de fin!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                }

                //Aucune année choisie

                if(SMulAnnee.Checked)
                {
                    if (ChkCmbMulAn.Text==string.Empty)
                    {
                        MessageBox.Show("Veuillez choisir une année d'analyse!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                //Liste des années
                if (SMulAnnee.Checked)
                {
                    ListeAnnees = GetListeAnneesMulticheck();
                }
              else
                {
                    ListeAnnees = GetListeAnnees(Int32.Parse(sNumAnneeDeb.EditValue.ToString()), Int32.Parse(sNumAnneeFin.EditValue.ToString()));
                }

              if(!IsTrimetre && !IsSemetre && !IsPeriodeAdefinir && !IsAnnuel &&!IsMensuel)
                {
                    MessageBox.Show("Veuillez définir une période d'analyse!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

              //cas des multiples ,on ne doit pas gérer le annuel par mois

                if(IsCommercialAxeMultiple|| IsMarqueAxeMultiple || IsFamilleAxeMultiple || IsRevendeurAxeMultiple)
                {
                    if (!IsTrimetre && !IsSemetre && !IsPeriodeAdefinir && !IsAnnuel )
                    {
                        MessageBox.Show("Veuillez définir une période d'analyse!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }


              
                //Periode Début et Fin

                //Mensuel
              
                if (IsMensuel)
                {
                    //JAnv -Juin
                    JourMoisDeb = "01-01";
                    JourMoisFin = "31-12";

                }
                
                //trimestre
                if (IsTrimetre)
                {
                    int choix = Int32.Parse(CmbTrimestre.EditValue.ToString());

                    switch (choix)
                    {
                        case 1://JAnv -Mars
                            JourMoisDeb = "01-01";
                            JourMoisFin = "31-03";
                            break;
                        case 2://Avril -Juin
                            JourMoisDeb = "01-04";
                            JourMoisFin = "30-06";
                            break;

                        case 3://Juillet -Sept
                            JourMoisDeb = "01-07";
                            JourMoisFin = "30-09";
                            break;

                        case 4://Oct -Décembre
                            JourMoisDeb = "01-10";
                            JourMoisFin = "31-12";
                            break;
                            

                    }
                   
                }

                //semestre
                if (IsSemetre)
                {
                  
                    int choixsem = Int32.Parse(CmbSemestre.EditValue.ToString());
                    switch (choixsem)
                    {
                        case 1://JAnv -Juin
                            JourMoisDeb = "01-01";
                            JourMoisFin = "30-06";
                            break;
                        case 2://Juillet -Décembre
                            JourMoisDeb = "01-07";
                            JourMoisFin = "31-12";
                            break;
                    }
                }

                //Annuel
                if (IsAnnuel)
                {
                       //JAnv -Juin
                            JourMoisDeb = "01-01";
                            JourMoisFin = "31-12";
                       
                }

              
                //Cas période définie

                if (IsPeriodeAdefinir)
                {
                    JourMoisDeb = sJourDeb.EditValue.ToString() + "-" + LMoisDeb.EditValue.ToString();

                    int jdeb = Int32.Parse(sJourDeb.EditValue.ToString());
                    int moisdeb= Int32.Parse(LMoisDeb.EditValue.ToString());

                    if(jdeb==28 && moisdeb==2)
                    {
                        IsFinFevrierDeb = true;
                        JourMoisDebBIX = "29-02";
                    }
               
                    JourMoisFin = sJourFin.EditValue.ToString() + "-" + LMoisFin.EditValue.ToString();

                    int jfin = Int32.Parse(sJourFin.EditValue.ToString());
                    int moisfin = Int32.Parse(LMoisFin.EditValue.ToString());

                    if (jfin == 28 && moisfin == 2)
                    {
                        IsFinFevrierFin = true;
                        JourMoisFinBIX = "29-02";
                    }

                    //Voir si mois deb > mois Fin

                     if(moisdeb>moisfin)
                    {
                        MessageBox.Show("Votre période d'analyse n'est pas cohérante!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                     //cas ou les mois sont pareils ,voir les jours

                     if(moisdeb==moisfin)
                    {
                        if(jdeb>jfin)
                        {
                            MessageBox.Show("Votre période d'analyse n'est pas cohérante!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                    }



                }

                if (JourMoisDeb == string.Empty || JourMoisFin == string.Empty)
                {
                    MessageBox.Show("Veuillez définir une période d'analyse!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                //Type Document

                string TypDocId = CmbTypeDoc.EditValue.ToString();

                if (!splashScreenManager1.IsSplashFormVisible) splashScreenManager1.ShowWaitForm();

                //Affichage courbe commercial+++++++++++++++++++++++++++++++++++++++++++++++
                if (IsCommercialAxe)
                {
                    //Cas BUNDLE(facture)

                    if(Int16.Parse(TypDocId)==8)
                    {
                        LResult = daoReport.GetComparativeCommercialBUNDLE(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousCommerciauxAxe.Checked, DAONomCommercial, DAOPrenomCommercial, SMulMarque.Checked, chkTousMarque.Checked, LEMarqueDE.Text, LEMarqueA.Text, ListMultiMarqueCritere, chkConfigFamilles.Checked, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDECritere, ListNomLEFamilleACritere, ListMultiFamilleCritere, TypDocId, SMulRevendeur.Checked, chkTousRevendeur.Checked, LERevendeurDE.Text, LERevendeurA.Text, ListMultiRevendeurCritere, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);

                    }
                    else
                    {
                        LResult = daoReport.GetComparativeCommercial(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousCommerciauxAxe.Checked, DAONomCommercial, DAOPrenomCommercial, SMulMarque.Checked, chkTousMarque.Checked, LEMarqueDE.Text, LEMarqueA.Text, ListMultiMarqueCritere, chkConfigFamilles.Checked, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDECritere, ListNomLEFamilleACritere, ListMultiFamilleCritere, TypDocId, SMulRevendeur.Checked, chkTousRevendeur.Checked, LERevendeurDE.Text, LERevendeurA.Text, ListMultiRevendeurCritere, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);

                    }
                    
                    // if(!SMulFamille.Checked && chkTousFamille.Checked && !SMulRevendeur.Checked && chkTousRevendeur.Checked)
                    chartControlAnalCom.Series.Clear();
                    //cas du mensuel ,on groupe par mois

                    if (IsMensuel)
                    {
                        if (LResult.Count > 0)
                        {
                            #region chartControlAnalCom
                            
                            foreach (var item in ListeAnnees)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                Series ser = new Series(item.ToString(), ViewType.Bar);

                                double SommeMois = 0;

                                for(int i=1;i<13;i++)
                                {
                                    SommeMois = 0;
                                    Ltemp = LResult.Where(c => c.mDatePiece.Month == i && c.mDatePiece.Year == Int32.Parse(item.ToString())).ToList();

                                    foreach (var obj in Ltemp)
                                    {
                                        SommeMois += obj.mMontantHT;
                                    }

                                    ComDataComp CData = new ComDataComp();
                                    

                                    if (Ltemp != null )
                                    {
                                        //Series ser = new Series(item.ToString(), ViewType.Bar);
                                        string NomMois = string.Empty;

                                        NomMois = GetNomById(i);

                                        CData.mAnnee = Int32.Parse(item.ToString());
                                        CData.mMois = NomMois;
                                        CData.mMontant = SommeMois;

                                        LData.Add(CData);


                                        if (chartControlAnalCom.Series.Count > 0)
                                        {
                                            //if (!chartControlAnalCom.Series[0].Name.Equals(ser.Name))
                                            //{
                                            //foreach (var obj in Ltemp)
                                            //{
                                            //     ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeMois));
                                            
                                            ser.Points.Add(new SeriesPoint(NomMois, SommeMois));

                                            // }

                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            ////tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            //Annotations dans la legende
                                            //AxisLabel TextPattern = "{}{V:#,##0}"

                                          //  ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            ////ser.CrosshairLabelPattern = "{V:n0}";
                                            ////  }
                                        }
                                        else
                                        {
                                            //foreach (var obj in Ltemp)
                                            //{
                                            //  ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeMois));
                                            // }
                                            ser.Points.Add(new SeriesPoint(NomMois, SommeMois));

                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            ////tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            //Annotations dans la legende
                                            //AxisLabel TextPattern = "{}{V:#,##0}"

                                           // ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            //ser.CrosshairLabelPattern = "{V:n0}";
                                        }

                                       

                                    }

                                }

                              

                                //tester que la serie existe ou pas dans le graphe
                                chartControlAnalCom.Series.Add(ser);
                                ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";


                            }

                            if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            //grid DATA
                            gridControlData.DataSource = LData;


                            //Masquer colonnes
                            //Masquer colonnes
                            gridView1.Columns["mAnnee"].Visible = true;
                            gridView1.Columns["mMontant"].Visible = true;
                            gridView1.Columns["mMois"].Visible = true;
                            gridView1.Columns["mCommercial"].Visible = false;
                            gridView1.Columns["mMarque"].Visible = false;
                            gridView1.Columns["mFamille"].Visible = false;
                            gridView1.Columns["mRevendeur"].Visible = false;

                            #endregion

                        }

                    }
                    else
                    {
                        if (LResult.Count > 0)
                        {
                            #region chartControlAnalCom

                            chartControlAnalCom.Series.Clear();

                            foreach (var item in ListeAnnees)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                double SommeTotal = 0;

                                Ltemp = LResult.Where(c => c.mDatePiece.Year == Int32.Parse(item.ToString())).ToList();

                                foreach (var obj in Ltemp)
                                {
                                    SommeTotal += obj.mMontantHT;
                                }

                                ComDataComp CData = new ComDataComp();

                                CData.mAnnee = Int32.Parse(item.ToString());
                                CData.mMois = string.Empty;
                                CData.mMontant = SommeTotal;

                                LData.Add(CData);

                                if (Ltemp != null )
                                {
                                    Series ser = new Series(item.ToString(), ViewType.Bar);
                                    
                                    if (chartControlAnalCom.Series.Count > 0)
                                    {
                                        //if (!chartControlAnalCom.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in Ltemp)
                                        //{
                                        if(chkTousCommerciauxAxe.Checked)
                                        {
                                            ser.Points.Add(new SeriesPoint("Tous les Commerciaux", SommeTotal));
                                        }
                                        else
                                        {
                                            ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeTotal));
                                        }

                                    
                                        // }

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        chartControlAnalCom.Series.Add(ser);
                                        //Annotations dans la legende
                                        //AxisLabel TextPattern = "{}{V:#,##0}"


                                        ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in Ltemp)
                                        //{
                                        if (chkTousCommerciauxAxe.Checked)
                                        {
                                            ser.Points.Add(new SeriesPoint("Tous les Commerciaux", SommeTotal));
                                        }
                                        else
                                        {
                                            ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeTotal));
                                        }

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        chartControlAnalCom.Series.Add(ser);
                                        //Annotations dans la legende
                                        //AxisLabel TextPattern = "{}{V:#,##0}"
                                        
                                        ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                    }

                                }

                            }

                            if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            gridControlData.DataSource = LData;
                            
                            //Masquer colonnes
                            gridView1.Columns["mAnnee"].Visible = true;
                            gridView1.Columns["mMontant"].Visible = true;
                            gridView1.Columns["mMois"].Visible = false;
                            gridView1.Columns["mCommercial"].Visible = false;
                            gridView1.Columns["mMarque"].Visible = false;
                            gridView1.Columns["mFamille"].Visible = false;
                            gridView1.Columns["mRevendeur"].Visible = false;

                            #endregion

                        }

                    }
                    
                }

                //Affichage courbe Marque+++++++++++++++++++++++++++++++++++++++++++++++
                if (IsMarqueAxe)
                {
                    LResult = daoReport.GetComparativeMarque(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees,chkTousMarqueAxe.Checked,LEMarque.Text, chkTousCom.Checked,SMulCom.Checked, ListNomMultiCommerciauxCritere, ListPrenomMultiCommerciauxCritere, LECommercialDE.Text, LECommercialA.Text,SMulFamille.Checked, chkTousFamille.Checked, LEFamilleDE.Text,LEFamilleA.Text, ListMultiFamilleCritere,  TypDocId, SMulRevendeur.Checked,chkTousRevendeur.Checked, LERevendeurDE.Text, LERevendeurA.Text, ListMultiRevendeurCritere, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);
           
            // if(!SMulFamille.Checked && chkTousFamille.Checked && !SMulRevendeur.Checked && chkTousRevendeur.Checked)
            chartControlAnalCom.Series.Clear();
                    //cas du mensuel ,on groupe par mois

                    if (IsMensuel)
                    {
                        if (LResult.Count > 0)
                        {
                            #region chartControlAnalCom

                            foreach (var item in ListeAnnees)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                Series ser = new Series(item.ToString(), ViewType.Bar);

                                double SommeMois = 0;

                                for (int i = 1; i < 13; i++)
                                {
                                    SommeMois = 0;
                                    Ltemp = LResult.Where(c => c.mDatePiece.Month == i && c.mDatePiece.Year == Int32.Parse(item.ToString())).ToList();

                                    foreach (var obj in Ltemp)
                                    {
                                        SommeMois += obj.mMontantHT;
                                    }

                                    ComDataComp CData = new ComDataComp();

                                    if (Ltemp != null)
                                    {
                                        //Series ser = new Series(item.ToString(), ViewType.Bar);
                                        string NomMois = string.Empty;

                                        NomMois = GetNomById(i);

                                        CData.mAnnee = Int32.Parse(item.ToString());
                                        CData.mMois = NomMois;
                                        CData.mMontant = SommeMois;

                                        LData.Add(CData);

                                        if (chartControlAnalCom.Series.Count > 0)
                                        {
                                            //if (!chartControlAnalCom.Series[0].Name.Equals(ser.Name))
                                            //{
                                            //foreach (var obj in Ltemp)
                                            //{
                                            //     ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeMois));

                                            ser.Points.Add(new SeriesPoint(NomMois, SommeMois));

                                            // }

                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            ////tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            //Annotations dans la legende
                                            //AxisLabel TextPattern = "{}{V:#,##0}"

                                            //  ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            ////ser.CrosshairLabelPattern = "{V:n0}";
                                            ////  }
                                        }
                                        else
                                        {
                                            //foreach (var obj in Ltemp)
                                            //{
                                            //  ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeMois));
                                            // }
                                            ser.Points.Add(new SeriesPoint(NomMois, SommeMois));

                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            ////tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            //Annotations dans la legende
                                            //AxisLabel TextPattern = "{}{V:#,##0}"

                                            // ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            //ser.CrosshairLabelPattern = "{V:n0}";
                                        }

                                    }

                                }
                                
                                //tester que la serie existe ou pas dans le graphe
                                chartControlAnalCom.Series.Add(ser);
                                ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";
                                
                            }

                            if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            gridControlData.DataSource = LData;

                            //Masquer colonnes
                            gridView1.Columns["mAnnee"].Visible = true;
                            gridView1.Columns["mMontant"].Visible = true;
                            gridView1.Columns["mMois"].Visible = true;
                            gridView1.Columns["mCommercial"].Visible = false;
                            gridView1.Columns["mMarque"].Visible = false;
                            gridView1.Columns["mFamille"].Visible = false;
                            gridView1.Columns["mRevendeur"].Visible = false;
                            #endregion

                        }

                    }
                    else
                    {
                        if (LResult.Count > 0)
                        {
                            #region chartControlAnalCom

                            chartControlAnalCom.Series.Clear();

                            foreach (var item in ListeAnnees)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                double SommeTotal = 0;

                                Ltemp = LResult.Where(c => c.mDatePiece.Year == Int32.Parse(item.ToString())).ToList();

                                foreach (var obj in Ltemp)
                                {
                                    SommeTotal += obj.mMontantHT;
                                }

                                ComDataComp CData = new ComDataComp();

                                CData.mAnnee = Int32.Parse(item.ToString());
                                CData.mMois = string.Empty;
                                CData.mMontant = SommeTotal;

                                LData.Add(CData);

                                if (Ltemp != null)
                                {
                                    Series ser = new Series(item.ToString(), ViewType.Bar);

                                    if (chartControlAnalCom.Series.Count > 0)
                                    {
                                        //if (!chartControlAnalCom.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in Ltemp)
                                        //{
                                        if (chkTousMarqueAxe.Checked)
                                        {
                                            ser.Points.Add(new SeriesPoint("Toutes les Marques", SommeTotal));
                                        }
                                        else
                                        {
                                            ser.Points.Add(new SeriesPoint(LEMarque.Text, SommeTotal));
                                        }
                                        
                                        // }

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        chartControlAnalCom.Series.Add(ser);
                                        //Annotations dans la legende
                                        //AxisLabel TextPattern = "{}{V:#,##0}"

                                        ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in Ltemp)
                                        //{
                                        if (chkTousMarqueAxe.Checked)
                                        {
                                            ser.Points.Add(new SeriesPoint("Toutes les Marques", SommeTotal));
                                        }
                                        else
                                        {
                                            ser.Points.Add(new SeriesPoint(LEMarque.Text, SommeTotal));
                                        }

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        chartControlAnalCom.Series.Add(ser);
                                        //Annotations dans la legende
                                        //AxisLabel TextPattern = "{}{V:#,##0}"

                                        ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                    }

                                }

                            }
                            if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            gridControlData.DataSource = LData;
                            //Masquer colonnes
                            gridView1.Columns["mAnnee"].Visible = true;
                            gridView1.Columns["mMontant"].Visible = true;
                            gridView1.Columns["mMois"].Visible = false;
                            gridView1.Columns["mCommercial"].Visible = false;
                            gridView1.Columns["mMarque"].Visible = false;
                            gridView1.Columns["mFamille"].Visible = false;
                            gridView1.Columns["mRevendeur"].Visible = false;
                            #endregion

                        }

                    }

                }


                //Affichage courbe Famille+++++++++++++++++++++++++++++++++++++++++++++++
                if (IsFamilleAxe)
                {
                    LResult = daoReport.GetComparativeFamille(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousCom.Checked,SMulCom.Checked,LECommercialDE.Text,LECommercialA.Text, ListNomMultiCommerciauxCritere, ListPrenomMultiCommerciauxCritere,chkTousFamilleAxe.Checked,DAOFamille, TypDocId, SMulRevendeur.Checked, chkTousRevendeur.Checked, LERevendeurDE.Text, LERevendeurA.Text, ListMultiRevendeurCritere, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);

                    // if(!SMulFamille.Checked && chkTousFamille.Checked && !SMulRevendeur.Checked && chkTousRevendeur.Checked)
                    chartControlAnalCom.Series.Clear();
                    //cas du mensuel ,on groupe par mois

                    if (IsMensuel)
                    {
                        if (LResult.Count > 0)
                        {
                            #region chartControlAnalCom

                            foreach (var item in ListeAnnees)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                Series ser = new Series(item.ToString(), ViewType.Bar);

                                double SommeMois = 0;

                                for (int i = 1; i < 13; i++)
                                {
                                    SommeMois = 0;
                                    Ltemp = LResult.Where(c => c.mDatePiece.Month == i && c.mDatePiece.Year == Int32.Parse(item.ToString())).ToList();

                                    foreach (var obj in Ltemp)
                                    {
                                        SommeMois += obj.mMontantHT;
                                    }


                                    ComDataComp CData = new ComDataComp();

                                    if (Ltemp != null)
                                    {
                                        //Series ser = new Series(item.ToString(), ViewType.Bar);
                                        string NomMois = string.Empty;

                                        NomMois = GetNomById(i);


                                        CData.mAnnee = Int32.Parse(item.ToString());
                                        CData.mMois = NomMois;
                                        CData.mMontant = SommeMois;

                                        LData.Add(CData);

                                        if (chartControlAnalCom.Series.Count > 0)
                                        {
                                            //if (!chartControlAnalCom.Series[0].Name.Equals(ser.Name))
                                            //{
                                            //foreach (var obj in Ltemp)
                                            //{
                                            //     ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeMois));

                                            ser.Points.Add(new SeriesPoint(NomMois, SommeMois));

                                            // }

                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            ////tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            //Annotations dans la legende
                                            //AxisLabel TextPattern = "{}{V:#,##0}"

                                            //  ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            ////ser.CrosshairLabelPattern = "{V:n0}";
                                            ////  }
                                        }
                                        else
                                        {
                                            //foreach (var obj in Ltemp)
                                            //{
                                            //  ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeMois));
                                            // }
                                            ser.Points.Add(new SeriesPoint(NomMois, SommeMois));

                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            ////tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            //Annotations dans la legende
                                            //AxisLabel TextPattern = "{}{V:#,##0}"

                                            // ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            //ser.CrosshairLabelPattern = "{V:n0}";
                                        }

                                    }

                                }
                                
                                //tester que la serie existe ou pas dans le graphe
                                chartControlAnalCom.Series.Add(ser);
                                ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";
                                
                            }
                            if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            gridControlData.DataSource = LData;
                            //Masquer colonnes
                            gridView1.Columns["mAnnee"].Visible = true;
                            gridView1.Columns["mMontant"].Visible = true;
                            gridView1.Columns["mMois"].Visible = true;
                            gridView1.Columns["mCommercial"].Visible = false;
                            gridView1.Columns["mMarque"].Visible = false;
                            gridView1.Columns["mFamille"].Visible = false;
                            gridView1.Columns["mRevendeur"].Visible = false;
                            #endregion

                        }

                    }
                    else
                    {
                        if (LResult.Count > 0)
                        {
                            #region chartControlAnalCom

                            chartControlAnalCom.Series.Clear();

                            foreach (var item in ListeAnnees)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                double SommeTotal = 0;

                                Ltemp = LResult.Where(c => c.mDatePiece.Year == Int32.Parse(item.ToString())).ToList();

                                foreach (var obj in Ltemp)
                                {
                                    SommeTotal += obj.mMontantHT;
                                }

                                ComDataComp CData = new ComDataComp();

                                CData.mAnnee = Int32.Parse(item.ToString());
                                CData.mMois = string.Empty;
                                CData.mMontant = SommeTotal;

                                LData.Add(CData);

                                if (Ltemp != null)
                                {
                                    Series ser = new Series(item.ToString(), ViewType.Bar);

                                    if (chartControlAnalCom.Series.Count > 0)
                                    {
                                        //if (!chartControlAnalCom.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in Ltemp)
                                        //{
                                        if (chkTousFamilleAxe.Checked)
                                        {
                                            ser.Points.Add(new SeriesPoint("Toutes les Familles", SommeTotal));
                                        }
                                        else
                                        {
                                            ser.Points.Add(new SeriesPoint(LEFamille.Text , SommeTotal));
                                        }


                                        // }

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        chartControlAnalCom.Series.Add(ser);
                                        //Annotations dans la legende
                                        //AxisLabel TextPattern = "{}{V:#,##0}"

                                        ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in Ltemp)
                                        //{
                                        if (chkTousFamilleAxe.Checked)
                                        {
                                            ser.Points.Add(new SeriesPoint("Toutes les Familles", SommeTotal));
                                        }
                                        else
                                        {
                                            ser.Points.Add(new SeriesPoint(LEFamille.Text, SommeTotal));
                                        }

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        chartControlAnalCom.Series.Add(ser);
                                        //Annotations dans la legende
                                        //AxisLabel TextPattern = "{}{V:#,##0}"

                                        ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                    }

                                }

                            }
                            if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            gridControlData.DataSource = LData;

                            //Masquer colonnes
                            gridView1.Columns["mAnnee"].Visible = true;
                            gridView1.Columns["mMontant"].Visible = true;
                            gridView1.Columns["mMois"].Visible = false;
                            gridView1.Columns["mCommercial"].Visible = false;
                            gridView1.Columns["mMarque"].Visible = false;
                            gridView1.Columns["mFamille"].Visible = false;
                            gridView1.Columns["mRevendeur"].Visible = false;

                            #endregion

                        }

                    }

                }

                //Affichage courbe Revendeur+++++++++++++++++++++++++++++++++++++++++++++++
                if (IsRevendeurAxe)
                {
                    
                    if (Int16.Parse(TypDocId) == 8)
                    {
                        LResult = daoReport.GetComparativeRevendeurBUNDLE(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousRevendeurAxe.Checked, DAORevendeur, chkTousCom.Checked, SMulCom.Checked, LECommercialDE.Text, LECommercialA.Text, ListNomMultiCommerciauxCritere, ListPrenomMultiCommerciauxCritere, SMulMarque.Checked, chkTousMarque.Checked, LEMarqueDE.Text, LEMarqueA.Text, ListMultiMarqueCritere, chkConfigFamilles.Checked, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDECritere, ListNomLEFamilleACritere, ListMultiFamilleCritere, TypDocId, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);

                    }
                    else
                    {
                        LResult = daoReport.GetComparativeRevendeur(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousRevendeurAxe.Checked, DAORevendeur, chkTousCom.Checked, SMulCom.Checked, LECommercialDE.Text, LECommercialA.Text, ListNomMultiCommerciauxCritere, ListPrenomMultiCommerciauxCritere, SMulMarque.Checked, chkTousMarque.Checked, LEMarqueDE.Text, LEMarqueA.Text, ListMultiMarqueCritere, chkConfigFamilles.Checked, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDECritere, ListNomLEFamilleACritere, ListMultiFamilleCritere, TypDocId, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);

                    }
                    
                    // if(!SMulFamille.Checked && chkTousFamille.Checked && !SMulRevendeur.Checked && chkTousRevendeur.Checked)
                    chartControlAnalCom.Series.Clear();
                    //cas du mensuel ,on groupe par mois

                    if (IsMensuel)
                    {
                        if (LResult.Count > 0)
                        {
                            #region chartControlAnalCom

                            foreach (var item in ListeAnnees)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                Series ser = new Series(item.ToString(), ViewType.Bar);

                                double SommeMois = 0;

                                for (int i = 1; i < 13; i++)
                                {
                                    SommeMois = 0;
                                    Ltemp = LResult.Where(c => c.mDatePiece.Month == i && c.mDatePiece.Year == Int32.Parse(item.ToString())).ToList();

                                    foreach (var obj in Ltemp)
                                    {
                                        SommeMois += obj.mMontantHT;
                                    }

                                    ComDataComp CData = new ComDataComp();


                                    if (Ltemp != null)
                                    {
                                        //Series ser = new Series(item.ToString(), ViewType.Bar);
                                        string NomMois = string.Empty;

                                        NomMois = GetNomById(i);

                                        CData.mAnnee = Int32.Parse(item.ToString());
                                        CData.mMois = NomMois;
                                        CData.mMontant = SommeMois;

                                        LData.Add(CData);

                                        if (chartControlAnalCom.Series.Count > 0)
                                        {
                                            //if (!chartControlAnalCom.Series[0].Name.Equals(ser.Name))
                                            //{
                                            //foreach (var obj in Ltemp)
                                            //{
                                            //     ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeMois));

                                            ser.Points.Add(new SeriesPoint(NomMois, SommeMois));

                                            // }

                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            ////tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            //Annotations dans la legende
                                            //AxisLabel TextPattern = "{}{V:#,##0}"

                                            //  ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            ////ser.CrosshairLabelPattern = "{V:n0}";
                                            ////  }
                                        }
                                        else
                                        {
                                            //foreach (var obj in Ltemp)
                                            //{
                                            //  ser.Points.Add(new SeriesPoint(LECommercial.Text + " " + LECommercial.EditValue.ToString(), SommeMois));
                                            // }
                                            ser.Points.Add(new SeriesPoint(NomMois, SommeMois));

                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            ////tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            //Annotations dans la legende
                                            //AxisLabel TextPattern = "{}{V:#,##0}"

                                            // ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            //ser.CrosshairLabelPattern = "{V:n0}";
                                        }



                                    }

                                }



                                //tester que la serie existe ou pas dans le graphe
                                chartControlAnalCom.Series.Add(ser);
                                ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";


                            }
                            if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            //grid DATA
                            gridControlData.DataSource = LData;

                            //Masquer colonnes
                            gridView1.Columns["mAnnee"].Visible = true;
                            gridView1.Columns["mMontant"].Visible = true;
                            gridView1.Columns["mMois"].Visible = true;
                            gridView1.Columns["mCommercial"].Visible = false;
                            gridView1.Columns["mMarque"].Visible = false;
                            gridView1.Columns["mFamille"].Visible = false;
                            gridView1.Columns["mRevendeur"].Visible = false;
                            #endregion

                        }

                    }
                    else
                    {
                        if (LResult.Count > 0)
                        {
                            #region chartControlAnalCom

                            chartControlAnalCom.Series.Clear();

                            foreach (var item in ListeAnnees)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                double SommeTotal = 0;

                                Ltemp = LResult.Where(c => c.mDatePiece.Year == Int32.Parse(item.ToString())).ToList();

                                foreach (var obj in Ltemp)
                                {
                                    SommeTotal += obj.mMontantHT;
                                }

                                ComDataComp CData = new ComDataComp();

                                CData.mAnnee = Int32.Parse(item.ToString());
                                CData.mMois = string.Empty;
                                CData.mMontant = SommeTotal;

                                LData.Add(CData);

                                if (Ltemp != null)
                                {
                                    Series ser = new Series(item.ToString(), ViewType.Bar);
                                    
                                    if (chartControlAnalCom.Series.Count > 0)
                                    {
                                        //if (!chartControlAnalCom.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in Ltemp)
                                        //{
                                        if (chkTousRevendeurAxe.Checked)
                                        {
                                            ser.Points.Add(new SeriesPoint("Tous les Revendeurs", SommeTotal));
                                        }
                                        else
                                        {
                                            ser.Points.Add(new SeriesPoint(LERevendeur.Text, SommeTotal));
                                        }


                                        // }

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        chartControlAnalCom.Series.Add(ser);
                                        //Annotations dans la legende
                                        //AxisLabel TextPattern = "{}{V:#,##0}"


                                        ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in Ltemp)
                                        //{
                                        if (chkTousRevendeurAxe.Checked)
                                        {
                                            ser.Points.Add(new SeriesPoint("Tous les Revendeurs", SommeTotal));
                                        }
                                        else
                                        {
                                            ser.Points.Add(new SeriesPoint(LERevendeur.Text , SommeTotal));
                                        }

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        chartControlAnalCom.Series.Add(ser);
                                        //Annotations dans la legende
                                        //AxisLabel TextPattern = "{}{V:#,##0}"

                                        ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                    }

                                }

                            }
                            if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            gridControlData.DataSource = LData;

                            //Masquer colonnes
                            gridView1.Columns["mAnnee"].Visible = true;
                            gridView1.Columns["mMontant"].Visible = true;
                            gridView1.Columns["mMois"].Visible = false;
                            gridView1.Columns["mCommercial"].Visible = false;
                            gridView1.Columns["mMarque"].Visible = false;
                            gridView1.Columns["mFamille"].Visible = false;
                            gridView1.Columns["mRevendeur"].Visible = false;

                            #endregion

                        }

                    }

                }

                //MULTIPLE COMPARAISON =============================================
                

                if(IsCommercialAxeMultiple)
                {

                    //Liste des commerciaux(Nom +" "+Prenoms)

                    if (Int16.Parse(TypDocId) == 8)
                    {
                        LResult = daoReport.GetComparativeMultiCommercialAxeBUNDLE(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousComMultAxe.Checked, SMulComMultAxe.Checked, LECommercialDEMultAxe.Text, LECommercialAMultAxe.Text, ListNomMultiCommerciauxMultipleAxe, ListPrenomMultiCommerciauxMultipleAxe, SMulMarque.Checked, chkTousMarque.Checked, LEMarqueDE.Text, LEMarqueA.Text, ListMultiMarqueCritere, chkConfigFamilles.Checked, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDECritere, ListNomLEFamilleACritere, ListMultiFamilleCritere, TypDocId, SMulRevendeur.Checked, chkTousRevendeur.Checked, LERevendeurDE.Text, LERevendeurA.Text, ListMultiRevendeurCritere, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);

                    }
                    else
                    {
                        LResult = daoReport.GetComparativeMultiCommercialAxe(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousComMultAxe.Checked, SMulComMultAxe.Checked, LECommercialDEMultAxe.Text, LECommercialAMultAxe.Text, ListNomMultiCommerciauxMultipleAxe, ListPrenomMultiCommerciauxMultipleAxe, SMulMarque.Checked, chkTousMarque.Checked, LEMarqueDE.Text, LEMarqueA.Text, ListMultiMarqueCritere, chkConfigFamilles.Checked, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDECritere, ListNomLEFamilleACritere, ListMultiFamilleCritere, TypDocId, SMulRevendeur.Checked, chkTousRevendeur.Checked, LERevendeurDE.Text, LERevendeurA.Text, ListMultiRevendeurCritere, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);

                    }



                    Dictionary<string, string> ListeCommerciaux = new Dictionary<string, string>();

                    if (LResult.Count > 0)
                    {

                        #region Liste des Commerciaux


                        foreach (var item in LResult)
                        {
                            if (ListeCommerciaux.Count == 0)
                            {
                                ListeCommerciaux.Add(item.mNomCommercial, item.mPrenomCommercial);
                            }
                            else
                            {
                                var a = ListeCommerciaux.ContainsKey(item.mNomCommercial);

                                if (a == false)
                                {
                                    var b = ListeCommerciaux.ContainsKey(item.mPrenomCommercial);

                                    if (b == false)
                                    {
                                        ListeCommerciaux.Add(item.mNomCommercial, item.mPrenomCommercial);
                                    }
                                }

                            }
                        }

                        #endregion

                    }
                    
                    // if(!SMulFamille.Checked && chkTousFamille.Checked && !SMulRevendeur.Checked && chkTousRevendeur.Checked)
                    chartControlAnalCom.Series.Clear();
                    
                        if (LResult.Count > 0)
                        {
                            #region chartControlAnalCom

                            chartControlAnalCom.Series.Clear();

                            foreach (var item in ListeAnnees)
                            {
                            Series ser = new Series(item.ToString(), ViewType.Bar);
                            foreach (var elt in ListeCommerciaux)
                                {
                                    List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                    double SommeTotal = 0;

                                    Ltemp = LResult.Where(c => c.mDatePiece.Year == Int32.Parse(item.ToString()) && c.mNomCommercial==elt.Key && c.mPrenomCommercial==elt.Value ).ToList();

                                    foreach (var obj in Ltemp)
                                    {
                                        SommeTotal += obj.mMontantHT;
                                    }

                                    ComDataComp CData = new ComDataComp();

                                    CData.mAnnee = Int32.Parse(item.ToString());
                                    CData.mMois = string.Empty;
                                    CData.mMontant = SommeTotal;
                                    CData.mCommercial = elt.Key + " " + elt.Value;
                                    CData.mMarque = string.Empty;
                                    CData.mFamille = string.Empty;
                                CData.mRevendeur = string.Empty;

                                    LData.Add(CData);

                                    if (Ltemp != null)
                                    {
                                        //Series ser = new Series(item.ToString(), ViewType.Bar);

                                        if (chartControlAnalCom.Series.Count > 0)
                                        {
                                            ser.Points.Add(new SeriesPoint(elt.Key+" "+elt.Value, SommeTotal));
                                        
                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            //tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            ////Annotations dans la legende
                                            ////AxisLabel TextPattern = "{}{V:#,##0}"

                                            //((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            //ser.CrosshairLabelPattern = "{V:n0}";
                                            //  }
                                        }
                                        else
                                        {
                                            ser.Points.Add(new SeriesPoint(elt.Key + " " + elt.Value, SommeTotal));

                                            ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                            ////tester que la serie existe ou pas dans le graphe
                                            //chartControlAnalCom.Series.Add(ser);
                                            ////Annotations dans la legende
                                            ////AxisLabel TextPattern = "{}{V:#,##0}"

                                            //((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                            //ser.CrosshairLabelPattern = "{V:n0}";
                                        }

                                    }

                                }

                            chartControlAnalCom.Series.Add(ser);
                            //Annotations dans la legende
                            //AxisLabel TextPattern = "{}{V:#,##0}"

                            ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                        }
                        if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            gridControlData.DataSource = LData;

                        //Masquer colonnes
                        gridView1.Columns["mAnnee"].Visible = true;
                        gridView1.Columns["mMontant"].Visible = true;
                        gridView1.Columns["mMois"].Visible = false;
                        gridView1.Columns["mCommercial"].Visible = true;
                        gridView1.Columns["mMarque"].Visible = false;
                        gridView1.Columns["mFamille"].Visible = false;
                        gridView1.Columns["mRevendeur"].Visible = false;

                        #endregion

                    }
                        

                }

                //Famille Centrale

                if (IsMarqueAxeMultiple)
                {
                    //Liste des Marques(Nom )
                    
                    LResult = daoReport.GetComparativeMultiMarqueAxe(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousCom.Checked, SMulCom.Checked, LECommercialDE.Text, LECommercialA.Text, ListNomMultiCommerciauxCritere, ListPrenomMultiCommerciauxCritere, SMulMarqueMultAxe.Checked, chkTousMarqueMultAxe.Checked, LEMarqueDEMultAxe.Text, LEMarqueAMultAxe.Text, ListMultiMarqueAxe, chkConfigFamilles.Checked, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDECritere, ListNomLEFamilleACritere, ListMultiFamilleCritere, TypDocId, SMulRevendeur.Checked, chkTousRevendeur.Checked, LERevendeurDE.Text, LERevendeurA.Text, ListMultiRevendeurCritere, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);

                    
                 List<string> ListeMarque = new List<string>();

                    if (LResult.Count > 0)
                    {
                        #region Liste des Marques
                        
                        foreach (var item in LResult)
                        {
                            
                            if(item.mFamilleCentral==string.Empty)
                            {
                                if (!ListeMarque.Contains("NO MARK"))
                                {
                                    ListeMarque.Add("NO MARK");
                                }

                            }
                            else
                            {
                                if (!ListeMarque.Contains(item.mFamilleCentral.Trim()))
                                {
                                    ListeMarque.Add(item.mFamilleCentral.Trim());
                                }
                            }
                                

                        }

                        #endregion

                    }

                    // if(!SMulFamille.Checked && chkTousFamille.Checked && !SMulRevendeur.Checked && chkTousRevendeur.Checked)
                    chartControlAnalCom.Series.Clear();

                    if (LResult.Count > 0)
                    {
                        #region chartControlAnalCom

                        chartControlAnalCom.Series.Clear();

                        foreach (var item in ListeAnnees)
                        {
                            Series ser = new Series(item.ToString(), ViewType.Bar);
                            foreach (var elt in ListeMarque)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                double SommeTotal = 0;

                                if(elt=="NO MARK")
                                {
                                    Ltemp = LResult.Where(c => c.mDatePiece.Year == Int32.Parse(item.ToString()) && c.mFamilleCentral == string.Empty).ToList();

                                }
                                else
                                {
                                    Ltemp = LResult.Where(c => c.mDatePiece.Year == Int32.Parse(item.ToString()) && c.mFamilleCentral.Trim() == elt.Trim()).ToList();

                                }


                                foreach (var obj in Ltemp)
                                {
                                    SommeTotal += obj.mMontantHT;
                                }

                                ComDataComp CData = new ComDataComp();

                                CData.mAnnee = Int32.Parse(item.ToString());
                                CData.mMois = string.Empty;
                                CData.mMontant = SommeTotal;

                                CData.mCommercial = string.Empty;
                                CData.mMarque = elt;
                                CData.mFamille = string.Empty;
                                CData.mRevendeur = string.Empty;

                                LData.Add(CData);

                                if (Ltemp != null)
                                {
                                    //Series ser = new Series(item.ToString(), ViewType.Bar);

                                    if (chartControlAnalCom.Series.Count > 0)
                                    {
                                        ser.Points.Add(new SeriesPoint(elt, SommeTotal));

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlAnalCom.Series.Add(ser);
                                        ////Annotations dans la legende
                                        ////AxisLabel TextPattern = "{}{V:#,##0}"

                                        //((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                        //  }
                                    }
                                    else
                                    {
                                        ser.Points.Add(new SeriesPoint(elt, SommeTotal));

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        ////tester que la serie existe ou pas dans le graphe
                                        //chartControlAnalCom.Series.Add(ser);
                                        ////Annotations dans la legende
                                        ////AxisLabel TextPattern = "{}{V:#,##0}"

                                        //((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                    }

                                }

                            }

                            chartControlAnalCom.Series.Add(ser);
                            //Annotations dans la legende
                            //AxisLabel TextPattern = "{}{V:#,##0}"

                            ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                        }
                        if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                        gridControlData.DataSource = LData;

                        //Masquer colonnes
                        gridView1.Columns["mAnnee"].Visible = true;
                        gridView1.Columns["mMontant"].Visible = true;
                        gridView1.Columns["mMois"].Visible = false;
                        gridView1.Columns["mCommercial"].Visible = false;
                        gridView1.Columns["mMarque"].Visible = true;
                        gridView1.Columns["mFamille"].Visible = false;
                        gridView1.Columns["mRevendeur"].Visible = false;

                        #endregion

                    }


                }

                
                //Revendeur

                if (IsRevendeurAxeMultiple)
                {
                    //Liste des Marques(Nom )

                    if (Int16.Parse(TypDocId) == 8)
                    {
                        LResult = daoReport.GetComparativeMultiRevendeurAxeBUNDLE(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousCom.Checked, SMulCom.Checked, LECommercialDE.Text, LECommercialA.Text, ListNomMultiCommerciauxCritere, ListPrenomMultiCommerciauxCritere, SMulMarque.Checked, chkTousMarque.Checked, LEMarqueDE.Text, LEMarqueA.Text, ListMultiMarqueCritere, chkConfigFamilles.Checked, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDECritere, ListNomLEFamilleACritere, ListMultiFamilleCritere, TypDocId, SMulRevendeurMultAxe.Checked, chkTousRevendeurMultAxe.Checked, LERevendeurDEMultAxe.Text, LERevendeurAMultAxe.Text, ListMultiRevendeurAxe, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);

                    }
                    else
                    {
                        LResult = daoReport.GetComparativeMultiRevendeurAxe(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousCom.Checked, SMulCom.Checked, LECommercialDE.Text, LECommercialA.Text, ListNomMultiCommerciauxCritere, ListPrenomMultiCommerciauxCritere, SMulMarque.Checked, chkTousMarque.Checked, LEMarqueDE.Text, LEMarqueA.Text, ListMultiMarqueCritere, chkConfigFamilles.Checked, SMulFamille.Checked, chkTousFamille.Checked, ListNomLEFamilleDECritere, ListNomLEFamilleACritere, ListMultiFamilleCritere, TypDocId, SMulRevendeurMultAxe.Checked, chkTousRevendeurMultAxe.Checked, LERevendeurDEMultAxe.Text, LERevendeurAMultAxe.Text, ListMultiRevendeurAxe, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);
                        
                    }
                    

                    List<string> ListeRevendeur = new List<string>();

                    if (LResult.Count > 0)
                    {
                        #region Liste des Revendeurs

                        foreach (var item in LResult)
                        {

                            if (!ListeRevendeur.Contains(item.mCT_Num))
                            {
                                ListeRevendeur.Add(item.mCT_Num);
                            }
                            

                        }

                        #endregion

                    }

                    // if(!SMulFamille.Checked && chkTousFamille.Checked && !SMulRevendeur.Checked && chkTousRevendeur.Checked)
                    chartControlAnalCom.Series.Clear();

                    if (LResult.Count > 0)
                    {
                        #region chartControlAnalCom

                        chartControlAnalCom.Series.Clear();

                        foreach (var item in ListeAnnees)
                        {
                            Series ser = new Series(item.ToString(), ViewType.Bar);
                            foreach (var elt in ListeRevendeur)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                double SommeTotal = 0;

                                    Ltemp = LResult.Where(c => c.mDatePiece.Year == Int32.Parse(item.ToString()) && c.mCT_Num.Trim() == elt.Trim()).ToList();
                                
                                foreach (var obj in Ltemp)
                                {
                                    SommeTotal += obj.mMontantHT;
                                }

                                ComDataComp CData = new ComDataComp();

                                CData.mAnnee = Int32.Parse(item.ToString());
                                CData.mMois = string.Empty;
                                CData.mMontant = SommeTotal;

                                CData.mCommercial = string.Empty;
                                CData.mMarque = string.Empty;
                                CData.mFamille = string.Empty;
                                CData.mRevendeur = elt;

                                LData.Add(CData);

                                if (Ltemp != null)
                                {
                                    //Series ser = new Series(item.ToString(), ViewType.Bar);

                                    if (chartControlAnalCom.Series.Count > 0)
                                    {
                                        ser.Points.Add(new SeriesPoint(elt, SommeTotal));

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlAnalCom.Series.Add(ser);
                                        ////Annotations dans la legende
                                        ////AxisLabel TextPattern = "{}{V:#,##0}"

                                        //((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                        //  }
                                    }
                                    else
                                    {
                                        ser.Points.Add(new SeriesPoint(elt, SommeTotal));

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        ////tester que la serie existe ou pas dans le graphe
                                        //chartControlAnalCom.Series.Add(ser);
                                        ////Annotations dans la legende
                                        ////AxisLabel TextPattern = "{}{V:#,##0}"

                                        //((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                    }

                                }

                            }

                            chartControlAnalCom.Series.Add(ser);
                            //Annotations dans la legende
                            //AxisLabel TextPattern = "{}{V:#,##0}"

                            ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                        }
                        if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                        gridControlData.DataSource = LData;
                        //Masquer colonnes
                        gridView1.Columns["mAnnee"].Visible = true;
                        gridView1.Columns["mMontant"].Visible = true;
                        gridView1.Columns["mMois"].Visible = false;
                        gridView1.Columns["mCommercial"].Visible = false;
                        gridView1.Columns["mMarque"].Visible = false;
                        gridView1.Columns["mFamille"].Visible = false;
                        gridView1.Columns["mRevendeur"].Visible = true;
                        #endregion

                    }


                }

                //Famille

                if (IsFamilleAxeMultiple)
                {
                    //Liste des Marques(Nom )

                    LResult = daoReport.GetComparativeMultiFamilleAxe(ChaineconnexionBABI, JourMoisDeb, JourMoisFin, myObjectFilialeChoisies, ListeAnnees, chkTousCom.Checked, SMulCom.Checked, LECommercialDE.Text, LECommercialA.Text, ListNomMultiCommerciauxCritere, ListPrenomMultiCommerciauxCritere,  SMulFamilleMultAxe.Checked, chkTousFamilleMultAxe.Checked, LEFamilleDEMultAxe.Text, LEFamilleAMultAxe.Text, ListMultiFamilleAxe, TypDocId, SMulRevendeur.Checked, chkTousRevendeur.Checked, LERevendeurDE.Text, LERevendeurA.Text, ListMultiRevendeurCritere, IsFinFevrierDeb, JourMoisDebBIX, IsFinFevrierFin, JourMoisFinBIX);


                    List<string> ListeFamille = new List<string>();

                    if (LResult.Count > 0)
                    {
                        #region Liste des Familles

                        foreach (var item in LResult)
                        {

                            if (!ListeFamille.Contains(item.mFamille))
                            {
                                ListeFamille.Add(item.mFamille);
                            }
                            
                        }

                        #endregion

                    }

                    // if(!SMulFamille.Checked && chkTousFamille.Checked && !SMulRevendeur.Checked && chkTousRevendeur.Checked)
                    chartControlAnalCom.Series.Clear();

                    if (LResult.Count > 0)
                    {
                        #region chartControlAnalCom

                        chartControlAnalCom.Series.Clear();

                        foreach (var item in ListeAnnees)
                        {
                            Series ser = new Series(item.ToString(), ViewType.Bar);
                            foreach (var elt in ListeFamille)
                            {
                                List<ComRevendeur> Ltemp = new List<ComRevendeur>();

                                double SommeTotal = 0;

                                Ltemp = LResult.Where(c => c.mDatePiece.Year == Int32.Parse(item.ToString()) && c.mFamille.Trim() == elt.Trim()).ToList();

                                foreach (var obj in Ltemp)
                                {
                                    SommeTotal += obj.mMontantHT;
                                }

                                ComDataComp CData = new ComDataComp();

                                CData.mAnnee = Int32.Parse(item.ToString());
                                CData.mMois = string.Empty;
                                CData.mMontant = SommeTotal;

                                CData.mCommercial = string.Empty;
                                CData.mMarque = string.Empty;
                                CData.mFamille = elt;
                                CData.mRevendeur = string.Empty;

                                LData.Add(CData);

                                if (Ltemp != null)
                                {
                                    //Series ser = new Series(item.ToString(), ViewType.Bar);

                                    if (chartControlAnalCom.Series.Count > 0)
                                    {
                                        ser.Points.Add(new SeriesPoint(elt, SommeTotal));

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlAnalCom.Series.Add(ser);
                                        ////Annotations dans la legende
                                        ////AxisLabel TextPattern = "{}{V:#,##0}"

                                        //((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                        //  }
                                    }
                                    else
                                    {
                                        ser.Points.Add(new SeriesPoint(elt, SommeTotal));

                                        ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                        ////tester que la serie existe ou pas dans le graphe
                                        //chartControlAnalCom.Series.Add(ser);
                                        ////Annotations dans la legende
                                        ////AxisLabel TextPattern = "{}{V:#,##0}"

                                        //((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                        //ser.CrosshairLabelPattern = "{V:n0}";
                                    }

                                }

                            }

                            chartControlAnalCom.Series.Add(ser);
                            //Annotations dans la legende
                            //AxisLabel TextPattern = "{}{V:#,##0}"

                            ((XYDiagram)chartControlAnalCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                        }
                        if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                        gridControlData.DataSource = LData;

                        //Masquer colonnes
                        gridView1.Columns["mAnnee"].Visible = true;
                        gridView1.Columns["mMontant"].Visible = true;
                        gridView1.Columns["mMois"].Visible = false;
                        gridView1.Columns["mCommercial"].Visible = false;
                        gridView1.Columns["mMarque"].Visible = false;
                        gridView1.Columns["mFamille"].Visible = true;
                        gridView1.Columns["mRevendeur"].Visible = false;

                        #endregion

                    }


                }

                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();

                timer1.Enabled = false;

                sBtnApercu.ForeColor = Color.Black;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> sBtnApercu_Click -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

       
        private string GetNomById(int IdMois)
        {
            string ret = string.Empty;
            try
            {
                if(IdMois>0)
                {
                    switch (IdMois)
                    {
                        case 1:
                            ret = "Janvier";
                            break;
                        case 2:
                            ret = "Février";
                            break;

                        case 3:
                            ret = "Mars";
                            break;

                        case 4:
                            ret = "Avril";
                            break;

                        case 5:
                            ret = "Mai";
                            break;

                        case 6:
                            ret = "Juin";
                            break;

                        case 7:
                            ret = "Juillet";
                            break;

                        case 8:
                            ret = "Août";
                            break;

                        case 9:
                            ret = "Septembre";
                            break;

                        case 10:
                            ret = "Octobre";
                            break;

                        case 11:
                            ret = "Novembre";
                            break;

                        case 12:
                            ret = "Décembre";
                            break;

                    }
                }

                
                return ret;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> GetNomById -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return ret;
            }

        }



        private void radioSemestre_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioSemestre.Checked)
                {
                    panSemestre.Visible = true;
                    panTrimestre.Visible = false;
                
                    panPeriode.Visible = false;

                    CmbSemestre.ItemIndex = 0;

                IsTrimetre=false;
                IsSemetre=true;
                IsPeriodeAdefinir=false;
                    IsAnnuel = false;
                    IsMensuel = false;

                    CleanGrid();
                    timer1.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> radioSemestre_Click -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void radioTrimestre_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioTrimestre.Checked)
                {
                    panTrimestre.Visible = true;
                    panSemestre.Visible = false;
                    panPeriode.Visible = false;
                    CmbTrimestre.ItemIndex = 0;


                    IsTrimetre = true;
                    IsSemetre = false;
                    IsPeriodeAdefinir = false;
                    IsAnnuel = false;
                    IsMensuel = false;

                    CleanGrid();
                    timer1.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> radioTrimestre_Click -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void radioADefinir_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioADefinir.Checked)
                {
                    panTrimestre.Visible = false;
                    panSemestre.Visible = false;
                    panPeriode.Visible = true;


                    IsTrimetre = false;
                    IsSemetre = false;
                    IsPeriodeAdefinir = true;
                    IsAnnuel = false;
                    IsMensuel = false;

                    LMoisDeb.ItemIndex = 0;
                    LMoisFin.ItemIndex = 0;

                    CleanGrid();
                    timer1.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> radioADefinir_Click -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void radioAnnuel_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioAnnuel.Checked)
                {
                    panTrimestre.Visible = false;
                    panSemestre.Visible = false;
                    panPeriode.Visible = false;


                    IsTrimetre = false;
                    IsSemetre = false;
                    IsPeriodeAdefinir = false;
                    IsAnnuel = true;
                    IsMensuel = false;

                    CleanGrid();
                    timer1.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> radioAnnuel_Click -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void radioMensuel_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioMensuel.Checked)
                {
                    panTrimestre.Visible = false;
                    panSemestre.Visible = false;
                    panPeriode.Visible = false;


                    IsTrimetre = false;
                    IsSemetre = false;
                    IsPeriodeAdefinir = false;
                    IsAnnuel = false;
                    IsMensuel = true;

                    CleanGrid();
                    timer1.Enabled = true;
                }
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> radioMensuel_Click -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }


        private void HideAnneMois()
        {
            try
            {
                if (radioMensuel.Checked) radioMensuel.Checked = false;

                radioMensuel.Visible = false;

            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> HideAnneMois -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }


        private void ShowAnneMois()
        {
            try
            {
               
                radioMensuel.Visible = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> ShowAnneMois -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEAxeAnalyse_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                int choix = Int32.Parse(LEAxeAnalyse.EditValue.ToString());
                
                switch (choix)
                {
                    case 1://Commercial
                        pCommercialAxe.Visible = true;
                        pMarqueAxe.Visible = false;
                        pFamilleAxe.Visible = false;
                        pRevendeurAxe.Visible = false;

                        //A cause de la gestion des BUNDLE ,on tiendra pas compte des Familles et marques
                        //dans les critères
                        pCommercialCritere.Visible = false;
                        pFamilleCritere.Visible = false;
                        pRevendeurCritere.Visible = true;
                        pMarqueCritere.Visible = false;
                        
                        //Multiple
                        pCommercialAxeMult.Visible = false;
                        pMarqueMultAxe.Visible = false;
                        pRevendeurMultAxe.Visible = false;
                        pFamilleMultAxe.Visible = false;

                        IsCommercialAxe = true;
                        IsFamilleAxe = false;
                        IsMarqueAxe = false;
                        IsRevendeurAxe = false;

                        //Multiple
                        IsCommercialAxeMultiple = false;
                        IsFamilleAxeMultiple = false;
                        IsMarqueAxeMultiple = false;
                        IsRevendeurAxeMultiple = false;

                        ShowAnneMois();

                        break;
                    case 2://Marque
                        pCommercialAxe.Visible = false;
                        pMarqueAxe.Visible = true;
                        pFamilleAxe.Visible = false;
                        pRevendeurAxe.Visible = false;

                        pCommercialCritere.Visible = true;
                        pFamilleCritere.Visible = true;
                        pRevendeurCritere.Visible = true;
                        pMarqueCritere.Visible = false;

                        //Multiple
                        pCommercialAxeMult.Visible = false;
                        pMarqueMultAxe.Visible = false;
                        pRevendeurMultAxe.Visible = false;
                        pFamilleMultAxe.Visible = false;

                        IsCommercialAxe = false;
                        IsFamilleAxe = false;
                        IsMarqueAxe = true;
                        IsRevendeurAxe = false;

                        //Multiple
                        IsCommercialAxeMultiple = false;
                        IsFamilleAxeMultiple = false;
                        IsMarqueAxeMultiple = false;
                        IsRevendeurAxeMultiple = false;

                        //Charger la liste des famille pour critère par rapport a marque par defaut
                        FillComboFamilleByChoixMarque();
                        ShowAnneMois();
                        break;

                    case 3://Famille
                        pCommercialAxe.Visible = false;
                        pMarqueAxe.Visible = false;
                        pFamilleAxe.Visible = true;
                        pRevendeurAxe.Visible = false;

                        pCommercialCritere.Visible = true;
                        pFamilleCritere.Visible = false;
                        pRevendeurCritere.Visible = true;
                        pMarqueCritere.Visible = false;

                        //Multiple
                        pCommercialAxeMult.Visible = false;
                        pMarqueMultAxe.Visible = false;
                        pRevendeurMultAxe.Visible = false;
                        pFamilleMultAxe.Visible = false;

                        IsCommercialAxe = false;
                        IsFamilleAxe = true;
                        IsMarqueAxe = false;
                        IsRevendeurAxe = false;

                        //Multiple
                        IsCommercialAxeMultiple = false;
                        IsFamilleAxeMultiple = false;
                        IsMarqueAxeMultiple = false;
                        IsRevendeurAxeMultiple = false;
                        ShowAnneMois();
                        break;

                    case 4://Revendeur
                        pCommercialAxe.Visible = false;
                        pMarqueAxe.Visible = false;
                        pFamilleAxe.Visible = false;
                        pRevendeurAxe.Visible = true;

                        //A cause de la gestion des BUNDLE ,on tiendra pas compte des Familles et marques
                        //dans les critères
                        pCommercialCritere.Visible = true;
                        pFamilleCritere.Visible = false;
                        pRevendeurCritere.Visible = false;
                       // pMarqueCritere.Visible = true;
                        pMarqueCritere.Visible = false;

                        //Multiple
                        pCommercialAxeMult.Visible = false;
                        pMarqueMultAxe.Visible = false;
                        pRevendeurMultAxe.Visible = false;
                        pFamilleMultAxe.Visible = false;

                        IsCommercialAxe = false;
                        IsFamilleAxe = false;
                        IsMarqueAxe = false;
                        IsRevendeurAxe = true;

                        //Multiple
                        IsCommercialAxeMultiple = false;
                        IsFamilleAxeMultiple = false;
                        IsMarqueAxeMultiple = false;
                        IsRevendeurAxeMultiple = false;
                        ShowAnneMois();
                        break;


                    //MULTIPLE

                    case 5://Commercial multiple
                        pCommercialAxe.Visible = false;
                        pMarqueAxe.Visible = false;
                        pFamilleAxe.Visible = false;
                        pRevendeurAxe.Visible = false;

                        //A cause de la gestion des BUNDLE ,on tiendra pas compte des Familles et marques
                        //dans les critères
                        pCommercialCritere.Visible = false;
                        pFamilleCritere.Visible = false;
                        pRevendeurCritere.Visible = true;
                       // pMarqueCritere.Visible = true;
                        pMarqueCritere.Visible = false;

                        //Multiple
                        pCommercialAxeMult.Visible = true;
                        pMarqueMultAxe.Visible = false;
                        pRevendeurMultAxe.Visible = false;
                        pFamilleMultAxe.Visible = false;

                        IsCommercialAxe = false;
                        IsFamilleAxe = false;
                        IsMarqueAxe = false;
                        IsRevendeurAxe = false;

                        //Multiple
                        IsCommercialAxeMultiple = true;
                        IsFamilleAxeMultiple = false;
                        IsMarqueAxeMultiple = false;
                        IsRevendeurAxeMultiple = false;

                        HideAnneMois();

                        break;
                    case 6://Marque
                        pCommercialAxe.Visible = false;
                        pMarqueAxe.Visible = false;
                        pFamilleAxe.Visible = false;
                        pRevendeurAxe.Visible = false;

                        pCommercialCritere.Visible = true;
                        pFamilleCritere.Visible = true;
                        pRevendeurCritere.Visible = true;
                        pMarqueCritere.Visible = false;

                        //Multiple
                        pCommercialAxeMult.Visible = false;
                        pMarqueMultAxe.Visible = true;
                        pRevendeurMultAxe.Visible = false;
                        pFamilleMultAxe.Visible = false;

                        IsCommercialAxe = false;
                        IsFamilleAxe = false;
                        IsMarqueAxe = false;
                        IsRevendeurAxe = false;

                        //Multiple
                        IsCommercialAxeMultiple = false;
                        IsFamilleAxeMultiple = false;
                        IsMarqueAxeMultiple = true;
                        IsRevendeurAxeMultiple = false;

                        //Charger la liste des famille pour critère par rapport a marque par defaut
                        FillComboFamilleByChoixMarque();
                        HideAnneMois();
                        break;

                    case 7://Famille
                        pCommercialAxe.Visible = false;
                        pMarqueAxe.Visible = false;
                        pFamilleAxe.Visible = false;
                        pRevendeurAxe.Visible = false;

                        pCommercialCritere.Visible = true;
                        pFamilleCritere.Visible = false;
                        pRevendeurCritere.Visible = true;
                        pMarqueCritere.Visible = false;

                        //Multiple
                        pCommercialAxeMult.Visible = false;
                        pMarqueMultAxe.Visible = false;
                        pRevendeurMultAxe.Visible = false;
                        pFamilleMultAxe.Visible = true;

                        IsCommercialAxe = false;
                        IsFamilleAxe = false;
                        IsMarqueAxe = false;
                        IsRevendeurAxe = false;

                        //Multiple
                        IsCommercialAxeMultiple = false;
                        IsFamilleAxeMultiple = true;
                        IsMarqueAxeMultiple = false;
                        IsRevendeurAxeMultiple = false;
                        HideAnneMois();
                        break;

                    case 8://Revendeur
                        pCommercialAxe.Visible = false;
                        pMarqueAxe.Visible = false;
                        pFamilleAxe.Visible = false;
                        pRevendeurAxe.Visible = false;

                        //A cause de la gestion des BUNDLE ,on tiendra pas compte des Familles et marques
                        //dans les critères
                        pCommercialCritere.Visible = true;
                        pFamilleCritere.Visible = false;
                        pRevendeurCritere.Visible = false;
                       // pMarqueCritere.Visible = true;
                        pMarqueCritere.Visible = false;

                        //Multiple
                        pCommercialAxeMult.Visible = false;
                        pMarqueMultAxe.Visible = false;
                        pRevendeurMultAxe.Visible = true;
                        pFamilleMultAxe.Visible = false;

                        IsCommercialAxe = false;
                        IsFamilleAxe = false;
                        IsMarqueAxe = false;
                        IsRevendeurAxe = false;

                        //Multiple
                        IsCommercialAxeMultiple = false;
                        IsFamilleAxeMultiple = false;
                        IsMarqueAxeMultiple = false;
                        IsRevendeurAxeMultiple = true;
                        HideAnneMois();
                        break;

                }
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEAxeAnalyse_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkTousMarque_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousMarque.CheckState == CheckState.Checked)
                {
                    LEMarqueDE.Enabled = false;
                    LEMarqueA.Enabled = false;
                }
                else
                {
                    LEMarqueDE.Enabled = true;
                    LEMarqueA.Enabled = true;
                }

                CancelConfigFamille();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousMarque_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

            }
        }

        private void chkTousCom_CheckedChanged(object sender, EventArgs e)
        {

            try
            {
                if (chkTousCom.CheckState == CheckState.Checked)
                {
                    LECommercialDE.Enabled = false;
                    LECommercialA.Enabled = false;
                }
                else
                {
                    LECommercialDE.Enabled = true;
                    LECommercialA.Enabled = true;
                }
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousCom_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

            }
        }

        private void SMulCom_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulCom.CheckState == CheckState.Checked)
                {
                    chkCmbMultCom.Visible = true;

                    chkTousCom.Visible = false;

                    LECommercialA.Visible = false;

                    LECommercialDE.Visible = false;

                    lblCommercialA.Visible = false;
                    lblCommercialDE.Visible = false;

                }
                else
                {
                    chkCmbMultCom.Visible = false;

                    chkTousCom.Visible = true;

                    LECommercialA.Visible = true;

                    LECommercialDE.Visible = true;

                    lblCommercialA.Visible = true;
                    lblCommercialDE.Visible = true;
                }
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> SMulCom_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void SMulMarque_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulMarque.CheckState == CheckState.Checked)
                {
                    chkCmbMultMarque.Visible = true;

                    chkTousMarque.Visible = false;

                    LEMarqueA.Visible = false;

                    LEMarqueDE.Visible = false;

                    lblMarqueDe.Visible = false;
                    lblMarqueA.Visible = false;

                }
                else
                {
                    chkCmbMultMarque.Visible = false;

                    chkTousMarque.Visible = true;

                    LEMarqueA.Visible = true;

                    LEMarqueDE.Visible = true;

                    lblMarqueDe.Visible = true;
                    lblMarqueA.Visible = true;
                }

                CancelConfigFamille();
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> SMulMarque_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkTousRevendeurAxe_CheckStateChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousRevendeurAxe.CheckState == CheckState.Checked)
                {
                    LERevendeur.Enabled = false;

                }
                else
                {
                    LERevendeur.Enabled = true;

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousRevendeurAxe_CheckStateChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkTousFamilleAxe_CheckStateChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousFamilleAxe.CheckState == CheckState.Checked)
                {
                    LEFamille.Enabled = false;

                }
                else
                {
                    LEFamille.Enabled = true;

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousFamilleAxe_CheckStateChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkTousMarqueAxe_CheckStateChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousMarqueAxe.CheckState == CheckState.Checked)
                {
                    LEMarque.Enabled = false;

                }
                else
                {
                    LEMarque.Enabled = true;

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousMarqueAxe_CheckStateChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkConfigFamilles_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkConfigFamilles.CheckState == CheckState.Checked)
                {
                    pFamilleCritere.Visible = true;

                    ReLoadFamilleByMarque(chkTousMarque.Checked,SMulMarque.Checked,LEMarqueDE.Text,LEMarqueA.Text,chkCmbMultMarque.Text);

                }
                else
                {
                    pFamilleCritere.Visible = false;

                }

                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkConfigFamilles_CheckedChanged -> TypeErreur: " + ex.Message; ;
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
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> CancelConfigFamille -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

      


        private void ReLoadFamilleByMarque(bool IsTousMarque, bool IsMultiMarque, string MarqueDE, string MarqueA, string MultiMarque)
        {
            List<string> Lidpays = new List<string>();
            List<string> LFamCentral = new List<string>();
            List<ComFamille> LFamilleFill = new List<ComFamille>();
            try
            {
                //foreach (CheckedListBoxItem item in ChkListBoxControlFiliale.Items)
                //{
                //    if (item.CheckState == CheckState.Checked)
                //    {
                //        var Choose = ListeFiliales.FirstOrDefault(c => c.mId == Convert.ToInt32(item.Value.ToString()));

                //        Lidpays.Add(item.Value.ToString());
                //    }

                //}

                foreach(var elt in myObjectFilialeChoisies)
                {
                    Lidpays.Add(elt.mId.ToString());
                }

                if (IsMultiMarque)
                {
                    var tab = MultiMarque.Split(',');

                    for (int i = 0; i < tab.Length; i++)
                    {
                        LFamCentral.Add(tab[i].Trim());
                    }

                    LFamilleFill = daoReport.GetListByListIdPaysFamCentrFam(Lidpays, myObjectFamille, LFamCentral);
                    LEFamilleDE.Properties.DataSource = LFamilleFill;
                    LEFamilleA.Properties.DataSource = LFamilleFill;

                    //Vider le multifamille ET LES DE_A au cas où il y aurait eu un précédent choix
                    ListMultiFamilleCritere = string.Empty;
                    ListNomLEFamilleDECritere = string.Empty;
                    ListNomLEFamilleACritere = string.Empty;
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
                            LFamilleFill = daoReport.GetListByListIdFam(Lidpays, myObjectFamille);

                            LEFamilleDE.Properties.DataSource = LFamilleFill;
                            LEFamilleA.Properties.DataSource = LFamilleFill;
                            ////Vider le multifamille ET LES DE_A au cas où il y aurait eu un précédent choix
                            ListMultiFamilleCritere = string.Empty;
                            ListNomLEFamilleDECritere = string.Empty;
                            ListNomLEFamilleACritere = string.Empty;
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
                            ListFC = myObjectMarque.Where(c => c.mIdPays == Convert.ToInt32(item.ToString())).ToList();

                            ListFCTMP.AddRange(ListFC);
                        }


                        if (ListFCTMP.Count > 0)
                        {
                            ListFCTMP = ListFCTMP.OrderBy(x => x.mFa_CodeFamille).ToList();

                            foreach (var obj in ListFCTMP)
                            {
                                if (MarqueDE == obj.mFa_CodeFamille)
                                {
                                    trouveDeb = true;

                                    ListFamCentralisatrice.Add(obj.mFa_CodeFamille);
                                }

                                if (trouveDeb && !trouveFin)
                                {
                                    var IsExist = ListFamCentralisatrice.FirstOrDefault(c => c.Equals(obj.mFa_CodeFamille));

                                    if (IsExist == null) ListFamCentralisatrice.Add(obj.mFa_CodeFamille);
                                }

                                if (MarqueA == obj.mFa_CodeFamille)
                                {
                                    trouveFin = true;
                                    var IsExist = ListFamCentralisatrice.FirstOrDefault(c => c.Equals(obj.mFa_CodeFamille));

                                    if (IsExist == null) ListFamCentralisatrice.Add(obj.mFa_CodeFamille);
                                }

                                if (trouveDeb && trouveFin)
                                {
                                    break;
                                }

                            }

                            LFamilleFill = daoReport.GetListByListIdPaysFamCentrFam(Lidpays, myObjectFamille, ListFamCentralisatrice);

                            LEFamilleDE.Properties.DataSource = LFamilleFill;
                            LEFamilleA.Properties.DataSource = LFamilleFill;

                            ////Vider le multifamille ET LES DE_A au cas où il y aurait eu un précédent choix
                            ListMultiFamilleCritere = string.Empty;
                            ListNomLEFamilleDECritere = string.Empty;
                            ListNomLEFamilleACritere = string.Empty;
                            /////////////////////////////////////////////////////////

                            chkCmbMultFam.Properties.DataSource = LFamilleFill;
                            FillMultiCheckComboFamille(chkCmbMultFam, LFamilleFill);

                            LEFamilleDE.Properties.DisplayMember = "mFa_CodeFamille";
                            LEFamilleA.Properties.DisplayMember = "mFa_CodeFamille";
                        }



                    }


                }



            }
            catch (Exception ex)
            {

                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> ReLoadFamilleByMarque -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

            }
        }

        private void LEMarque_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                //Charger la liste des famille pour critère par rapport a marque par defaut
                FillComboFamilleByChoixMarque();

                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEMarque_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

       

        private void chkCmbMultCom_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                
                ListNomMultiCommerciauxCritere = string.Empty;

                ListPrenomMultiCommerciauxCritere = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultCom.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        //if (item.Value.ToString() != string.Empty)
                        if (!item.Value.ToString().Equals(item.Description))
                        {
                            //le value c'est le prenom ,
                            ListPrenomMultiCommerciauxCritere += item.Value.ToString().Trim() + ",";

                            //si la description ramenée= prénom donc on a pas de nom
                            if (item.Value.ToString().Equals(item.Description))
                            {
                                ListNomMultiCommerciauxCritere += " " + ",";
                            }
                            else
                            {

                                //on peut donc retirer le nom par déduction d'avec le texte(description =Nom +" "+Prenom)
                                string tmp = " " + item.Value.ToString().Trim();
                                ListNomMultiCommerciauxCritere += item.Description.Replace(tmp, "") + ",";
                            }


                        }
                        else
                        {
                            //On a que le nom

                            //le value c'est le prenom donc il prend vide,
                            ListPrenomMultiCommerciauxCritere += " " + ",";

                            ListNomMultiCommerciauxCritere += item.Description + ",";


                        }

                    }

                }

                if (ListPrenomMultiCommerciauxCritere.Length > 0) ListPrenomMultiCommerciauxCritere = ListPrenomMultiCommerciauxCritere.Substring(0, ListPrenomMultiCommerciauxCritere.Length - 1);
                if (ListNomMultiCommerciauxCritere.Length > 0) ListNomMultiCommerciauxCritere = ListNomMultiCommerciauxCritere.Substring(0, ListNomMultiCommerciauxCritere.Length - 1);
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                timer1.Enabled = true;
                CleanGrid();
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkCmbMultCom_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleDE_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
              
                ListNomLEFamilleDECritere = LEFamilleDE.Text;
                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEFamilleDE_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleA_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListNomLEFamilleACritere = LEFamilleA.Text;
                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEFamilleA_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultFam_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListMultiFamilleCritere = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultFam.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if (item.Value != null)
                        {

                            ListMultiFamilleCritere += item.Value.ToString().Trim() + ",";

                        }

                    }

                }

                if (ListMultiFamilleCritere.Length > 0) ListMultiFamilleCritere = ListMultiFamilleCritere.Substring(0, ListMultiFamilleCritere.Length - 1);

                
                CleanGrid();
                timer1.Enabled = true;

            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkCmbMultFam_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void sJourFin_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                int mois = Int32.Parse(LMoisFin.EditValue.ToString());

                switch (mois)
                {
                    case 1://Janvier
                        sJourFin.Properties.MaxValue = 31;
                        break;
                    case 2://Fevrier
                        sJourFin.Properties.MaxValue = 28;
                        break;

                    case 3://MARS
                        sJourFin.Properties.MaxValue = 31;
                        break;

                    case 4://AVRIL
                        sJourFin.Properties.MaxValue = 30;
                        break;

                    case 5://MAI
                        sJourFin.Properties.MaxValue = 31;
                        break;

                    case 6://Juin
                        sJourFin.Properties.MaxValue = 30;
                        break;

                    case 7://JUILLET
                        sJourFin.Properties.MaxValue = 31;
                        break;
                    case 8://Aout
                        sJourFin.Properties.MaxValue = 31;
                        break;
                    case 9://Semptemre
                        sJourFin.Properties.MaxValue = 30;
                        break;

                    case 10://Octobre
                        sJourFin.Properties.MaxValue = 31;
                        break;

                    case 11://Novembre
                        sJourFin.Properties.MaxValue = 30;
                        break;

                    case 12://dECEMBRE
                        sJourFin.Properties.MaxValue = 31;
                        break;

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> sJourFin_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void sJourDeb_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                int mois = Int32.Parse(LMoisDeb.EditValue.ToString());

                switch (mois)
                {

                    case 1://Janvier
                        sJourDeb.Properties.MaxValue = 31;
                        break;
                    case 2://Fevrier
                        sJourDeb.Properties.MaxValue = 28;
                        break;

                    case 3://MARS
                        sJourDeb.Properties.MaxValue = 31;
                        break;

                    case 4://AVRIL
                        sJourDeb.Properties.MaxValue = 30;
                        break;

                    case 5://MAI
                        sJourDeb.Properties.MaxValue = 31;
                        break;

                    case 6://Juin
                        sJourDeb.Properties.MaxValue = 30;
                        break;

                    case 7://JUILLET
                        sJourDeb.Properties.MaxValue = 31;
                        break;
                    case 8://Aout
                        sJourDeb.Properties.MaxValue = 31;
                        break;
                    case 9://Semptemre
                        sJourDeb.Properties.MaxValue = 30;
                        break;

                    case 10://Octobre
                        sJourDeb.Properties.MaxValue = 31;
                        break;

                    case 11://Novembre
                        sJourDeb.Properties.MaxValue = 30;
                        break;

                    case 12://dECEMBRE
                        sJourDeb.Properties.MaxValue = 31;
                        break;

                }
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> sJourDeb_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LMoisFin_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                int mois = Int32.Parse(LMoisFin.EditValue.ToString());

                switch (mois)
                {
                    case 1://Janvier
                        sJourFin.Properties.MaxValue = 31;
                        break;
                    case 2://Fevrier
                        sJourFin.Properties.MaxValue = 28;
                        break;

                    case 3://MARS
                        sJourFin.Properties.MaxValue = 31;
                        break;

                    case 4://AVRIL
                        sJourFin.Properties.MaxValue = 30;
                        break;

                    case 5://MAI
                        sJourFin.Properties.MaxValue = 31;
                        break;

                    case 6://Juin
                        sJourFin.Properties.MaxValue = 30;
                        break;

                    case 7://JUILLET
                        sJourFin.Properties.MaxValue = 31;
                        break;
                    case 8://Aout
                        sJourFin.Properties.MaxValue = 31;
                        break;
                    case 9://Semptemre
                        sJourFin.Properties.MaxValue = 30;
                        break;

                    case 10://Octobre
                        sJourFin.Properties.MaxValue = 31;
                        break;

                    case 11://Novembre
                        sJourFin.Properties.MaxValue = 30;
                        break;

                    case 12://dECEMBRE
                        sJourFin.Properties.MaxValue = 31;
                        break;

                }
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LMoisFin_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LMoisDeb_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                int mois = Int32.Parse(LMoisDeb.EditValue.ToString());

                switch (mois)
                {
                    case 1://Janvier
                        sJourDeb.Properties.MaxValue = 31;
                        break;
                    case 2://Fevrier
                        sJourDeb.Properties.MaxValue = 28;
                        break;

                    case 3://MARS
                        sJourDeb.Properties.MaxValue = 31;
                        break;

                    case 4://AVRIL
                        sJourDeb.Properties.MaxValue = 30;
                        break;

                    case 5://MAI
                        sJourDeb.Properties.MaxValue = 31;
                        break;

                    case 6://Juin
                        sJourDeb.Properties.MaxValue = 30;
                        break;

                    case 7://JUILLET
                        sJourDeb.Properties.MaxValue = 31;
                        break;
                    case 8://Aout
                        sJourDeb.Properties.MaxValue = 31;
                        break;
                    case 9://Semptemre
                        sJourDeb.Properties.MaxValue = 30;
                        break;

                    case 10://Octobre
                        sJourDeb.Properties.MaxValue = 31;
                        break;

                    case 11://Novembre
                        sJourDeb.Properties.MaxValue = 30;
                        break;

                    case 12://dECEMBRE
                        sJourDeb.Properties.MaxValue = 31;
                        break;

                }

                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LMoisDeb_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void CmbTypeDoc_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> CmbTypeDoc_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LERevendeur_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LERevendeur_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamille_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEFamille_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void sNumAnneeDeb_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> sNumAnneeDeb_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void sNumAnneeFin_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> sNumAnneeFin_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void CmbTrimestre_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> CmbTrimestre_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void CmbSemestre_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> CmbSemestre_EditValueChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LECommercialDE_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LECommercialDE_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LECommercialA_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LECommercialA_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LERevendeurDE_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LERevendeurDE_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LERevendeurA_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LERevendeurA_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultRev_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                ListMultiRevendeurCritere = string.Empty;
                
                foreach (CheckedListBoxItem item in chkCmbMultRev.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if (item.Value != null)
                        {

                            ListMultiRevendeurCritere += item.Value.ToString().Trim() + ",";

                        }

                    }

                }

                if (ListMultiRevendeurCritere.Length > 0) ListMultiRevendeurCritere = ListMultiRevendeurCritere.Substring(0, ListMultiRevendeurCritere.Length - 1);



                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkCmbMultRev_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEMarqueDE_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEMarqueDE_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEMarqueA_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEMarqueA_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultMarque_Closed(object sender, ClosedEventArgs e)
        {
            try
            {

                ListMultiMarqueCritere = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultMarque.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if (item.Value != null)
                        {

                            ListMultiMarqueCritere += item.Value.ToString().Trim() + ",";

                        }

                    }

                }

                if (ListMultiMarqueCritere.Length > 0) ListMultiMarqueCritere = ListMultiMarqueCritere.Substring(0, ListMultiMarqueCritere.Length - 1);
                
                CleanGrid();
                CancelConfigFamille();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkCmbMultMarque_Closed -> TypeErreur: " + ex.Message; ;
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
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                var msg = "FenAnalyseComp -> GetPidExcelAvant -> TypeErreur: " + ex.Message; ;
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
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                var msg = "FenAnalyseComp -> TuerProcessus -> TypeErreur: " + ex.Message; ;
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
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                var msg = "FenAnalyseComp -> GetPidExcelApres -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return ret;
            }

        }


        public ComFiltreAnalComp GetFiltre()
        {
            ComFiltreAnalComp cf = new ComFiltreAnalComp();
            try
            {
                cf.mAxeAnalyse = LEAxeAnalyse.Text;

                //élément Analysé

                int choix = Int32.Parse(LEAxeAnalyse.EditValue.ToString());
                switch (choix)
                {
                    case 1://Commercial

                        if(chkTousCommerciauxAxe.Checked)
                        {

                            cf.mElementAnalyse = "Commercial : Tous les Commerciaux";
                        }
                        else
                        {
                            cf.mElementAnalyse = "Commercial : " + LECommercial.Text;
                        }

                        break;
                    case 2://Marque

                        if (chkTousMarqueAxe.Checked)
                        {

                            cf.mElementAnalyse = "Marque : Toutes les Marques";
                        }
                        else
                        {
                            cf.mElementAnalyse = "Marque : " + LEMarque.Text;
                        }


                        break;

                    case 3://Famille

                        if (chkTousFamilleAxe.Checked)
                        {

                            cf.mElementAnalyse = "Famille : Toutes les Familles";
                        }
                        else
                        {

                            cf.mElementAnalyse = "Famille : " + LEFamille.Text;
                        }


                        break;

                    case 4://Revendeur

                        if (chkTousRevendeurAxe.Checked)
                        {

                            cf.mElementAnalyse = "Revendeurs : Tous les Revendeurs";
                        }
                        else
                        {
                            cf.mElementAnalyse = "Revendeur : " + LERevendeur.Text;
                        }
                  
                        break;


                    case 5://Multi Commerciaux

                        if (SMulComMultAxe.Checked)
                        {
                            cf.mElementAnalyse = "Commerciaux : " + chkCmbMultComAxe.Text;
                        }
                        else
                        {
                            if (chkTousComMultAxe.Checked)
                            {
                                cf.mElementAnalyse = "Commerciaux : Tous les Commerciaux";
                            }
                            else
                            {
                                //DE_A
                                cf.mElementAnalyse = "Commerciaux : De " + LECommercialDEMultAxe.Text + " A " + LECommercialAMultAxe.Text;
                            }
                        }


                        break;


                    case 6://Multi Marque

                        if (SMulMarqueMultAxe.Checked)
                        {
                            cf.mElementAnalyse = "Marques : " + chkCmbMultMarqueAxe.Text;
                        }
                        else
                        {
                            if (chkTousMarqueMultAxe.Checked)
                            {
                                cf.mElementAnalyse = "Marques : Toutes les Marques";
                            }
                            else
                            {
                                //DE_A
                                cf.mElementAnalyse = "Marques : De " + LEMarqueDEMultAxe.Text + " A " + LEMarqueAMultAxe.Text;
                            }
                        }

                        break;

                    case 7://Multi Famille

                        if (SMulFamilleMultAxe.Checked)
                        {
                            cf.mElementAnalyse = "Familles : " + chkCmbMultFamAxe.Text;
                        }
                        else
                        {
                            if (chkTousFamilleMultAxe.Checked)
                            {
                                cf.mElementAnalyse = "Familles : Toutes les Familles";
                            }
                            else
                            {
                                //DE_A
                                cf.mElementAnalyse = "Familles : De " + LEFamilleDEMultAxe.Text + " A " + LEFamilleAMultAxe.Text;
                            }
                        }

                        break;

                    case 8://Multi Revendeurs

                        if (SMulRevendeurMultAxe.Checked)
                        {
                            cf.mElementAnalyse = "Revendeurs : " + chkCmbMultRevAxe.Text;
                        }
                        else
                        {
                            if (chkTousRevendeurMultAxe.Checked)
                            {
                                cf.mElementAnalyse = "Revendeurs : Tous les Revendeurs";
                            }
                            else
                            {
                                //DE_A
                                cf.mElementAnalyse = "Revendeurs : De " + LERevendeurDEMultAxe.Text + " A " + LERevendeurAMultAxe.Text;
                            }
                        }

                        break;

                }


                //Type de document
             cf.mTypeDoc = "Type Document : " + CmbTypeDoc.Text;
                    
                //Periode
                if(SMulAnnee.Checked)
                {
                    cf.mPeriode = "Période : " + ChkCmbMulAn.Text;
                }
                else
                {

                    if (sNumAnneeDeb.Text != string.Empty && sNumAnneeFin.Text != string.Empty)
                    {
                        cf.mPeriode = "Période : De " + sNumAnneeDeb.Text + " A " + sNumAnneeFin.Text;
                    }
                }



                //Periode Analyse

                if(IsTrimetre)
                {
                    cf.mPeriodeAnalyse = "Saisonnalité : Trimestrielle";
                }

                if (IsSemetre)
                {
                    cf.mPeriodeAnalyse = "Saisonnalité : Semestrielle";
                }
                if (IsPeriodeAdefinir)
                {
                    cf.mPeriodeAnalyse = "Saisonnalité : A définir";

                }

                if (IsAnnuel)
                {
                    cf.mPeriodeAnalyse = "Saisonnalité : Annuelle";
                }

                if (IsMensuel)
                {
                    cf.mPeriodeAnalyse = "Saisonnalité : Annuelle(Par mois)";
                }

                //Periode analyse définition

                if (IsTrimetre)
                {
                    cf.mPeriodeAnalyseDefinition = "Période Analyse : "+CmbTrimestre.Text;
                }

                if (IsSemetre)
                {
                    cf.mPeriodeAnalyseDefinition = "Période Analyse : " + CmbSemestre.Text;
                }

                if (IsPeriodeAdefinir)
                {
                    cf.mPeriodeAnalyseDefinition = "Période : Du " + sJourDeb.Text + "  " + LMoisDeb.Text + " Au " + sJourFin.Text + " " + LMoisFin.Text;

                }

                if (IsAnnuel)
                {
                    cf.mPeriodeAnalyseDefinition = "Période Analyse : Annuelle";
                }

                if (IsMensuel)
                {
                    cf.mPeriodeAnalyseDefinition = "Période Analyse : Annuelle(Par mois)";
                }


                //Critère

                if (pCommercialCritere.Visible)
                {

                    if (SMulCom.Checked)
                    {
                        cf.mCommercialCritere = "Commerciaux : " + chkCmbMultCom.Text;
                    }
                    else
                    {
                        if (chkTousCom.Checked)
                        {
                            cf.mCommercialCritere = "Commerciaux : Tous les Commerciaux";
                        }
                        else
                        {
                            //DE_A
                            cf.mCommercialCritere = "Commerciaux : De " + LECommercialA.Text + " A " + LECommercialDE.Text;
                        }
                    }
                }


                //Famille

                if(pFamilleCritere.Visible)
                {
                    if (SMulFamille.Checked)
                    {
                        cf.mFamilleCritere = "Famille : " + chkCmbMultFam.Text;
                    }
                    else
                    {
                        if (chkTousFamille.Checked)
                        {
                            cf.mFamilleCritere = "Famille : Toutes les Familles";
                        }
                        else
                        {
                            //DE_A
                            cf.mFamilleCritere = "Famille : De " + LEFamilleDE.Text + " A " + LEFamilleA.Text;
                        }
                    }
                }


                //Famille Centralisatrice(Marque)
                if (pMarqueCritere.Visible)
                {

                    if (SMulMarque.Checked)
                    {
                        cf.mMarqueCritere = "Marque : " + chkCmbMultMarque.Text;
                    }
                    else
                    {
                        if (chkTousMarque.Checked)
                        {
                            cf.mMarqueCritere = "Marque : Toutes les Marques";
                        }
                        else
                        {
                            //DE_A
                            cf.mMarqueCritere = "Marque : De " + LEMarqueDE.Text + " A " + LEMarqueA.Text;
                        }
                    }
                }

             

                //Revendeur

                if(pRevendeurCritere.Visible)
                {
                    if (SMulRevendeur.Checked)
                    {
                        cf.mRevendeurCritere = "Revendeur : " + chkCmbMultRev.Text;
                    }
                    else
                    {
                        if (chkTousRevendeur.Checked)
                        {
                            cf.mRevendeurCritere = "Revendeur : Tous les Revendeurs";
                        }
                        else
                        {
                            //DE_A
                            cf.mRevendeurCritere = "Revendeur : De " + LERevendeurDE.Text + " A " + LERevendeurA.Text;
                        }
                    }
                }

                
                //Site choisi

                if (myObjectFilialeChoisies.Count > 0)
                {

                    string listeSite = string.Empty;

                    foreach (var item in myObjectFilialeChoisies)
                    {
                        listeSite = item.mVille + ",";

                    }

                    if (listeSite.Length > 1) listeSite = listeSite.Substring(0, listeSite.Length - 1);

                    cf.mSite = "Site(s) : " + listeSite;
                }

                return cf;
            }
            catch (Exception ex)
            {

                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAnalyseComp ->GetFiltre -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }


        }


        
         


        private void sBtnExportExcel_Click(object sender, EventArgs e)
        {

            try
            {
                string cheminFichier = string.Empty;

                string Titre = string.Empty;

                int processIdAvant = 0;

                int PidExcelproc = 0;

                string chaineBABI = string.Empty;

                ComFiltreAnalComp FiltreObject=new ComFiltreAnalComp();


                if (xtraTabControlApercu.SelectedTabPage == xtraTabPageData)
                {
                    //Tester si la grid est vide ou pas
                    if (gridView1.RowCount == 0)
                    {
                        MessageBox.Show("Il n'y a aucune données à exporter en Excel!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
               

                }

                //Statistiques
                if (xtraTabControlApercu.SelectedTabPage == xtraTabPageStat)
                {
                    if (chartControlAnalCom.Series.Count == 0)
                    {
                        MessageBox.Show("Il n'y a aucune données à exporter en Excel!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                   

                }

                //Les filtres
                FiltreObject= GetFiltre();


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
                    
                    if (!splashScreenManager1.IsSplashFormVisible) splashScreenManager1.ShowWaitForm();

                    if (xtraTabControlApercu.SelectedTabPage == xtraTabPageData)
                    {
                        //Tester si la grid est vide ou pas
                       
                            gridControlData.ExportToXlsx(cheminFichier);

                            Titre = "Comparatif des Données";
                        
                    }

                    //Statistiques
                    if (xtraTabControlApercu.SelectedTabPage == xtraTabPageStat)
                    {
                            chartControlAnalCom.ExportToXlsx(cheminFichier);

                            Titre = "Comparatif des Statistiques ";
                        
                    }


                    splashScreenManager1.SetWaitFormDescription("Traitement en cours...20%");

                    //Obtenir la somme des pid Excel ouvert avant qu'on initialise notre objet
                    processIdAvant = GetPidExcelAvant();

                    Microsoft.Office.Interop.Excel.Application xlsApp;
                    Microsoft.Office.Interop.Excel.Workbook xlsClasseur;
                    Microsoft.Office.Interop.Excel.Worksheet xlsFeuille;


                    // Créer un objet Excel
                    xlsApp = new Microsoft.Office.Interop.Excel.Application();

                    //recupérer le Pid de l'objet pour le tuer à la fin ou en cas d'erreur

                    int pidTot = 0;

                    pidTot = GetPidExcelApres();

                    //On le tuera à la fin ou en cas d'exception
                    PidExcelproc = pidTot + processIdAvant;

                    splashScreenManager1.SetWaitFormDescription("Traitement en cours...40%");

                    if (xlsApp == null)
                    {
                        ////installer Microsoft Excel sur le poste
                        if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                        MessageBox.Show("Veuiller installer Excel sur votre poste au préalable!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        //tuer le processus créé

                        bool iskok = false;

                        iskok = TuerProcessus(PidExcelproc);

                        if (!iskok)
                        {
                            if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                            MessageBox.Show("Une erreur est survenue lors de la suppression du processus!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        }

                        return;

                    }

                    //Ouvrir le fichier et inscrire les filtres

                    // Ne pas tenir compte des alertes
                    xlsApp.DisplayAlerts = false;

                    //// Ajout d'un classeur
                    xlsClasseur = xlsApp.Workbooks.Open(cheminFichier);
                    splashScreenManager1.SetWaitFormDescription("Traitement en cours...60%");

                    ((Microsoft.Office.Interop.Excel.Worksheet)xlsClasseur.Worksheets[1]).Select();

                    xlsFeuille = (Microsoft.Office.Interop.Excel.Worksheet)xlsClasseur.Worksheets[1];

                    splashScreenManager1.SetWaitFormDescription("Traitement en cours...65%");

                    bool isImputsOK = false;
                    
                    
                    //Données
                    if (xtraTabControlApercu.SelectedTabPage == xtraTabPageData)
                    {

                        int ligne = 1;

                        int colonne = 5;

                        //titre
                        isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne, colonne, Titre);
                        //sites
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 1, colonne, FiltreObject.mSite);

                        //Element Analysé
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 2, colonne, FiltreObject.mElementAnalyse);
                        
                        //periode
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 3, colonne, FiltreObject.mPeriode);

                        //Saisonnalité
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 4, colonne, FiltreObject.mPeriodeAnalyse);

                        //Période analyse définit
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 5, colonne, FiltreObject.mPeriodeAnalyseDefinition);

                        int choix = Int32.Parse(LEAxeAnalyse.EditValue.ToString());
                        switch (choix)
                        {
                            case 1://Commercial

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mMarqueCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);


                                break;
                            case 2://Marque

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);
                                
                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);
                                
                                break;

                            case 3://Famille

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);
                                
                                break;

                            case 4://Revendeur

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mMarqueCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);
                                
                                break;

                            case 5://Multi Commerciaux

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mMarqueCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);

                                break;

                            case 6://Multi Marque

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);

                                break;

                            case 7://Multi Famille

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);

                                break;

                            case 8://Multi Revendeur

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mMarqueCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);

                                break;

                        }



                        splashScreenManager1.SetWaitFormDescription("Traitement en cours...80%");


                    }

                    
                    //Statistiques
                    if (xtraTabControlApercu.SelectedTabPage == xtraTabPageStat)
                    {

                        int ligne = 2;

                        int colonne = 1;

                        //titre
                        isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne, colonne, Titre);
                        //sites
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 1, colonne, FiltreObject.mSite);

                        //Element Analysé
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 2, colonne, FiltreObject.mElementAnalyse);


                        //periode
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 3, colonne, FiltreObject.mPeriode);

                        //Saisonnalité
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 4, colonne, FiltreObject.mPeriodeAnalyse);

                        //Période analyse définit
                        if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 5, colonne, FiltreObject.mPeriodeAnalyseDefinition);

                        int choix = Int32.Parse(LEAxeAnalyse.EditValue.ToString());
                        switch (choix)
                        {
                            case 1://Commercial

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mMarqueCritere);
                                
                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);


                                break;
                            case 2://Marque

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);


                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);
                                
                                break;

                            case 3://Famille

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);

                                
                                break;

                            case 4://Revendeur

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mMarqueCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);
                                
                                break;


                            case 5://Multi Commercial

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mMarqueCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);


                                break;
                            case 6://Multi Marque

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);


                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);

                                break;

                            case 7://Multi Famille

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeurCritere);


                                break;

                            case 8://Multi Revendeur

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mCommercialCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mMarqueCritere);

                                if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 8, colonne, FiltreObject.mFamilleCritere);

                                break;

                        }



                        splashScreenManager1.SetWaitFormDescription("Traitement en cours...80%");
                   

                    }

                    splashScreenManager1.SetWaitFormDescription("Traitement en cours...90%");
                    xlsApp.ActiveWorkbook.SaveAs(cheminFichier);
                    xlsApp.ActiveWorkbook.Close();

                    bool isExcelok = false;

                    isExcelok = TuerProcessus(PidExcelproc);

                    if (!isExcelok)
                    {
                        if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                        MessageBox.Show("Une erreur est survenue lors de la suppression du processus!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    }
                    splashScreenManager1.SetWaitFormDescription("Traitement en cours...100%");
                    //Generation Excel
                    if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                    MessageBox.Show("Export Excel effectuée avec succès!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAnalyseComp ->sBtnExportExcel_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
              
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                if (sBtnApercu.ForeColor == Color.Black)
                {
                    sBtnApercu.ForeColor = Color.Red;
                }
                else
                {
                    if (sBtnApercu.ForeColor == Color.Red)
                    {
                        sBtnApercu.ForeColor = Color.Black;
                    }
                }
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAnalyseComp ->timer1_Tick -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
           
            }
        }

        private void chkTousRevendeurMultAxe_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousComMultAxe.CheckState == CheckState.Checked)
                {
                    LECommercialDEMultAxe.Enabled = false;
                    LECommercialAMultAxe.Enabled = false;

                }
                else
                {
                    LECommercialDEMultAxe.Enabled = true;
                    LECommercialAMultAxe.Enabled = true;

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousRevendeurMultAxe_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }
        

        private void SMulComMultAxe_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulComMultAxe.CheckState == CheckState.Checked)
                {
                    chkCmbMultComAxe.Visible = true;

                    chkTousComMultAxe.Visible = false;

                    LECommercialAMultAxe.Visible = false;

                    LECommercialDEMultAxe.Visible = false;

                    lblComAMultAxe.Visible = false;
                    lblComDEMultAxe.Visible = false;

                }
                else
                {
                    chkCmbMultComAxe.Visible = false;

                    chkTousComMultAxe.Visible = true;

                    LECommercialAMultAxe.Visible = true;

                    LECommercialDEMultAxe.Visible = true;

                    lblComAMultAxe.Visible = true;
                    lblComDEMultAxe.Visible = true;
                }
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> SMulComMultAxe_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkTousMarqueMultAxe_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousMarqueMultAxe.CheckState == CheckState.Checked)
                {
                    LEMarqueDEMultAxe.Enabled = false;
                    LEMarqueAMultAxe.Enabled = false;

                }
                else
                {
                    LEMarqueDEMultAxe.Enabled = true;
                    LEMarqueAMultAxe.Enabled = true;

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousMarqueMultAxe_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void SMulMarqueMultAxe_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulMarqueMultAxe.CheckState == CheckState.Checked)
                {
                    chkCmbMultMarqueAxe.Visible = true;

                    chkTousMarqueMultAxe.Visible = false;

                    LEMarqueAMultAxe.Visible = false;

                    LEMarqueDEMultAxe.Visible = false;

                    lblMarqueAMultAxe.Visible = false;
                    lblMarqueDEMultAxe.Visible = false;

                }
                else
                {
                    chkCmbMultMarqueAxe.Visible = false;

                    chkTousMarqueMultAxe.Visible = true;

                    LEMarqueAMultAxe.Visible = true;

                    LEMarqueDEMultAxe.Visible = true;

                    lblMarqueAMultAxe.Visible = true;
                    lblMarqueDEMultAxe.Visible = true;
                }
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> SMulMarqueMultAxe_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkTousRevendeurMultAxe_CheckedChanged_1(object sender, EventArgs e)
        {
            try
            {
                if (chkTousRevendeurMultAxe.CheckState == CheckState.Checked)
                {
                    LERevendeurDEMultAxe.Enabled = false;
                    LERevendeurAMultAxe.Enabled = false;

                }
                else
                {
                    LERevendeurDEMultAxe.Enabled = true;
                    LERevendeurAMultAxe.Enabled = true;

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousRevendeurMultAxe_CheckedChanged_1 -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void SMulRevendeurMultAxe_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulRevendeurMultAxe.CheckState == CheckState.Checked)
                {
                    chkCmbMultRevAxe.Visible = true;

                    chkTousRevendeurMultAxe.Visible = false;

                    LERevendeurAMultAxe.Visible = false;

                    LERevendeurDEMultAxe.Visible = false;

                    lblRevendeurAMultAxe.Visible = false;
                    lblRevendeurDEMultAxe.Visible = false;

                }
                else
                {
                    chkCmbMultRevAxe.Visible = false;

                    chkTousRevendeurMultAxe.Visible = true;

                    LERevendeurAMultAxe.Visible = true;

                    LERevendeurDEMultAxe.Visible = true;

                    lblRevendeurAMultAxe.Visible = true;
                    lblRevendeurDEMultAxe.Visible = true;
                }
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> SMulRevendeurMultAxe_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkTousFamilleMultAxe_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (chkTousFamilleMultAxe.CheckState == CheckState.Checked)
                {
                    LEFamilleDEMultAxe.Enabled = false;
                    LEFamilleAMultAxe.Enabled = false;

                }
                else
                {
                    LEFamilleDEMultAxe.Enabled = true;
                    LEFamilleAMultAxe.Enabled = true;

                }
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkTousFamilleMultAxe_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void SMulFamilleMultAxe_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (SMulFamilleMultAxe.CheckState == CheckState.Checked)
                {
                    chkCmbMultFamAxe.Visible = true;

                    chkTousFamilleMultAxe.Visible = false;

                    LEFamilleAMultAxe.Visible = false;

                    LEFamilleDEMultAxe.Visible = false;

                    lblFamilleAMultAxe.Visible = false;
                    lblFamilleDEMultAxe.Visible = false;

                }
                else
                {
                    chkCmbMultFamAxe.Visible = false;

                    chkTousFamilleMultAxe.Visible = true;

                    LEFamilleAMultAxe.Visible = true;

                    LEFamilleDEMultAxe.Visible = true;

                    lblFamilleAMultAxe.Visible = true;
                    lblFamilleDEMultAxe.Visible = true;
                }
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                CleanGrid();
                timer1.Enabled = true;

            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> SMulFamilleMultAxe_CheckedChanged -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        //Gérer les filtres familles pour multi Marque AXE
        
        private void xtraTabControl1_Click(object sender, EventArgs e)
        {

            try
            {

                if(xtraTabControl1.SelectedTabPage== xtraTabPageFiltres)
                {

                    if (IsMarqueAxeMultiple)
                    {
                        pFamilleCritere.Visible = true;

                        ReLoadFamilleByMarque(chkTousMarqueMultAxe.Checked, SMulMarqueMultAxe.Checked, LEMarqueDEMultAxe.Text, LEMarqueAMultAxe.Text, chkCmbMultMarqueAxe.Text);

                    }
                    else//Pas de filtyres sur les marques donc on charge tout comme au départ
                    {

                        if (!chkConfigFamilles.Checked && chkTousMarque.Checked && !IsMarqueAxe)
                        {
                            //Remplir Combo Famille==============================================

                            LEFamilleDE.Properties.DataSource = myObjectFamille;
                            LEFamilleA.Properties.DataSource = myObjectFamille;

                            chkCmbMultFam.Properties.DataSource = myObjectFamille;
                            FillMultiCheckComboFamille(chkCmbMultFam, myObjectFamille);

                            LEFamilleDE.Properties.DisplayMember = "mFa_CodeFamille";
                            LEFamilleA.Properties.DisplayMember = "mFa_CodeFamille";

                            //Choisir les premières valeurs
                            LEFamilleDE.ItemIndex = 0;
                            LEFamilleA.ItemIndex = 0;

                        }
                        else
                        {
                            if(chkConfigFamilles.Checked)
                            {
                                CancelConfigFamille();
                            }
                        }

                    }
                }


            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> xtraTabControl1_Click -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultComAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {

                ListNomMultiCommerciauxMultipleAxe = string.Empty;

                ListPrenomMultiCommerciauxMultipleAxe = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultComAxe.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        //if (item.Value.ToString() != string.Empty)
                        if (!item.Value.ToString().Equals(item.Description))
                        {
                            //le value c'est le prenom ,
                            ListPrenomMultiCommerciauxMultipleAxe += item.Value.ToString().Trim() + ",";

                            //si la description ramenée= prénom donc on a pas de nom
                            if (item.Value.ToString().Equals(item.Description))
                            {
                                ListNomMultiCommerciauxMultipleAxe += " " + ",";
                            }
                            else
                            {
                                //on peut donc retirer le nom par déduction d'avec le texte(description =Nom +" "+Prenom)
                                string tmp = " " + item.Value.ToString().Trim();
                                ListNomMultiCommerciauxMultipleAxe += item.Description.Replace(tmp, "") + ",";
                            }


                        }
                        else
                        {
                            //On a que le nom

                            //le value c'est le prenom donc il prend vide,
                            ListPrenomMultiCommerciauxMultipleAxe += " " + ",";

                            ListNomMultiCommerciauxMultipleAxe += item.Description + ",";


                        }

                    }

                }

                if (ListPrenomMultiCommerciauxMultipleAxe.Length > 0) ListPrenomMultiCommerciauxMultipleAxe = ListPrenomMultiCommerciauxMultipleAxe.Substring(0, ListPrenomMultiCommerciauxMultipleAxe.Length - 1);
                if (ListNomMultiCommerciauxMultipleAxe.Length > 0) ListNomMultiCommerciauxMultipleAxe = ListNomMultiCommerciauxMultipleAxe.Substring(0, ListNomMultiCommerciauxMultipleAxe.Length - 1);
                //lblControlTotalHT.Text = "0";
                //lblControlTotalHT2.Text = "0";
                timer1.Enabled = true;
                CleanGrid();
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkCmbMultComAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultMarqueAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {

                ListMultiMarqueAxe = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultMarqueAxe.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if (item.Value != null)
                        {

                            ListMultiMarqueAxe += item.Value.ToString().Trim() + ",";

                        }

                    }

                }

                if (ListMultiMarqueAxe.Length > 0) ListMultiMarqueAxe = ListMultiMarqueAxe.Substring(0, ListMultiMarqueAxe.Length - 1);

                CleanGrid();
                CancelConfigFamille();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkCmbMultMarqueAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultRevAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {

                ListMultiRevendeurAxe = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultRevAxe.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if (item.Value != null)
                        {

                            ListMultiRevendeurAxe += item.Value.ToString().Trim() + ",";

                        }

                    }

                }

                if (ListMultiRevendeurAxe.Length > 0) ListMultiRevendeurAxe = ListMultiRevendeurAxe.Substring(0, ListMultiRevendeurAxe.Length - 1);

                CleanGrid();
    
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkCmbMultRevAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LECommercialAMultAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LECommercialAMultAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LECommercialDEMultAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LECommercialDEMultAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEMarqueAMultAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEMarqueAMultAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEMarqueDEMultAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEMarqueDEMultAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LERevendeurAMultAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LERevendeurAMultAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LERevendeurDEMultAxe_Closed(object sender, ClosedEventArgs e)
        {

            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LERevendeurDEMultAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void chkCmbMultFamAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {

                ListMultiFamilleAxe = string.Empty;

                foreach (CheckedListBoxItem item in chkCmbMultFamAxe.Properties.Items)
                {
                    if (item.CheckState == CheckState.Checked)
                    {
                        if (item.Value != null)
                        {

                            ListMultiFamilleAxe += item.Value.ToString().Trim() + ",";

                        }

                    }

                }

                if (ListMultiFamilleAxe.Length > 0) ListMultiFamilleAxe = ListMultiFamilleAxe.Substring(0, ListMultiFamilleAxe.Length - 1);

                CleanGrid();

                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> chkCmbMultFamAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleDEMultAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEFamilleDEMultAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }

        private void LEFamilleAMultAxe_Closed(object sender, ClosedEventArgs e)
        {
            try
            {
                CleanGrid();
                timer1.Enabled = true;
            }
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                var msg = "FenAnalyseComp -> LEFamilleAMultAxe_Closed -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
        }
    }
}
