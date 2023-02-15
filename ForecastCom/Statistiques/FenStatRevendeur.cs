using DevExpress.XtraCharts;
using ForecastCom.Models;
using ForecastCom.Services;
using ForecastCom.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ForecastCom.Statistiques
{
    public partial class FenStatRevendeur : Form
    {
        private ComFiltre FiltreObject;

        //type doc -Montant-Revendeur
        private List<ComRevendeur> myObjectTypeDocMtantRev;

        //Activité Commerciaux pour Revendeur
        private List<ComRevendeur> myObjectComMtantRev;

        //Top 10 meilleurs Revendeurs
        private List<ComRevendeur> myObjectTOPMtantRevendeur;

        //les 10 mauvais Revendeurs
        private List<ComRevendeur> myObjectBADMtantRevendeur;

        private readonly FormatageForecastCom FormatMain = new FormatageForecastCom();


        private bool TestIsOnlyFact;

        public FenStatRevendeur(List<ComRevendeur> LcomStatTypeDocRevendeur, ComFiltre Cfiltre, List<ComRevendeur> LcomStatComRevendeur, List<ComRevendeur> LcomStatTOPRevendeur, List<ComRevendeur> LcomStatBADRevendeur,bool IsOnlyFact)
        {
            try
            {
                InitializeComponent();

                this.FiltreObject = Cfiltre;
                this.myObjectTypeDocMtantRev = LcomStatTypeDocRevendeur;
                this.myObjectComMtantRev = LcomStatComRevendeur;
                this.myObjectTOPMtantRevendeur = LcomStatTOPRevendeur;
                this.myObjectBADMtantRevendeur = LcomStatBADRevendeur;

                this.TestIsOnlyFact = IsOnlyFact;
            }
            catch(Exception ex)
            {
                var msg = "FenStatRevendeur -> FenStatRevendeur -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
            }
     
        }

        private void FenStatRevendeur_Load(object sender, EventArgs e)
        {
            try
            {
                #region Type Documents
                List<string> ListTypeDoc = new List<string>();
                
                if (FiltreObject.mIsPipe == false)
                {
                    foreach (var elt in myObjectTypeDocMtantRev)
                    {
                        if (ListTypeDoc.Count == 0)
                        {
                            ListTypeDoc.Add(elt.mTypeDoc);
                        }
                        else
                        {
                            var a = ListTypeDoc.Contains(elt.mTypeDoc);

                            if (a == false)
                            {
                                ListTypeDoc.Add(elt.mTypeDoc);

                            }

                        }
                    }
                }
                else
                {
                    //cas du pipe

                    foreach (var elt in myObjectTypeDocMtantRev)
                    {
                        if (ListTypeDoc.Count == 0)
                        {
                            ListTypeDoc.Add(elt.mStatut);
                        }
                        else
                        {
                            var a = ListTypeDoc.Contains(elt.mStatut);

                            if (a == false)
                            {
                                ListTypeDoc.Add(elt.mStatut);

                            }

                        }
                    }


                }

                #endregion

                #region Liste des Revendeurs

                Dictionary<string, string> ListeRevendeurs = new Dictionary<string, string>();

                foreach (var item in myObjectTypeDocMtantRev)
                {
                    if (ListeRevendeurs.Count == 0)
                    {
                        ListeRevendeurs.Add(item.mCT_Num, item.mCT_Intitule);
                    }
                    else
                    {
                        var a = ListeRevendeurs.ContainsKey(item.mCT_Num);

                        if (a == false)
                        {
                            var b = ListeRevendeurs.ContainsKey(item.mCT_Intitule);

                            if (b == false)
                            {
                                ListeRevendeurs.Add(item.mCT_Num, item.mCT_Intitule);
                            }
                        }

                    }
                }

                #endregion

                #region chartControlTypeDoc

                if(TestIsOnlyFact)
                {
                    if (FiltreObject.mIsPipe == false)
                    {
                        //foreach (var it in ListTypeDoc)
                        //{
                            Series ser = new Series("Facture[FA+FC]" + " (En Montant)", ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlNbreTypeDocRev.Series.Add(ser);

                            ((XYDiagram)chartControlNbreTypeDocRev.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            foreach (var item in ListeRevendeurs)
                            {
                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                List<ComRevendeur> LtempTypeDoc = new List<ComRevendeur>();

                                LtempTypeDoc = myObjectTypeDocMtantRev.Where(c => c.mCT_Num == item.Key && c.mCT_Intitule == item.Value).ToList();
                                
                                double SommeTotal = 0;
                                

                                foreach (var obj in LtempTypeDoc)
                                {
                                    SommeTotal += obj.mMontantHT;

                                }

                                ComRevendeur LtempCom = new ComRevendeur();

                                LtempCom = myObjectTypeDocMtantRev.FirstOrDefault(c => c.mCT_Num == item.Key && c.mCT_Intitule == item.Value);



                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlNbreTypeDocRev.Series.Count > 0)
                                    {
                                        
                                        ser.Points.Add(new SeriesPoint(LtempCom.mCT_Intitule, SommeTotal));
                                       
                                    }
                                    else
                                    {
                                       
                                        ser.Points.Add(new SeriesPoint(LtempCom.mCT_Intitule, SommeTotal));
                                        
                                    }



                                }

                            }

                      //  }

                    }
                    else
                    {
                        //cas du pipe

                        foreach (var it in ListTypeDoc)
                        {
                            Series ser = new Series(it, ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlNbreTypeDocRev.Series.Add(ser);

                            ((XYDiagram)chartControlNbreTypeDocRev.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            foreach (var item in ListeRevendeurs)

                            {
                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                ComRevendeur LtempTypeDoc = new ComRevendeur();

                                LtempTypeDoc = myObjectTypeDocMtantRev.FirstOrDefault(c => c.mCT_Num == item.Key && c.mCT_Intitule == item.Value && c.mStatut == it);

                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlNbreTypeDocRev.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mCT_Intitule, LtempTypeDoc.mMontantHT));
                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mCT_Intitule, LtempTypeDoc.mMontantHT));

                                        //}
                                        //chartControlTypeDoc.Series.Add(ser);
                                    }



                                }

                            }

                        }


                    }
                }
                else
                {
                    if (FiltreObject.mIsPipe == false)
                    {
                        foreach (var it in ListTypeDoc)
                        {
                            Series ser = new Series(it, ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlNbreTypeDocRev.Series.Add(ser);

                            ((XYDiagram)chartControlNbreTypeDocRev.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            foreach (var item in ListeRevendeurs)
                            {
                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                ComRevendeur LtempTypeDoc = new ComRevendeur();

                                LtempTypeDoc = myObjectTypeDocMtantRev.FirstOrDefault(c => c.mCT_Num == item.Key && c.mCT_Intitule == item.Value && c.mTypeDoc == it);

                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlNbreTypeDocRev.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mCT_Intitule, LtempTypeDoc.mMontantHT));
                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mCT_Intitule, LtempTypeDoc.mMontantHT));

                                        //}
                                        //chartControlTypeDoc.Series.Add(ser);
                                    }



                                }

                            }

                        }

                    }
                    else
                    {
                        //cas du pipe

                        foreach (var it in ListTypeDoc)
                        {
                            Series ser = new Series(it, ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlNbreTypeDocRev.Series.Add(ser);

                            ((XYDiagram)chartControlNbreTypeDocRev.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            foreach (var item in ListeRevendeurs)

                            {
                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                ComRevendeur LtempTypeDoc = new ComRevendeur();

                                LtempTypeDoc = myObjectTypeDocMtantRev.FirstOrDefault(c => c.mCT_Num == item.Key && c.mCT_Intitule == item.Value && c.mStatut == it);

                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlNbreTypeDocRev.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mCT_Intitule, LtempTypeDoc.mMontantHT));
                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mCT_Intitule, LtempTypeDoc.mMontantHT));

                                        //}
                                        //chartControlTypeDoc.Series.Add(ser);
                                    }



                                }

                            }

                        }


                    }
                }


              

                #endregion
                
                #region Top10Commerciaux qui transactent avec Revendeur Montant

                if (myObjectComMtantRev != null && myObjectComMtantRev.Count > 0)
                {
                    // Create a pie series.
                    Series SeriesComRevMtant = new Series("Top 10 des Transactions Commerciales avec le(s) Revendeur(s) en Montant", ViewType.Pie);

                    //Top 10 Commerciaux par montant=====================================
                    foreach (var item in myObjectComMtantRev)

                    {
                        if(item.mMontantHT>0)
                        {
                            SeriesComRevMtant.Points.Add(new SeriesPoint(item.mPrenomCommercial + " " + item.mNomCommercial, item.mMontantHT));

                        }

                    }
                    
                    // Add the series to the chart.
                    chartControlTopComRev.Series.Add(SeriesComRevMtant);

                    // Format the the series labels(Annotations sur le graphique)
                    SeriesComRevMtant.Label.TextPattern = "{A}: {VP:p0}";

                    //Annotations dans la legende
                    SeriesComRevMtant.LegendTextPattern = "{A}: {V:n0}";

                    // Access the view-type-specific options of the series.
                    PieSeriesView myView = (PieSeriesView)SeriesComRevMtant.View;

                    // Show a title for the series.
                    myView.Titles.Add(new SeriesTitle());
                    myView.Titles[0].Text = SeriesComRevMtant.Name;

                }

                #endregion


                #region Top10Revendeurs

                if (myObjectTOPMtantRevendeur != null && myObjectTOPMtantRevendeur.Count > 0)
                {
                    // Create a pie series.
                    Series SeriesTopRevMtant = new Series("Top 10 des Revendeurs en Montant", ViewType.Pie);

                    //Top 10 Commerciaux par montant=====================================
                    foreach (var item in myObjectTOPMtantRevendeur)

                    {
                       if(item.mMontantHT>0) SeriesTopRevMtant.Points.Add(new SeriesPoint(item.mCT_Num, item.mMontantHT));

                    }

                    // Add the series to the chart.
                    chartControlTopRev.Series.Add(SeriesTopRevMtant);

                    // Format the the series labels(Annotations sur le graphique)
                    SeriesTopRevMtant.Label.TextPattern = "{A}: {VP:p0}";

                    //Annotations dans la legende
                    SeriesTopRevMtant.LegendTextPattern = "{A}: {V:n0}";

                    // Access the view-type-specific options of the series.
                    PieSeriesView myView = (PieSeriesView)SeriesTopRevMtant.View;

                    // Show a title for the series.
                    myView.Titles.Add(new SeriesTitle());
                    myView.Titles[0].Text = SeriesTopRevMtant.Name;

                }

                #endregion

                #region BAD10Revendeurs
                
                    if (myObjectBADMtantRevendeur != null && myObjectBADMtantRevendeur.Count > 0)
                    {
                  
                    
                        // Create a pie series.
                        
                        Series ser = new Series("Les 10 Mauvais Revendeurs en Montant", ViewType.Bar);

                        //Top 10 famille par montant=====================================
                        foreach (var obj in myObjectBADMtantRevendeur)
                        {
                        
                            if (obj.mMontantHT >= 0) ser.Points.Add(new SeriesPoint(obj.mCT_Num, obj.mMontantHT.ToString("n0")));
                         
                        }

                        ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                        // Add the series to the chart.
                        chartControlBadRevendeur.Series.Add(ser);


                        //((XYDiagram)chartControlBadRevendeur.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                        //// Rotate the diagram (if necessary).
                        //((XYDiagram)chartControlBadRevendeur.Diagram).Rotated = true;

                    
                    }

                    #endregion

                }
            catch (Exception ex)
            {
                var msg = "FenStatRevendeur -> FenStatRevendeur_Load -> TypeErreur: " + ex.Message; 
                CAlias.Log(msg);
            }
        }

        private void BtnExportExcelStatRev_Click(object sender, EventArgs e)
        {
            try
            {
                string cheminFichier = string.Empty;

                string Titre = string.Empty;

                int processIdAvant = 0;

                int PidExcelproc = 0;

                string chaineBABI = string.Empty;

                //On a des eléments ,on peut imprimer dans excel

                if (sFDExcelStatRev.ShowDialog() == DialogResult.OK)
                {
                    cheminFichier = sFDExcelStatRev.FileName;

                    if (cheminFichier.Length >= 203)
                    {
                        //chemin trop long

                        MessageBox.Show("Le chemin spécifié est trop long !Veuillez choisir un autre dossier!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    if (!splashScreenManager1.IsSplashFormVisible) splashScreenManager1.ShowWaitForm();

                    if (xtraTabControlStatRev.SelectedTabPage == xtraTabPageTypeDocMtantRev)
                    {
                        chartControlNbreTypeDocRev.ExportToXlsx(cheminFichier);

                        Titre = "Statistiques Montant par Revendeur et Type Document";
                    }
                   
                    if (xtraTabControlStatRev.SelectedTabPage == xtraTabPageComRev)
                    {
                        chartControlTopComRev.ExportToXlsx(cheminFichier);

                        Titre = "Top 10 des transactions commerciales ";
                    }
                    if (xtraTabControlStatRev.SelectedTabPage == xtraTabPageTopRevendeur)
                    {
                        chartControlTopRev.ExportToXlsx(cheminFichier);

                        Titre = "Top 10 des Revendeurs en Montant ";
                    }

                    if (xtraTabControlStatRev.SelectedTabPage == xtraTabPageBadRevendeur)
                    {
                        chartControlBadRevendeur.ExportToXlsx(cheminFichier);

                        Titre = "Les 10 Mauvais Revendeurs ";
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

                    int ligne = 2;

                    int colonne = 1;

                    //titre
                    isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne, colonne, Titre);
                    //sites
                    if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 1, colonne, FiltreObject.mSite);
                    //periode
                    if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 2, colonne, FiltreObject.mPeriode);

                    splashScreenManager1.SetWaitFormDescription("Traitement en cours...80%");
                    //Commerciaux
                    if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 3, colonne, FiltreObject.mCommerciaux);

                    //famille Centrale
                    if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 4, colonne, FiltreObject.mFamilleCentral);
                    
                    //famille
                    if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 5, colonne, FiltreObject.mFamille);
                    //Type docu
                    if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 6, colonne, FiltreObject.mTypeDoc);

                    //Revendeur
                    if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne + 7, colonne, FiltreObject.mRevendeur);


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
            catch (Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenStatistiques ->BtnExportExcelStat_Click -> TypeErreur: " + ex.Message;
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
                //if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                var msg = "MainForm -> GetPidExcelAvant -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return ret;
            }

        }

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

                    if (!Process.GetProcessById(processToKill).HasExited)
                    {
                        Process.GetProcessById(processToKill).Kill();
                    }


                    test = true;

                    return test;
                }

            }
            catch (Exception ex)
            {
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
                // if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                var msg = "MainForm -> GetPidExcelApres -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return ret;
            }

        }


    }
}
