using DevExpress.XtraCharts;
using DevExpress.XtraSplashScreen;
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

namespace ForecastCom
{
    public partial class FenStatistiques : Form
    {
        private List<FCommercial> myObject;

        private readonly FormatageForecastCom FormatMain = new FormatageForecastCom();

        private List<FCommercial> myObjectTypeDoc;
        private List<FCommercial> myObjectNbreDoc;
        private List<FCommercial> myObjectNbreTypeDoc;

        private ComFiltre FiltreObject;

        private List<FCommercial> myObjectTopFamMtant;
        private List<FCommercial> myObjectTopFamCentralisatriceMtant;

        private List<FCommercial> myObjectTopFamQtite;
        private List<FCommercial> myObjectTopFamCentralisatriceQtite;
        private List<FCommercial> myObjectRankingCom;

        private bool TestIsOnlyFact;


        public FenStatistiques(List<FCommercial> ListcomStat, List<FCommercial> LcomStatTypeDoc, List<FCommercial> LcomStatNbreDoc, List<FCommercial> LcomStatNbreTypeDocCom,ComFiltre Cfiltre, List<FCommercial> LcomTopFamMtant, List<FCommercial> LcomTopFamCentralisatriceMtant, List<FCommercial> LcomTopFamQtite,List<FCommercial> LcomTopFamCentralisatriceQtite, List<FCommercial> LcomRanking,bool IsOnlyFact)
        {
            InitializeComponent();

            this.myObject = ListcomStat;
            this.myObjectTypeDoc = LcomStatTypeDoc;
            this.myObjectNbreDoc = LcomStatNbreDoc;
            this.myObjectNbreTypeDoc = LcomStatNbreTypeDocCom;
            this.myObjectTopFamMtant = LcomTopFamMtant;
            this.myObjectTopFamCentralisatriceMtant = LcomTopFamCentralisatriceMtant;
            this.myObjectTopFamQtite = LcomTopFamQtite;
            this.myObjectTopFamCentralisatriceQtite=LcomTopFamCentralisatriceQtite;
            this.myObjectRankingCom = LcomRanking;

            this.FiltreObject = Cfiltre;

            this.TestIsOnlyFact = IsOnlyFact;
        }

        private void FenStatistiques_Load(object sender, EventArgs e)
        {
            try
            {

                // Create an empty chart.
                //   ChartControl sideBySideBarChart = new ChartControl();

                List<FCommercial> LSom = new List<FCommercial>();

                //sommer par date et par commerciaux le revenu de ses ventes

                //Liste des commerciaux(Nom +" "+Prenoms)

                #region Liste des Commerciaux

                Dictionary<string, string> ListeCommerciaux = new Dictionary<string, string>();

                foreach (var item in myObject)
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

                //liste des type de documents

                #region Type Documents
                List<string> ListTypeDoc = new List<string>();


                if (FiltreObject.mIsPipe == false)
                {
                    foreach (var elt in myObjectTypeDoc)
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

                    foreach (var elt in myObjectTypeDoc)
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

                if (ListeCommerciaux.Count > 0)
                {
                    #region chartControlCom
                    foreach (var item in ListeCommerciaux)
                    {

                        List<FCommercial> Ltemp = new List<FCommercial>();

                        Ltemp = myObject.Where(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value).ToList();

                        if (Ltemp != null)
                        {
                            Series ser = new Series(item.Key + " " + item.Value, ViewType.Bar);

                            if (chartControlCom.Series.Count > 0)
                            {
                                if (!chartControlCom.Series[0].Name.Equals(ser.Name))
                                {
                                    foreach (var obj in Ltemp)
                                    {
                                        ser.Points.Add(new SeriesPoint(obj.mDatePiece, obj.mMontantHT.ToString("n0")));
                                    }

                                    ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                    //tester que la serie existe ou pas dans le graphe
                                    chartControlCom.Series.Add(ser);
                                    //Annotations dans la legende
                                    //AxisLabel TextPattern = "{}{V:#,##0}"

                                    ((XYDiagram)chartControlCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                    //ser.CrosshairLabelPattern = "{V:n0}";
                                }
                            }
                            else
                            {
                                foreach (var obj in Ltemp)
                                {
                                    ser.Points.Add(new SeriesPoint(obj.mDatePiece, obj.mMontantHT.ToString("n0")));
                                }

                                ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                chartControlCom.Series.Add(ser);
                                //Annotations dans la legende
                                // ser.CrosshairLabelPattern = "{V:n0}";
                                ((XYDiagram)chartControlCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";
                            }

                        }

                    }

                    #endregion

                }
                
                //graphe type de document

                if (1<0)
                {

                    #region chartControlTypeDoc
                    if (FiltreObject.mIsPipe == false)
                    {
                        foreach (var it in ListTypeDoc)
                        {
                            Series ser = new Series(it, ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlTypeDoc.Series.Add(ser);

                            ((XYDiagram)chartControlTypeDoc.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            foreach (var item in ListeCommerciaux)

                            {

                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                FCommercial LtempTypeDoc = new FCommercial();

                                LtempTypeDoc = myObjectTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mTypeDoc == it);

                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlTypeDoc.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));
                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));

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

                            chartControlTypeDoc.Series.Add(ser);

                            ((XYDiagram)chartControlTypeDoc.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            foreach (var item in ListeCommerciaux)

                            {

                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                FCommercial LtempTypeDoc = new FCommercial();

                                LtempTypeDoc = myObjectTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mStatut == it);

                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlTypeDoc.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));
                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));

                                        //}
                                        //chartControlTypeDoc.Series.Add(ser);
                                    }



                                }

                            }

                        }


                    }

                    #endregion


                }

                if(ListeCommerciaux.Count>0)
                {
                    #region chartControlDateNbDocCom
                    foreach (var item in ListeCommerciaux)

                    {

                        List<FCommercial> Ltemp = new List<FCommercial>();

                        Ltemp = myObjectNbreDoc.Where(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value).ToList();

                        if (Ltemp != null)
                        {
                            Series ser = new Series(item.Key + " " + item.Value, ViewType.Bar);

                            if (chartControlDateNbDocCom.Series.Count > 0)
                            {
                                if (!chartControlDateNbDocCom.Series[0].Name.Equals(ser.Name))
                                {
                                    foreach (var obj in Ltemp)
                                    {
                                        ser.Points.Add(new SeriesPoint(obj.mDatePiece, obj.mNbreDoc));
                                    }

                                    ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                    //tester que la serie existe ou pas dans le graphe
                                    chartControlDateNbDocCom.Series.Add(ser);


                                    ((XYDiagram)chartControlDateNbDocCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                                   // XYDiagram diagram = (XYDiagram)chartControlDateNbDocCom.Diagram;

                                    

                                }
                            }
                            else
                            {
                                foreach (var obj in Ltemp)
                                {
                                    ser.Points.Add(new SeriesPoint(obj.mDatePiece, obj.mNbreDoc));
                                }

                                ser.CrosshairLabelPattern = "{S}: {V:#,##0}";

                                chartControlDateNbDocCom.Series.Add(ser);

                                ((XYDiagram)chartControlDateNbDocCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                               
                            }

                        }

                    }

                    #endregion
                }

                
                if (1<0)
                {

                    #region chartControlNbreTypeDocCom
                    if (FiltreObject.mIsPipe == false)
                    {

                        foreach (var it in ListTypeDoc)
                        {
                            Series ser = new Series(it, ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlNbreTypeDocCom.Series.Add(ser);

                            ((XYDiagram)chartControlNbreTypeDocCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            foreach (var item in ListeCommerciaux)

                            {

                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                FCommercial LtempTypeDoc = new FCommercial();

                                LtempTypeDoc = myObjectNbreTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mTypeDoc == it);

                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlNbreTypeDocCom.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mNbreDoc));
                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mNbreDoc));

                                        //}
                                        //chartControlTypeDoc.Series.Add(ser);
                                    }



                                }

                            }

                        }
                    }
                    else
                    {

                        foreach (var it in ListTypeDoc)
                        {
                            Series ser = new Series(it, ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlNbreTypeDocCom.Series.Add(ser);

                            ((XYDiagram)chartControlNbreTypeDocCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            foreach (var item in ListeCommerciaux)

                            {

                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                FCommercial LtempTypeDoc = new FCommercial();

                                LtempTypeDoc = myObjectNbreTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mStatut == it);

                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlNbreTypeDocCom.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mNbreDoc));
                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mNbreDoc));


                                        //}
                                        //chartControlTypeDoc.Series.Add(ser);
                                    }



                                }

                            }

                        }

                    }

                    #endregion
                }
                
                #region Top10Famille Montant

                if (myObjectTopFamMtant != null && myObjectTopFamMtant.Count > 0)
                {
                    // Create a pie series.
                    Series SeriesFamMtant = new Series("Top 10 des Familles en Montant", ViewType.Pie);
                    
                    //Top 10 famille par montant=====================================
                    foreach (var item in myObjectTopFamMtant)

                    {
                        if(item.mMontantHT>0)
                        {
                            SeriesFamMtant.Points.Add(new SeriesPoint(item.mFamille, item.mMontantHT));
                            
                        }
                        

                    }
                    


                    // Add the series to the chart.
                    chartControlTopFamMtant.Series.Add(SeriesFamMtant);

                    // Format the the series labels(Annotations sur le graphique)
                    SeriesFamMtant.Label.TextPattern = "{A}: {VP:p0}";

                    //Annotations dans la legende
                    SeriesFamMtant.LegendTextPattern = "{A}: {V:n0}";

                    // Access the view-type-specific options of the series.
                    PieSeriesView myView = (PieSeriesView)SeriesFamMtant.View;

                    // Show a title for the series.
                    myView.Titles.Add(new SeriesTitle());
                    myView.Titles[0].Text = SeriesFamMtant.Name;

                }

                #endregion


                #region Top10FamilleCentralisatrice Montant

                if (myObjectTopFamCentralisatriceMtant != null && myObjectTopFamCentralisatriceMtant.Count > 0)
                {
                    // Create a pie series.
                    Series SeriesFamCentralisatriceMtant = new Series("Top 10 des Marques en Montant", ViewType.Pie3D);

                    //Top 10 marque par montant=====================================
                    foreach (var item in myObjectTopFamCentralisatriceMtant)

                    {
                        if (item.mMontantHT > 0)
                        {
                            if (item.mFamilleCentral == string.Empty) item.mFamilleCentral = "NO MARK";

                            SeriesFamCentralisatriceMtant.Points.Add(new SeriesPoint(item.mFamilleCentral, item.mMontantHT));
                        }


                    }


                    // Add the series to the chart.
                    chartControlTopFamCentrMtant.Series.Add(SeriesFamCentralisatriceMtant);

                    // Format the the series labels(Annotations sur le graphique)
                    SeriesFamCentralisatriceMtant.Label.TextPattern = "{A}: {VP:p0}";

                    //Annotations dans la legende
                    SeriesFamCentralisatriceMtant.LegendTextPattern = "{A}: {V:n0}";

                    // Access the view-type-specific options of the series.
                    Pie3DSeriesView myView = (Pie3DSeriesView)SeriesFamCentralisatriceMtant.View;

                    // Show a title for the series.
                    myView.Titles.Add(new SeriesTitle());
                    myView.Titles[0].Text = SeriesFamCentralisatriceMtant.Name;

                }

                #endregion


                #region Top10Famille Qtite

                if (myObjectTopFamQtite != null && myObjectTopFamQtite.Count > 0)
                {
                    // Create a pie series.
                    Series SeriesFamQtite = new Series("Top 10 des Familles en Quantité", ViewType.Pie);

                    //Top 10 famille par montant=====================================
                    foreach (var item in myObjectTopFamQtite)

                    {
                        if (item.mQtiteFamille > 0)
                        {
                            SeriesFamQtite.Points.Add(new SeriesPoint(item.mFamille, item.mQtiteFamille));
                        }


                    }


                    // Add the series to the chart.
                    chartControlTopFamQtite.Series.Add(SeriesFamQtite);
                    
                    // Format the the series labels(Annotations sur le graphique)
                    SeriesFamQtite.Label.TextPattern = "{A}: {VP:p0}";

                    //Annotations dans la legende
                    SeriesFamQtite.LegendTextPattern = "{A}: {V:n0}";

                    // Access the view-type-specific options of the series.
                    PieSeriesView myViewQte = (PieSeriesView)SeriesFamQtite.View;

                    // Show a title for the series.
                    myViewQte.Titles.Add(new SeriesTitle());
                    myViewQte.Titles[0].Text = SeriesFamQtite.Name;

                  

                    }

                #endregion


                #region Top10Famillecentralisatrice Qtite

                if (myObjectTopFamCentralisatriceQtite != null && myObjectTopFamCentralisatriceQtite.Count > 0)
                {
                    // Create a pie series.
                    Series SeriesFamCentrQtite = new Series("Top 10 des Marques en Quantité", ViewType.Pie3D);

                    //Top 10 famille par montant=====================================
                    foreach (var item in myObjectTopFamCentralisatriceQtite)

                    {
                        if (item.mQtiteFamille > 0)
                        {
                            if (item.mFamilleCentral == string.Empty) item.mFamilleCentral = "NO MARK";
                            SeriesFamCentrQtite.Points.Add(new SeriesPoint(item.mFamilleCentral, item.mQtiteFamille));
                        }


                    }


                    // Add the series to the chart.
                    chartControlTopFamCentrQtite.Series.Add(SeriesFamCentrQtite);

                    // Format the the series labels(Annotations sur le graphique)
                    SeriesFamCentrQtite.Label.TextPattern = "{A}: {VP:p0}";

                    //Annotations dans la legende
                    SeriesFamCentrQtite.LegendTextPattern = "{A}: {V:n0}";

                    // Access the view-type-specific options of the series.
                    Pie3DSeriesView myViewQte = (Pie3DSeriesView)SeriesFamCentrQtite.View;

                    // Show a title for the series.
                    myViewQte.Titles.Add(new SeriesTitle());
                    myViewQte.Titles[0].Text = SeriesFamCentrQtite.Name;



                }

                #endregion

                
                #region chartControlRankingCom

                if (myObjectRankingCom != null && myObjectRankingCom.Count > 0)
                {
                    // Create a pie series.
                    
                    Series ser = new Series("Classement Commerciaux", ViewType.Bar);

                    //Top 10 famille par montant=====================================
                    foreach (var obj in myObjectRankingCom)

                    {
                        ser.Points.Add(new SeriesPoint(obj.mPrenomCommercial + " " + obj.mNomCommercial, obj.mMontantHT.ToString("n0")));

                    }

                    ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                    // Add the series to the chart.
                    chartControlRankingCom.Series.Add(ser);


                    ((XYDiagram)chartControlRankingCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                    // Rotate the diagram (if necessary).
                    ((XYDiagram)chartControlRankingCom.Diagram).Rotated = true;


                }
                
                #endregion


                //Stat combiné typedoc/Mtant/com et Nbre doc/Com

                //constante d'agrandissement type doc
                int cste = 1;

                int k = 0;

                
                #region CombinechartControlTypeDocNbre

                if(TestIsOnlyFact)//On a que des factures pour le chartcontrol
                {
                    if (FiltreObject.mIsPipe == false)
                    {
                        //foreach (var it in ListTypeDoc)
                        //{
                        Series ser = new Series("Facture[FA+FC]" + " (En Montant)", ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            Series ser1 = new Series("Facture[FA + FC]" + " (En Nbre)", ViewType.Line);

                            ser1.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlTypeDocNbreMtant.Series.Add(ser1);

                            chartControlTypeDocNbreMtant.Series.Add(ser);


                            ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            // Cast the chart's diagram to the XYDiagram type, to access its axes.
                            XYDiagram diag = (XYDiagram)chartControlTypeDocNbreMtant.Diagram;

                            diag.AxisY.Title.Text = "Nombre de Documents";
                            diag.AxisY.Title.Visibility = DevExpress.Utils.DefaultBoolean.True;
                            diag.AxisY.Title.Alignment = StringAlignment.Center;
                            diag.AxisY.Title.TextColor = Color.Black;

                            SecondaryAxisY myAxisY = new SecondaryAxisY("my Y-Axis");

                            ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).SecondaryAxesY.Add(myAxisY);

                        //// Assign the series2 to the created axes.

                        ////  ((BarSeriesView)ser1.View).AxisY = myAxisY;
                        ((BarSeriesView)ser.View).AxisY = ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).SecondaryAxesY[0]; ;

                        ((BarSeriesView)ser.View).AxisY.Label.TextPattern = "{V:#,##0}";

                        if (k == 0)//Ajouter qu'une seule fois
                        {
                            //// Customize the appearance of the secondary axes (optional).
                            myAxisY.Title.Text = "Montant ";
                            myAxisY.Title.Visibility = DevExpress.Utils.DefaultBoolean.True;
                            myAxisY.Title.TextColor = Color.DarkBlue;
                            myAxisY.Label.TextColor = Color.DarkBlue;
                            myAxisY.Color = Color.DarkBlue;

                            k += 1;
                        }
                        else
                        {
                            myAxisY.Visibility = DevExpress.Utils.DefaultBoolean.False;
                        }


                        foreach (var item in ListeCommerciaux)
                            {
                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                List<FCommercial> LtempTypeDoc = new List<FCommercial>();
                                LtempTypeDoc = myObjectTypeDoc.Where(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value).ToList();

                                double SommeTotal = 0;

                                  int qtite = 0;

                                 foreach(var obj in LtempTypeDoc)
                                {
                                    SommeTotal += obj.mMontantHT;
                                
                                }

                            List<FCommercial> LtempTypeDoc1 = new List<FCommercial>();

                            LtempTypeDoc1 = myObjectNbreTypeDoc.Where(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value).ToList() ;

                            //foreach (var obj in LtempTypeDoc1)
                            //{
                            //    qtite += obj.mNbreDoc;
                            //}

                            //Modif car on tenait compte des documents en double++++++++++++++++++

                            var premChoix = LtempTypeDoc1.First();

                            qtite = premChoix.mNbreDoc;

                            //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                            FCommercial LtempCom = new FCommercial();

                            LtempCom = myObjectNbreTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value);


                            //FCommercial LtempTypeDoc1 = new FCommercial();

                            //    LtempTypeDoc1 = myObjectNbreTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mTypeDoc == it);


                            if (LtempTypeDoc != null)
                                {

                                    if (chartControlTypeDocNbreMtant.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempCom.mNomCommercial + " " + LtempCom.mPrenomCommercial, SommeTotal));

                                        ser1.Points.Add(new SeriesPoint(LtempCom.mNomCommercial + " " + LtempCom.mPrenomCommercial, qtite));

                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {

                                    ser.Points.Add(new SeriesPoint(LtempCom.mNomCommercial + " " + LtempCom.mPrenomCommercial, SommeTotal));

                                    ser1.Points.Add(new SeriesPoint(LtempCom.mNomCommercial + " " + LtempCom.mPrenomCommercial, qtite));

                                    //foreach (var obj in LtempTypeDoc)
                                    //{
                                    //ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));

                                    //ser1.Points.Add(new SeriesPoint(LtempTypeDoc1.mNomCommercial + " " + LtempTypeDoc1.mPrenomCommercial, LtempTypeDoc1.mNbreDoc));

                                    //}
                                    //chartControlTypeDoc.Series.Add(ser);
                                }



                            }

                            }

                       // }

                    }
                    else
                    {
                        //cas du pipe

                        foreach (var it in ListTypeDoc)
                        {
                            Series ser = new Series(it + " (En Montant)", ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            Series ser1 = new Series(it + " (En Nbre)", ViewType.Line);

                            ser1.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlTypeDocNbreMtant.Series.Add(ser1);

                            chartControlTypeDocNbreMtant.Series.Add(ser);

                            ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            // Cast the chart's diagram to the XYDiagram type, to access its axes.
                            XYDiagram diag = (XYDiagram)chartControlTypeDocNbreMtant.Diagram;

                            diag.AxisY.Title.Text = "Nombre de Documents";
                            diag.AxisY.Title.Visibility = DevExpress.Utils.DefaultBoolean.True;
                            diag.AxisY.Title.Alignment = StringAlignment.Center;
                            diag.AxisY.Title.TextColor = Color.Black;

                            // Create two secondary axes, and add them to the chart's Diagram.

                            SecondaryAxisY myAxisY = new SecondaryAxisY("my Y-Axis");

                            ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).SecondaryAxesY.Add(myAxisY);

                            //  Assign the series2 to the created axes.

                            ((BarSeriesView)ser.View).AxisY = ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).SecondaryAxesY[0]; ;

                            ((BarSeriesView)ser.View).AxisY.Label.TextPattern = "{V:#,##0}";

                            if (k == 0)//Ajouter qu'une seule fois
                            {
                                //// Customize the appearance of the secondary axes (optional).
                                myAxisY.Title.Text = "Montant";
                                myAxisY.Title.Visibility = DevExpress.Utils.DefaultBoolean.True;
                                myAxisY.Title.TextColor = Color.DarkBlue;
                                myAxisY.Label.TextColor = Color.DarkBlue;
                                myAxisY.Color = Color.DarkBlue;

                                k += 1;
                            }
                            else
                            {
                                myAxisY.Visibility = DevExpress.Utils.DefaultBoolean.False;
                            }

                            foreach (var item in ListeCommerciaux)

                            {

                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                FCommercial LtempTypeDoc = new FCommercial();

                                LtempTypeDoc = myObjectTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mStatut == it);

                                FCommercial LtempTypeDoc1 = new FCommercial();

                                LtempTypeDoc1 = myObjectNbreTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mStatut == it);


                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlTypeDocNbreMtant.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));

                                        ser1.Points.Add(new SeriesPoint(LtempTypeDoc1.mNomCommercial + " " + LtempTypeDoc1.mPrenomCommercial, LtempTypeDoc1.mNbreDoc));
                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));

                                        ser1.Points.Add(new SeriesPoint(LtempTypeDoc1.mNomCommercial + " " + LtempTypeDoc1.mPrenomCommercial, LtempTypeDoc1.mNbreDoc));

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
                            Series ser = new Series(it + " (En Montant)", ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            Series ser1 = new Series(it + " (En Nbre)", ViewType.Line);

                            ser1.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlTypeDocNbreMtant.Series.Add(ser1);

                            chartControlTypeDocNbreMtant.Series.Add(ser);


                            ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            // Cast the chart's diagram to the XYDiagram type, to access its axes.
                            XYDiagram diag = (XYDiagram)chartControlTypeDocNbreMtant.Diagram;

                            diag.AxisY.Title.Text = "Nombre de Documents";
                            diag.AxisY.Title.Visibility = DevExpress.Utils.DefaultBoolean.True;
                            diag.AxisY.Title.Alignment = StringAlignment.Center;
                            diag.AxisY.Title.TextColor = Color.Black;

                            SecondaryAxisY myAxisY = new SecondaryAxisY("my Y-Axis");

                            ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).SecondaryAxesY.Add(myAxisY);

                            // Assign the series2 to the created axes.

                            //  ((BarSeriesView)ser1.View).AxisY = myAxisY;
                            ((BarSeriesView)ser.View).AxisY = ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).SecondaryAxesY[0]; ;

                            ((BarSeriesView)ser.View).AxisY.Label.TextPattern = "{V:#,##0}";

                            if (k == 0)//Ajouter qu'une seule fois
                            {
                                //// Customize the appearance of the secondary axes (optional).
                                myAxisY.Title.Text = "Montant";
                                myAxisY.Title.Visibility = DevExpress.Utils.DefaultBoolean.True;
                                myAxisY.Title.TextColor = Color.DarkBlue;
                                myAxisY.Label.TextColor = Color.DarkBlue;
                                myAxisY.Color = Color.DarkBlue;

                                k += 1;
                            }
                            else
                            {
                                myAxisY.Visibility = DevExpress.Utils.DefaultBoolean.False;
                            }


                            foreach (var item in ListeCommerciaux)

                            {
                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                FCommercial LtempTypeDoc = new FCommercial();

                                LtempTypeDoc = myObjectTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mTypeDoc == it);

                                FCommercial LtempTypeDoc1 = new FCommercial();

                                LtempTypeDoc1 = myObjectNbreTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mTypeDoc == it);


                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlTypeDocNbreMtant.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));

                                        ser1.Points.Add(new SeriesPoint(LtempTypeDoc1.mNomCommercial + " " + LtempTypeDoc1.mPrenomCommercial, LtempTypeDoc1.mNbreDoc));

                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));

                                        ser1.Points.Add(new SeriesPoint(LtempTypeDoc1.mNomCommercial + " " + LtempTypeDoc1.mPrenomCommercial, LtempTypeDoc1.mNbreDoc));

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
                            Series ser = new Series(it + " (En Montant)", ViewType.Bar);

                            ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            Series ser1 = new Series(it + " (En Nbre)", ViewType.Line);

                            ser1.CrosshairLabelPattern = "{A}: {V:#,##0}";

                            chartControlTypeDocNbreMtant.Series.Add(ser1);

                            chartControlTypeDocNbreMtant.Series.Add(ser);

                            ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                            // Cast the chart's diagram to the XYDiagram type, to access its axes.
                            XYDiagram diag = (XYDiagram)chartControlTypeDocNbreMtant.Diagram;

                            diag.AxisY.Title.Text = "Nombre de Documents";
                            diag.AxisY.Title.Visibility = DevExpress.Utils.DefaultBoolean.True;
                            diag.AxisY.Title.Alignment = StringAlignment.Center;
                            diag.AxisY.Title.TextColor = Color.Black;

                            // Create two secondary axes, and add them to the chart's Diagram.

                            SecondaryAxisY myAxisY = new SecondaryAxisY("my Y-Axis");

                            ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).SecondaryAxesY.Add(myAxisY);

                            //  Assign the series2 to the created axes.

                            ((BarSeriesView)ser.View).AxisY = ((XYDiagram)chartControlTypeDocNbreMtant.Diagram).SecondaryAxesY[0]; ;

                            ((BarSeriesView)ser.View).AxisY.Label.TextPattern = "{V:#,##0}";

                            if (k == 0)//Ajouter qu'une seule fois
                            {
                                //// Customize the appearance of the secondary axes (optional).
                                myAxisY.Title.Text = "Montant";
                                myAxisY.Title.Visibility = DevExpress.Utils.DefaultBoolean.True;
                                myAxisY.Title.TextColor = Color.DarkBlue;
                                myAxisY.Label.TextColor = Color.DarkBlue;
                                myAxisY.Color = Color.DarkBlue;

                                k += 1;
                            }
                            else
                            {
                                myAxisY.Visibility = DevExpress.Utils.DefaultBoolean.False;
                            }

                            foreach (var item in ListeCommerciaux)

                            {

                                //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                                FCommercial LtempTypeDoc = new FCommercial();

                                LtempTypeDoc = myObjectTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mStatut == it);

                                FCommercial LtempTypeDoc1 = new FCommercial();

                                LtempTypeDoc1 = myObjectNbreTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mStatut == it);


                                if (LtempTypeDoc != null)
                                {

                                    if (chartControlTypeDocNbreMtant.Series.Count > 0)
                                    {
                                        //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                                        //{
                                        //foreach (var obj in LtempTypeDoc)
                                        //{

                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));

                                        ser1.Points.Add(new SeriesPoint(LtempTypeDoc1.mNomCommercial + " " + LtempTypeDoc1.mPrenomCommercial, LtempTypeDoc1.mNbreDoc));
                                        // }

                                        //tester que la serie existe ou pas dans le graphe
                                        //chartControlTypeDoc.Series.Add(ser);
                                        //  }
                                    }
                                    else
                                    {
                                        //foreach (var obj in LtempTypeDoc)
                                        //{
                                        ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mMontantHT));

                                        ser1.Points.Add(new SeriesPoint(LtempTypeDoc1.mNomCommercial + " " + LtempTypeDoc1.mPrenomCommercial, LtempTypeDoc1.mNbreDoc));

                                        //}
                                        //chartControlTypeDoc.Series.Add(ser);
                                    }



                                }

                            }

                        }


                    }

                }



                //if (FiltreObject.mIsPipe == false)
                //{

                //    foreach (var it in ListTypeDoc)
                //    {
                //        Series ser = new Series(it, ViewType.Bar);

                //        ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                //        chartControlNbreTypeDocCom.Series.Add(ser);

                //        ((XYDiagram)chartControlNbreTypeDocCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                //        foreach (var item in ListeCommerciaux)

                //        {

                //            //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                //            FCommercial LtempTypeDoc = new FCommercial();

                //            LtempTypeDoc = myObjectNbreTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mTypeDoc == it);

                //            if (LtempTypeDoc != null)
                //            {

                //                if (chartControlNbreTypeDocCom.Series.Count > 0)
                //                {
                //                    //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                //                    //{
                //                    //foreach (var obj in LtempTypeDoc)
                //                    //{

                //                    ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mNbreDoc));


                //                    // }

                //                    //tester que la serie existe ou pas dans le graphe
                //                    //chartControlTypeDoc.Series.Add(ser);
                //                    //  }
                //                }
                //                else
                //                {
                //                    //foreach (var obj in LtempTypeDoc)
                //                    //{
                //                    ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mNbreDoc));


                //                    //}
                //                    //chartControlTypeDoc.Series.Add(ser);
                //                }



                //            }

                //        }

                //    }
                //}
                //else
                //{

                //    foreach (var it in ListTypeDoc)
                //    {
                //        Series ser = new Series(it, ViewType.Bar);

                //        ser.CrosshairLabelPattern = "{A}: {V:#,##0}";

                //        chartControlNbreTypeDocCom.Series.Add(ser);

                //        ((XYDiagram)chartControlNbreTypeDocCom.Diagram).AxisY.Label.TextPattern = "{V:#,##0}";

                //        foreach (var item in ListeCommerciaux)

                //        {

                //            //List <FCommercial> LtempTypeDoc = new List<FCommercial>();
                //            FCommercial LtempTypeDoc = new FCommercial();

                //            LtempTypeDoc = myObjectNbreTypeDoc.FirstOrDefault(c => c.mNomCommercial == item.Key && c.mPrenomCommercial == item.Value && c.mStatut == it);

                //            if (LtempTypeDoc != null)
                //            {

                //                if (chartControlNbreTypeDocCom.Series.Count > 0)
                //                {
                //                    //if (!chartControlTypeDoc.Series[0].Name.Equals(ser.Name))
                //                    //{
                //                    //foreach (var obj in LtempTypeDoc)
                //                    //{

                //                    ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mNbreDoc));
                //                    // }

                //                    //tester que la serie existe ou pas dans le graphe
                //                    //chartControlTypeDoc.Series.Add(ser);
                //                    //  }
                //                }
                //                else
                //                {
                //                    //foreach (var obj in LtempTypeDoc)
                //                    //{
                //                    ser.Points.Add(new SeriesPoint(LtempTypeDoc.mNomCommercial + " " + LtempTypeDoc.mPrenomCommercial, LtempTypeDoc.mNbreDoc));


                //                    //}
                //                    //chartControlTypeDoc.Series.Add(ser);
                //                }



                //            }

                //        }

                //    }

                //}



                #endregion

                //==============================================================================

                chartControlRankingCom.Legend.Visibility = DevExpress.Utils.DefaultBoolean.False;

                    chartControlTopFamQtite.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;
                    chartControlTopFamMtant.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;

                    chartControlDateNbDocCom.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;

                    //// Hide the legend (if necessary).
                    chartControlTypeDoc.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;

                    chartControlCom.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;

                    chartControlNbreTypeDocCom.Legend.Visibility = DevExpress.Utils.DefaultBoolean.True;
                
            }
            catch (Exception ex)
            {
                //if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                var msg = "FenStatistiques -> FenStatistiques_Load -> TypeErreur: " + ex.Message; ;
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


        private void BtnExportExcelStat_Click(object sender, EventArgs e)
        {
            try
            {
                string cheminFichier = string.Empty;

                string Titre = string.Empty;
                
                int processIdAvant = 0;

                int PidExcelproc = 0;

                string chaineBABI = string.Empty;

                //On a des eléments ,on peut imprimer dans excel

                if (sFDExcelStat.ShowDialog() == DialogResult.OK)
                {
                    cheminFichier = sFDExcelStat.FileName;

                    if (cheminFichier.Length >= 203)
                    {
                        //chemin trop long

                        MessageBox.Show("Le chemin spécifié est trop long !Veuillez choisir un autre dossier!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    if (!splashScreenManager1.IsSplashFormVisible) splashScreenManager1.ShowWaitForm();

                    if (xtraTabControlStat.SelectedTabPage==xtraTabPageCom)
                    {
                        chartControlCom.ExportToXlsx(cheminFichier);

                        Titre = "Statistiques Commerciaux Par Date et Montant";
                    }
                    if (xtraTabControlStat.SelectedTabPage == xtraTabPageTypeDoc)
                    {
                        chartControlTypeDoc.ExportToXlsx(cheminFichier);

                        Titre = "Statistiques Commerciaux Par Type de document et Montant ";
                    }

                    if (xtraTabControlStat.SelectedTabPage == xtraTabPageDateNbDocCom)
                    {
                        chartControlDateNbDocCom.ExportToXlsx(cheminFichier);

                        Titre = "Statistiques Commerciaux Par Date et Nombre Document";
                    }

                    if (xtraTabControlStat.SelectedTabPage == xtraTabPageNbreTypeDocCom)
                    {
                        chartControlNbreTypeDocCom.ExportToXlsx(cheminFichier);

                        Titre = "Statistiques Commerciaux Par Nombre de Type de document ";
                    }
                    if (xtraTabControlStat.SelectedTabPage == xtraTabPageTopFamMtant)
                    {
                        chartControlTopFamMtant.ExportToXlsx(cheminFichier);
                      
                        Titre = "Top 10 des Familles en Montant ";
                    }
                    if (xtraTabControlStat.SelectedTabPage == xtraTabPageTopFamQte)
                    {
                        chartControlTopFamQtite.ExportToXlsx(cheminFichier);

                        Titre = "Top 10 des Familles en Quantité ";
                    }

                    if (xtraTabControlStat.SelectedTabPage == xtraTabPageRankingCom)
                    {
                        chartControlRankingCom.ExportToXlsx(cheminFichier);

                        Titre = "Classement Commerciaux ";
                    }

                    if (xtraTabControlStat.SelectedTabPage == xtraTabPageTypeDocNbreMtant)
                    {
                        chartControlTypeDocNbreMtant.ExportToXlsx(cheminFichier);

                        Titre = "Statistiques Commerciaux Par Montant et Nombre de Documents";
                    }

                    if (xtraTabControlStat.SelectedTabPage == xtraTabPageTopFamCentrMtant)
                    {
                        chartControlTopFamCentrMtant.ExportToXlsx(cheminFichier);

                        Titre = "Top 10 des Marques en Montant ";
                    }

                    if (xtraTabControlStat.SelectedTabPage == xtraTabPageTopFamCentrQtite)
                    {
                        chartControlTopFamCentrQtite.ExportToXlsx(cheminFichier);

                        Titre = "Top 10 des Marques en Quantité ";
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
                    if (isImputsOK) isImputsOK = FormatMain.EcrireDansExcel(xlsApp, ligne+1, colonne, FiltreObject.mSite);
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
            catch(Exception ex)
            {
                if (splashScreenManager1.IsSplashFormVisible) splashScreenManager1.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenStatistiques ->BtnExportExcelStat_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }
        
    }
}
