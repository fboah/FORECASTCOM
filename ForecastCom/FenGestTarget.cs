//using ForecastCom.ConnexionAdmin;
using ForecastCom.DAO;
using ForecastCom.Models;
using ForecastCom.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ForecastCom.ConnTarget;


namespace ForecastCom
{
    public partial class FenGestTarget : Form
    {

        private List<ComClass> myObjectCom;
        private List<ComFamille> myObjectFam;
        private List<CAlias> myObjectPays;

        private readonly DAOForecastCom daoReport = new DAOForecastCom();

        public FenGestTarget(List<ComClass>LCom, List<ComFamille> LFam, List<CAlias> ListeFilialeChoisies)
        {
            InitializeComponent();
            this.myObjectCom = LCom;
            this.myObjectFam= LFam;
            this.myObjectPays = ListeFilialeChoisies;
        }

        private void FenGestTarget_Load(object sender, EventArgs e)
        {
            try
            {
                //charger les comboBox

                LueCommercial.Properties.DataSource = myObjectCom;
                LueFamille.Properties.DataSource = myObjectFam;

                LuePays.Properties.DataSource = myObjectPays;


                LueCommercial.Properties.DisplayMember = "mNomCommercial";
                LueFamille.Properties.DisplayMember = "mFa_CodeFamille";

                LuePays.Properties.DisplayMember = "mVille";
                LuePays.Properties.ValueMember = "mId";

                //Choisir les premières valeurs
                LueCommercial.ItemIndex = 0;
                LueFamille.ItemIndex = 0;
                LuePays.ItemIndex = 0;

                //Initialiser Année

                sNumAnnee.EditValue = DateTime.Now.Year;

            }
            catch(Exception ex)
            {
              //  if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenGestTarget ->FenGestTarget_Load -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

        private void LuePays_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                ////Liste utilisée pour la génération du Report
                //myObjectPays.Clear();

                List<String> Lidpays = new List<string>();
                List<ComClass> LComboCom = new List<ComClass>();
                List<ComFamille> LComboFam = new List<ComFamille>();
                

                var id = LuePays.EditValue.ToString();

                Lidpays.Add(id);

                if(Lidpays!=null)
                {
                    //Remplir le combo commerciaux d'apres les pays choisis=============================
                    LComboCom = daoReport.GetListByListIdCom(Lidpays, myObjectCom);

                    LueCommercial.Properties.DataSource = LComboCom;

                    //Choisir les premières valeurs
                    LueCommercial.ItemIndex = 0;

                    //Remplir Combo famille===========================================================

                    LComboFam = daoReport.GetListByListIdFam(Lidpays, myObjectFam);

                    LueFamille.Properties.DataSource = LComboFam;

                    //Choisir les premières valeurs
                    LueFamille.ItemIndex = 0;
                }
               
                
            }
            catch (Exception ex)
            {
                //if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenGestTarget ->LuePays_EditValueChanged -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

        private void connexionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //if (!splashScreenManager2.IsSplashFormVisible) splashScreenManager2.ShowWaitForm();
                var FenConnexionTarget = new ConAdmin();

                //  if (splashScreenManager2.IsSplashFormVisible) splashScreenManager2.CloseWaitForm();

                FenConnexionTarget.ShowDialog();
            }
            catch(Exception ex)
            {

            }
        }

        private void sNumAnnuelCom_EditValueChanged(object sender, EventArgs e)
        {
            try
            {
                sNumTrimCom.Value = sNumAnnuelCom.Value / 4;

                sNumMensuelCom.Value = sNumAnnuelCom.Value / 12;

                sNumHebdoCom.Value = sNumAnnuelCom.Value / 52;

            }
            catch(Exception ex)
            {

            }
        }
    }
}
