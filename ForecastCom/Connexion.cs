using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ForecastCom.Utils;
using ForecastCom.Services;

namespace ForecastCom
{
    public partial class Connexion : Form
    {
        private readonly CAlias daoMain = new CAlias();
        private List<CAlias> ListeFiliales = new List<CAlias>();
        private readonly FormatageForecastCom FormatMain = new FormatageForecastCom();

        public int id = 0;

        public int idModif = 0;

        public static CModif mod = new CModif();


        public Connexion()
        {
            InitializeComponent();
        }

        private void sBtnAjoutCon_Click(object sender, EventArgs e)
        {
            try
            {
                CModif cl = new CModif();

                cl.mIsModif = false;
                cl.mIdModif = 0;

                FenAlias FAlias = new FenAlias();

                FAlias.ShowDialog();

                //Remplir la liste des filiales
                ListeFiliales = daoMain.GetAliasConnexions();

                gridConnexion.DataSource = ListeFiliales;
            }
            catch (Exception ex)
            {
                //  MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "REPORTING HPE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Connexion ->sBtnAjoutCon_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

        private void sBtnModifCon_Click(object sender, EventArgs e)
        {
            try
            {
                if (gridView1.DataRowCount != 0)
                {

                    var cid = gridView1.GetRowCellValue(gridView1.FocusedRowHandle, "mId").ToString();

                    idModif = Convert.ToInt32(cid);

                    if (idModif > 0)
                    {
                        CModif cl = new CModif();

                        cl.mIsModif = true;
                        cl.mIdModif = idModif;

                        FenAlias FAliasModif = new FenAlias(cl);

                        FAliasModif.ShowDialog();

                        //Remplir la liste des filiales
                        ListeFiliales = daoMain.GetAliasConnexions();

                        gridConnexion.DataSource = ListeFiliales;

                    }

                }

            }
            catch (Exception ex)
            {
                //  MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "REPORTING HPE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Connexion ->sBtnAjoutCon_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

        private void sBtnSupprCon_Click(object sender, EventArgs e)
        {
            var villeTodelete = string.Empty;
            try
            {

                if (gridView1.DataRowCount != 0)
                {

                    var cid = gridView1.GetRowCellValue(gridView1.FocusedRowHandle, "mId").ToString();

                    id = Convert.ToInt32(cid);

                    if (id > 0 && ListeFiliales.Count > 0)
                    {
                        var AliastoDelete = ListeFiliales.Where(c => c.mId == id).ToList();

                        if (AliastoDelete != null)
                        {
                            foreach (var item in AliastoDelete)
                            {
                                villeTodelete = item.mVille;

                            }

                        }

                        if (villeTodelete != string.Empty)
                        {
                            var res = MessageBox.Show("Voulez-vous supprimer la connexion configurée pour " + villeTodelete + " ?", "FORECASTCOM", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (res == DialogResult.Yes)
                            {
                                var strConnexion = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ressources/" + "DbConnexion.ini");

                                if (File.Exists(strConnexion)) File.Delete(strConnexion);

                                //recupérer une liste sans la ville a supprimer

                                var ret = false;

                                ListeFiliales = ListeFiliales.Where(c => c.mId != id).ToList();

                                if (ListeFiliales.Count > 0)
                                {
                                    var LesAlias = new List<string>();


                                    //Mettre le contenu de ListeFiliales dans une liste de string

                                    foreach (var item in ListeFiliales)
                                    {
                                        var text = item.mId.ToString() + "#" + item.mAliasName + "#" + item.mVille;

                                        LesAlias.Add(text);
                                    }

                                    ret = FormatMain.EcrireDansFichier(LesAlias, strConnexion);

                                    if (ret) MessageBox.Show("L'alias a été supprimé avec succès !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    //Rafraichir la grid

                                    gridConnexion.DataSource = ListeFiliales;
                                }
                                else//On supprime le dernier élément de la grid
                                {

                                    MessageBox.Show("L'alias a été supprimé avec succès !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    //Rafraichir la grid

                                    gridConnexion.DataSource = ListeFiliales;
                                }
                            }

                        }

                    }

                }

            }
            catch (Exception ex)
            {
                //  MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "REPORTING HPE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Connexion ->sBtnSupprCon_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

        private void Connexion_Load(object sender, EventArgs e)
        {
            //Remplir la liste des filiales
            ListeFiliales = daoMain.GetAliasConnexions();

            gridConnexion.DataSource = ListeFiliales;
        }
    }
}
