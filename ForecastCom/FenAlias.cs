using ForecastCom.Services;
using ForecastCom.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ForecastCom
{
    public partial class FenAlias : Form
    {
        private readonly FormatageForecastCom FormatMain = new FormatageForecastCom();

        private readonly CAlias daoMain = new CAlias();
        private List<CAlias> ListeFiliales = new List<CAlias>();
        private int idModif;
        private CModif cl;


        public FenAlias()
        {
            InitializeComponent();
        }

        public FenAlias(CModif cl)
        {
            InitializeComponent();
            this.cl = cl;
        }


        public void FillComboInstances()
        {
            if (CmbServeur.Items.Count == 0)
            {
                Cursor.Current = Cursors.WaitCursor;

                var oTable = SqlDataSourceEnumerator.Instance.GetDataSources();

                foreach (DataRow oRow in oTable.Rows)
                {
                    if (oRow["InstanceName"].ToString() == string.Empty)
                    {
                        CmbServeur.Items.Add(oRow["ServerName"]);
                    }
                    else
                    {
                        CmbServeur.Items.Add(oRow["ServerName"].ToString() + "\\" + oRow["InstanceName"].ToString());
                    }

                }


                Cursor.Current = Cursors.Default;
            }


        }

        public void FillComboBD()
        {
            var Auth = Convert.ToInt16(CmbAuthentification.EditValue.ToString());

            var Chaineconex = string.Empty;

            var requete = string.Empty;

            var SourceData = new DataTable();

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                Chaineconex = FormatageChaineConnexion(Auth, "Master", CmbServeur.Text, txtUser.Text, txtMotPasse.Text);

                requete = @"SELECT name,dbid FROM master.dbo.sysdatabases where dbid not in (1,2,3,4,5,6) order by name,dbid desc";

                var dta = new SqlDataAdapter(requete, Chaineconex);

                dta.Fill(SourceData);

                CmbBaseDonnees.DataSource = SourceData;

                CmbBaseDonnees.DisplayMember = SourceData.Columns["name"].ToString();

                CmbBaseDonnees.ValueMember = SourceData.Columns["dbid"].ToString();

                Cursor.Current = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur est survenue lors de l'établissement d'une connexion à SQL Server! Veuillez revoir vos paramètres!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAlias ->FillComboBD -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }

        }


        public void FillComboAuth()
        {
            var MySelect = new DataTable();

            MySelect.Columns.Add("Id");

            MySelect.Columns.Add("Libelle");

            MySelect.Rows.Add(0, "Authentification Windows");

            MySelect.Rows.Add(1, "Sql Serveur");

            CmbAuthentification.Properties.DataSource = MySelect;

            CmbAuthentification.Properties.ValueMember = "Id";

            CmbAuthentification.Properties.DisplayMember = "Libelle";

        }


        public string FormatageChaineConnexion(int Auth, string BD, string Serveur, string User, string MDP)
        {
            string ret = string.Empty;
            try
            {

                if (Auth == 1 && BD != string.Empty && Serveur != string.Empty && User != string.Empty && MDP != string.Empty)
                {
                    ret = "Initial Catalog=" + BD + ";Data Source=" + Serveur + ";user=" + User + ";password=" + MDP + " ";
                }

                if (Auth == 0 && BD != string.Empty && Serveur != string.Empty)

                {
                    ret = "Initial Catalog=" + BD + ";Data Source=" + Serveur + ";Integrated Security=SSPI";

                }

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAlias ->FormatageChaineConnexion -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return ret;
            }

        }


        public string FormatageChaineConnexionDbConn(int Auth, string BD, string Serveur, string User, string MDP, int pos, string Ville)
        {
            string ret = string.Empty;
            try
            {

                if (Auth == 1 && BD != string.Empty && Serveur != string.Empty && User != string.Empty && MDP != string.Empty)
                {
                    ret = "Initial Catalog=" + BD + ";Data Source=" + Serveur + ";user=" + User + ";password=" + MDP + " ";
                }

                if (Auth == 0 && BD != string.Empty && Serveur != string.Empty)

                {
                    ret = "Initial Catalog=" + BD + ";Data Source=" + Serveur + ";Integrated Security=SSPI";

                }

                ret = pos.ToString() + "#" + ret + "#" + Ville;

                return ret;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAlias ->FormatageChaineConnexionDbConn -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return ret;
            }

        }


        public Boolean TestConnexion(string conn)
        {
            var ret = false;

            var MaConnex = new SqlConnection();

            MaConnex.ConnectionString = conn;

            try
            {
                MaConnex.Open();
                ret = true;
                return ret;
            }
            catch (Exception ex)
            {
                ret = false;
                MaConnex.Close();
                return ret;
            }

        }


        private void sBtnTesterCon_Click(object sender, EventArgs e)
        {

            var chconn = string.Empty;

            var test = false;

            if (CmbServeur.Text == string.Empty)
            {
                MessageBox.Show("Erreur de connexion.Veuillez vérifier éventuellement que le serveur et/ou la base existent!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            if (CmbBaseDonnees.Text == string.Empty)
            {
                MessageBox.Show("Erreur de connexion.Veuillez vérifier éventuellement que le serveur et/ou la base existent!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
            }

            int Auth = Convert.ToInt16(CmbAuthentification.EditValue.ToString());

            chconn = FormatageChaineConnexion(Auth, CmbBaseDonnees.Text, CmbServeur.Text, txtUser.Text, txtMotPasse.Text);

            if (chconn != string.Empty) test = TestConnexion(chconn);

            if (test)
            {
                MessageBox.Show("Connexion réussie!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {

                MessageBox.Show("Erreur de connexion.Veuillez vérifier éventuellement que le serveur et/ou la base existent!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }


        }

        private void FenAlias_Load(object sender, EventArgs e)
        {
            try
            {

                FillComboAuth();

                //Cas d'une premiere connexion ,forcer à ABIDJAN pour ne pas avoir d'erreur de syntaxe

                if (cl == null)
                {
                    ListeFiliales = daoMain.GetAliasConnexions();

                    if (ListeFiliales.Count == 0)
                    {
                        txtVille.Text = "ABIDJAN";
                        txtVille.Enabled = false;
                    }

                }


                if (cl != null)
                {

                    if (cl.mIdModif > 0 && cl.mIsModif == true)
                    {
                        var idMod = cl.mIdModif;

                        //cas d'une modification

                        ListeFiliales = daoMain.GetAliasConnexions();

                        if (ListeFiliales.Count > 0)
                        {
                            var ObjToModify = ListeFiliales.FirstOrDefault(c => c.mId == idMod);

                            if (ObjToModify != null)
                            {
                                txtVille.Enabled = false;

                                txtVille.Text = ObjToModify.mVille;

                                if (ObjToModify.mAliasName.Contains("Integrated Security=SSPI"))
                                {
                                    CmbAuthentification.ItemIndex = 0;
                                    txtUser.Text = string.Empty;

                                    txtMotPasse.Text = string.Empty;

                                    txtUser.Enabled = false;

                                    txtMotPasse.Enabled = false;

                                }
                                else//SQL AUTH
                                {
                                    CmbAuthentification.ItemIndex = 1;
                                    var chaine = ObjToModify.mAliasName.Split(';');

                                    var chuser = chaine[2].Split('=');

                                    txtUser.Text = chuser[1].ToString();

                                    var chMDP = chaine[3].Split('=');

                                    txtMotPasse.Text = chMDP[1].ToString();

                                    txtUser.Enabled = true;

                                    txtMotPasse.Enabled = true;
                                }

                                //Serveur et BD

                                var line = ObjToModify.mAliasName.Split(';');

                                var chBD = line[0].Split('=');

                                CmbBaseDonnees.Text = chBD[1].ToString();

                                var chServeur = line[1].Split('=');

                                CmbServeur.Text = chServeur[1].ToString();
                            }

                        }

                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAlias ->FenAlias_Load -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

        private void sBtnQuitter_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void sBtnValider_Click(object sender, EventArgs e)
        {
            //Créer le fichier de DBconnexion sil nexiste pas et ecrire dedans

            var ret = false;

            var testconnexion = false;

            var Auth = Convert.ToInt16(CmbAuthentification.EditValue.ToString());

            var Chconn = FormatageChaineConnexion(Auth, CmbBaseDonnees.Text, CmbServeur.Text, txtUser.Text, txtMotPasse.Text);

            testconnexion = TestConnexion(Chconn);

            try
            {
                if (testconnexion)
                {
                    var strConnexion = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ressources/" + "DbConnexion.ini");

                    if (txtVille.Text == string.Empty)
                    {
                        MessageBox.Show("Veuillez renseigner une ville !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    if (!File.Exists(strConnexion))//Ajout premier élément
                    {
                        //La methode Ecrire va créer le fichier
                        //Ajouter Abidjan comme première connexion

                        var Alias = FormatageChaineConnexionDbConn(Auth, CmbBaseDonnees.Text, CmbServeur.Text, txtUser.Text, txtMotPasse.Text, 1, txtVille.Text);

                        var LAlias = new List<string>();

                        LAlias.Add(Alias);

                        ret = FormatMain.EcrireDansFichier(LAlias, strConnexion);

                        if (ret)
                        {
                            MessageBox.Show("Connexion configurée avec succès !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            Close();

                        }

                    }
                    else
                    {
                        // Ajout d'un autre élément'

                        ListeFiliales = daoMain.GetAliasConnexions();

                        if (ListeFiliales.Count > 0)
                        {

                            //Verifier si la ville n'existe pas déjà

                            foreach (var obj in ListeFiliales)
                            {
                                if (obj.mVille.ToUpper() == txtVille.Text.ToUpper())
                                {
                                    if (cl == null)//Ajout
                                    {
                                        MessageBox.Show("Cette ville a été déjà configurée!Veuillez la modifier au besoin.", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                        return;
                                    }


                                }

                            }

                            //Supprimer le fichier existant et le récréer en y ajoutant les nouvelles infos

                            var nbre = ListeFiliales.Count;

                            var LesAlias = new List<string>();

                            File.Delete(strConnexion);

                            //Mettre le contenu de ListeFiliales dans une liste de string

                            foreach (var item in ListeFiliales)
                            {
                                if (cl == null)//Ajout
                                {
                                    var text = item.mId.ToString() + "#" + item.mAliasName + "#" + item.mVille;

                                    LesAlias.Add(text);
                                }
                                else
                                {//Modif
                                    if (cl.mIdModif != item.mId)
                                    {
                                        var text = item.mId.ToString() + "#" + item.mAliasName + "#" + item.mVille;

                                        LesAlias.Add(text);
                                    }

                                }

                            }


                            if (cl != null)//Modif
                            {
                                var Alias = FormatageChaineConnexionDbConn(Auth, CmbBaseDonnees.Text, CmbServeur.Text, txtUser.Text, txtMotPasse.Text, cl.mIdModif, txtVille.Text);
                                LesAlias.Add(Alias);
                            }
                            else//Ajout
                            {
                                var Alias = FormatageChaineConnexionDbConn(Auth, CmbBaseDonnees.Text, CmbServeur.Text, txtUser.Text, txtMotPasse.Text, nbre + 1, txtVille.Text);
                                LesAlias.Add(Alias);
                            }



                            ret = FormatMain.EcrireDansFichier(LesAlias, strConnexion);

                            if (ret)
                            {
                                MessageBox.Show("Connexion configurée avec succès !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                Close();
                            }


                        }

                    }
                }
                else
                {
                    //Erreur de connexion
                    MessageBox.Show("Erreur de connexion.Veuillez vérifier éventuellement que le serveur et/ou la base existent!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur est survenue ! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "FenAlias ->sBtnValider_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }

        private void CmbBaseDonnees_Click(object sender, EventArgs e)
        {
            FillComboBD();
        }

        private void CmbAuthentification_EditValueChanged(object sender, EventArgs e)
        {
            if (Convert.ToInt16(CmbAuthentification.EditValue.ToString()) == 0)//windows
            {
                txtUser.Enabled = false;
                txtMotPasse.Enabled = false;
            }
            else
            {
                txtUser.Enabled = true;
                txtMotPasse.Enabled = true;
                txtUser.Text = string.Empty;
                txtMotPasse.Text = string.Empty;
            }

        }

        private void CmbServeur_Click(object sender, EventArgs e)
        {
            FillComboInstances();
        }
    }
}
