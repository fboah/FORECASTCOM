
using ForecastCom.ConnTarget;
using ForecastCom.Services;
using ForecastCom.Utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ForecastCom
{
    public partial class ConAdmin : Form
    {
        private readonly CAlias daoMain = new CAlias();
        private List<CAlias> ListeConnexions = new List<CAlias>();

        private readonly FormatageForecastCom FormatMain = new FormatageForecastCom();

        public int id = 0;

        public int idModif = 0;

        public static CModif mod = new CModif();


        public ConAdmin()
        {
            InitializeComponent();
        }

        private void sBtnAjoutConTarg_Click(object sender, EventArgs e)
        {
            try
            {
                CModif cl = new CModif();

                cl.mIsModif = false;
                cl.mIdModif = 0;

                FenConAdmin FAlias = new FenConAdmin();

                FAlias.ShowDialog();

                //Remplir la liste des filiales
                ListeConnexions = daoMain.GetAliasConnexionsTarget();

                gridControlConTarget.DataSource = ListeConnexions;
            }
            catch (Exception ex)
            {
                //  MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "REPORTING HPE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Connexion ->sBtnAjoutCon_Click -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
            }
        }


        private void sBtnSupprConTarg_Click(object sender, EventArgs e)
        {
            var villeTodelete = string.Empty;
            try
            {

                if (gridView1.DataRowCount != 0)
                {

                    var cid = gridView1.GetRowCellValue(gridView1.FocusedRowHandle, "mId").ToString();

                    id = Convert.ToInt32(cid);

                    if (id > 0 && ListeConnexions.Count > 0)
                    {
                        var AliastoDelete = ListeConnexions.Where(c => c.mId == id).ToList();

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
                                var strConnexion = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ressources/" + "DbConnexionTarget.ini");

                                if (File.Exists(strConnexion)) File.Delete(strConnexion);

                                //recupérer une liste sans la ville a supprimer

                                var ret = false;

                                ListeConnexions = ListeConnexions.Where(c => c.mId != id).ToList();

                                if (ListeConnexions.Count > 0)
                                {
                                    var LesAlias = new List<string>();


                                    //Mettre le contenu de ListeFiliales dans une liste de string

                                    foreach (var item in ListeConnexions)
                                    {
                                        var text = item.mId.ToString() + "#" + item.mAliasName + "#" + item.mVille;

                                        LesAlias.Add(text);
                                    }

                                    ret = FormatMain.EcrireDansFichier(LesAlias, strConnexion);

                                    if (ret) MessageBox.Show("L'alias a été supprimé avec succès !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                    //Rafraichir la grid

                                    gridControlConTarget.DataSource = ListeConnexions;
                                }
                                else//On supprime le dernier élément de la grid
                                {

                                    MessageBox.Show("L'alias a été supprimé avec succès !", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    //Rafraichir la grid

                                    gridControlConTarget.DataSource = ListeConnexions;
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

        

        private void ConnTarget_Load(object sender, EventArgs e)
        {
            //Remplir la liste des filiales
            ListeConnexions = daoMain.GetAliasConnexionsTarget();

            gridControlConTarget.DataSource = ListeConnexions;
        }

       
    }
}
