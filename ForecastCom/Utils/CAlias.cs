using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ForecastCom.Utils
{
   public  class CAlias
    {
        public int mId { get; set; }
        public string mVille { get; set; }
        public string mAliasName { get; set; }
        public string mBDName { get; set; }
        public bool IsAbidjan { get; set; }
        public string mPays { get; set; }


        public CAlias()
        {
            mId = 0;
            mVille = string.Empty;
            mAliasName = string.Empty;
            mBDName = string.Empty;
            mPays = string.Empty;
        }


        /// <summary>
        /// Récupérer les chaines de connexions dans le fichier de configuration
        /// </summary>
        /// <returns></returns>
        public List<CAlias> GetAlias()
        {
            var LAlias = new List<CAlias>();
            try
            {
                var strConnexion = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ressources/" + "DbConnexion.ini");

                if (File.Exists(strConnexion))
                {

                    if (strConnexion != null)
                    {
                        using (StreamReader sr = new StreamReader(strConnexion))
                        {
                            string line;
                            // Read and display lines from the file until the end of the file is reached.
                            while ((line = sr.ReadLine()) != null)
                            {
                                if (line != null)
                                {
                                    CAlias CAl = new CAlias();
                                    var chx = line.Split('#');

                                    CAl.mId = int.Parse(chx[0].ToString());
                                    CAl.mVille = chx[chx.Length - 1].Trim();
                                    CAl.mAliasName = chx[1].Trim();
                                    if (chx[1] != null && chx[1] != string.Empty)
                                    {
                                        var bdx = chx[1].Split(';');

                                        var bdn = bdx[0].Trim();

                                        if (bdn != null && bdn != string.Empty)
                                        {
                                            var bd = bdn.Split('=');

                                            CAl.mBDName = bd[1].Trim();

                                            //récupérer le pays de la BD

                                            string NomPays = string.Empty;

                                            if(CAl.mBDName!=string.Empty)
                                            {
                                                NomPays = CAl.mBDName;

                                                string a = NomPays.Replace("AITEK", "");

                                                NomPays = a.Trim();

                                                if (NomPays.Equals("DKR"))
                                                {
                                                    NomPays = "SN";
                                                }

                                                if (NomPays.Equals("MALI"))
                                                {
                                                    NomPays = "ML";
                                                }

                                                if (NomPays.Equals("BURKINA"))
                                                {
                                                    NomPays = "BF";
                                                }

                                                CAl.mPays = NomPays;

                                            }

                                        }

                                        if (CAl.mVille.ToUpper() == "ABIDJAN")
                                        {
                                            CAl.IsAbidjan = true;
                                        }
                                        else
                                        {
                                            CAl.IsAbidjan = false;
                                        }
                                        
                                    }

                                    LAlias.Add(CAl);
                                }

                            }
                        }

                    }

                }
                else
                {
                    //Configurer les connexions
                    MessageBox.Show("Veuillez configurer les connexions aux bases de données dans l'onglet Menu ->Connexions Reporting!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Utils -> CAlias ->GetAlias -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }

            return LAlias;

        }


        public List<CAlias> GetAliasConnexions()
        {
            var LAlias = new List<CAlias>();
            try
            {
                var strConnexion = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ressources/" + "DbConnexion.ini");

                if (File.Exists(strConnexion))
                {

                    if (strConnexion != null)
                    {
                        using (StreamReader sr = new StreamReader(strConnexion))
                        {
                            string line;
                            // Read and display lines from the file until the end of the file is reached.
                            while ((line = sr.ReadLine()) != null)
                            {
                                if (line != null)
                                {
                                    CAlias CAl = new CAlias();
                                    var chx = line.Split('#');

                                    CAl.mId = int.Parse(chx[0].ToString());
                                    CAl.mVille = chx[chx.Length - 1].Trim();
                                    CAl.mAliasName = chx[1].Trim();
                                    if (chx[1] != null && chx[1] != string.Empty)
                                    {
                                        var bdx = chx[1].Split(';');

                                        var bdn = bdx[0].Trim();

                                        if (bdn != null && bdn != string.Empty)
                                        {
                                            var bd = bdn.Split('=');

                                            CAl.mBDName = bd[1].Trim();

                                            //récupérer le pays de la BD

                                            string NomPays = string.Empty;

                                            if (CAl.mBDName != string.Empty)
                                            {
                                                NomPays = CAl.mBDName;

                                                NomPays = NomPays.Replace("AITEK", "");

                                                NomPays = NomPays.Trim();

                                                if (NomPays.Equals("DKR"))
                                                {
                                                    NomPays = "SEN";
                                                }

                                            }

                                        }

                                        if (CAl.mVille.ToUpper() == "ABIDJAN")
                                        {
                                            CAl.IsAbidjan = true;
                                        }
                                        else
                                        {
                                            CAl.IsAbidjan = false;
                                        }


                                    }

                                    LAlias.Add(CAl);
                                }

                            }
                        }

                    }

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Utils -> CAlias -> GetAliasConnexions-> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }

            return LAlias;

        }


        public List<CAlias> GetAliasConnexionsTarget()
        {
            var LAlias = new List<CAlias>();
            try
            {
                var strConnexion = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Ressources/" + "DbConnexionTarget.ini");

                if (File.Exists(strConnexion))
                {

                    if (strConnexion != null)
                    {
                        using (StreamReader sr = new StreamReader(strConnexion))
                        {
                            string line;
                            // Read and display lines from the file until the end of the file is reached.
                            while ((line = sr.ReadLine()) != null)
                            {
                                if (line != null)
                                {
                                    CAlias CAl = new CAlias();
                                    var chx = line.Split('#');

                                    CAl.mId = int.Parse(chx[0].ToString());
                                    CAl.mVille = chx[chx.Length - 1].Trim();
                                    CAl.mAliasName = chx[1].Trim();
                                    if (chx[1] != null && chx[1] != string.Empty)
                                    {
                                        var bdx = chx[1].Split(';');

                                        var bdn = bdx[0].Trim();

                                        if (bdn != null && bdn != string.Empty)
                                        {
                                            var bd = bdn.Split('=');

                                            CAl.mBDName = bd[1].Trim();

                                            //récupérer le pays de la BD

                                            string NomPays = string.Empty;

                                            if (CAl.mBDName != string.Empty)
                                            {
                                                NomPays = CAl.mBDName;

                                                NomPays = NomPays.Replace("AITEK", "");

                                                NomPays = NomPays.Trim();

                                                if (NomPays.Equals("DKR"))
                                                {
                                                    NomPays = "SEN";
                                                }

                                            }

                                        }

                                        if (CAl.mVille.ToUpper() == "ABIDJAN")
                                        {
                                            CAl.IsAbidjan = true;
                                        }
                                        else
                                        {
                                            CAl.IsAbidjan = false;
                                        }


                                    }

                                    LAlias.Add(CAl);
                                }

                            }
                        }

                    }

                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Utils -> CAlias -> GetAliasConnexionsTarget-> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                return null;
            }

            return LAlias;

        }





        //Ecrire les logs
        public static void Log(string msgToWrite)
        {
            var retourLigne = Environment.NewLine;
            const string space = "-------------------------------------------------------------------------------";
            var dateErreur = @"Erreur survenue le " + DateTime.Now.ToString(CultureInfo.InvariantCulture) + "   du poste " + Environment.UserDomainName + "\\" + Environment.UserName;
            //var dateErreur = "";
            var typeErreur = msgToWrite;

            if (!Directory.Exists(@"C:\Report_Log"))
            {
                Directory.CreateDirectory(@"C:\Report_Log");
            }

            // string[] msgfinal = { space , dateErreur , typeErreur , space };
            string msgfinal = space + retourLigne + dateErreur + retourLigne + typeErreur + retourLigne + space;

            var fs = new FileStream(@"C:\Report_Log\ForecastCom_Log.txt", FileMode.Append, FileAccess.Write, FileShare.None);

            var swFromFileStream = new StreamWriter(fs, Encoding.Default);

            swFromFileStream.Write(msgfinal);
            swFromFileStream.Flush();
            swFromFileStream.Close();
            

        }


    }
}
