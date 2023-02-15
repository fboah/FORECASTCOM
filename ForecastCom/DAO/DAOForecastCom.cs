using ForecastCom.Models;
using ForecastCom.Utils;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ForecastCom.DAO
{
    public class DAOForecastCom
    {

        public string GetSigleDoc(int typedocId)
        {
            string ret = string.Empty;
            try
            {
                //                0 = Devis
                //1 = Bon de
                //commande
                //2 = Préparation de
                //livraison
                //3 = Bon de livraison
                //4 = Bon de retour
                //5 = Bon d’avoir
                //6 = Facture
                //7 = Facture
                //comptabilisée
                //8 = Archive

                switch (typedocId)
                {
                    case 0://devis
                        ret = "DE";
                        break;
                    case 1:
                        ret = "BC";
                        break;

                    case 2:
                        ret = "PL";
                        break;

                    case 3:
                        ret = "BL";
                        break;

                    case 4:
                        ret = "BR";
                        break;

                    case 5:
                        ret = "BA";
                        break;

                    case 6:
                        ret = "FA";
                        break;

                    case 7:
                        ret = "FC";
                        break;

                    default:

                        break;
                }

                return ret;
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetSigleDoc -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

                return ret;
            }
        }

        //ramener le statut d'avancement du document
        public string GetSigleStatut(int StatutId, int TypeDoc)
        {
            string ret = string.Empty;
            try
            {
                //0 = Devis
                //1 = Bon de commande
                //2 = Préparation de livraison
                //3 = Bon de livraison
                //4 = Bon de retour
                //5 = Bon d’avoir
                //6 = Facture
                //7 = Facture
                //comptabilisée
                //8 = Archive

                //0 à 2
                //0 = Saisie
                //1 = Confirmé
                //2 = Accepté

                switch (TypeDoc)
                {
                    case 0://Devis
                        if (StatutId == 0)
                        {
                            ret = "60%   : Devis envoyé";
                        }
                        if (StatutId == 1)
                        {
                            ret = "65%   : Devis relancé";
                        }
                        if (StatutId == 2)
                        {
                            ret = "75%   : Négociation en cours";
                        }

                        break;
                    case 1:// Bon de commande
                        ret = "80%   : BC ou BPA reçu";
                        break;

                    case 2://Préparation de livraison
                        ret = "85%   : PL établie";
                        break;

                    case 3://Bon de livraison
                        ret = "90%   : BL établie";
                        break;

                    case 4:// Bon de retour
                        ret = "90%   : BR établie";
                        break;

                    case 5://Bon d’avoir
                        ret = "90%   : BA établie";
                        break;

                    case 6:// Facture
                        ret = "100%   : Facturé";
                        break;

                    case 7://facture comptabilisée
                        ret = "100%   : Facturé";
                        break;

                    case 8://Archive
                        ret = "100%   : Archive";
                        break;

                    default:

                        break;
                }

                return ret;
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetSigleStatut -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

                return ret;
            }
        }

        /// <summary>
        /// REtourne le pays du commerciale,si son pays est base est null ,on prendra son pays recupéré de la chaine de connexion
        /// </summary>
        /// <param name="P"></param>
        /// <returns></returns>
        public string GetPaysCommerciale(string valeurBD, string valeurChaine)
        {
            string ret = string.Empty;
            try
            {
                if (valeurBD != string.Empty)
                {
                    ret = valeurBD;
                }
                else
                {
                    ret = valeurChaine;
                }

                return ret;
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetPaysCommerciale -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return ret;
            }
        }


        public List<FCommercial> GetForecast1(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis)
        {
            try
            {
                var listeFORCAST = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.DO_Type,D.Do_piece,D.Do_date,Fa.FA_CodeFamille,D.CO_No " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        " ,D.CT_Num,E.DO_DateLivr,E.DO_Statut from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                       
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>'' AND D.DO_Type in (0,1,2,3,6,7,8,4,5) " +
                        " group by C.CO_Prenom,C.CO_nom,D.DO_Type,D.Do_piece,D.Do_date,Fa.FA_CodeFamille,D.CO_No,D.CT_Num,E.DO_DateLivr,E.DO_Statut " +

                        " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,do_date,CO_Prenom,CO_Nom,DO_Piece,FA_CodeFamille,CT_Num";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {


                                    var ClasseFCommercial = new FCommercial()
                                    {
                                        mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                        mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                        mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                        mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                        mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                        mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                        mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                        //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                        //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                        mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                        mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                        mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                        mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                        mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                        mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                        mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                        mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))
                                    };

                                    listeFORCAST.Add(ClasseFCommercial);
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecast -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }
        
        public List<FCommercial> GetForecastMARGE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                //Total des Marge

                double TotalMarge = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        NomMultiRevendeur = NomMultiRevendeur.Trim();

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }

                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                 //   FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement des factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }


                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.DO_Type,D.Do_piece,D.Do_date,Fa.FA_Central,Fa.FA_CodeFamille,D.CO_No " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT  ,SUM( case when D.do_type in (1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE* D.DL_PrixRU else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE* D.DL_PrixRU - D.DL_MontantHT  else case when D.DO_TYPE  in (0) THEN D.DL_MontantHT - D.DL_QTE * D.DL_CMUP  else 0 end end end)as Marge " +
                        " ,D.CT_Num,E.DO_DateLivr,E.DO_Statut,D.DO_Ref,fc.CT_Pays,E.DO_Coord01  from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                         " left join " + item.mBDName + ".dbo.F_COMPTET fc ON fc.CT_Num=D.CT_Num  " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                       " group by C.CO_Prenom,C.CO_nom,D.DO_Type,D.Do_piece,D.Do_date,Fa.FA_Central,Fa.FA_CodeFamille,D.CO_No,D.CT_Num,E.DO_DateLivr,E.DO_Statut,D.DO_Ref,fc.CT_Pays,E.DO_Coord01  " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,do_date,CO_Prenom,CO_Nom,DO_Piece,FA_Central,FA_CodeFamille,CT_Num";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mFamilleCentral = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;
                                    //Pays Revendeur
                                    ClasseFCommercial.mPaysRevendeur = reader["CT_Pays"] == DBNull.Value ? string.Empty : reader["CT_Pays"] as string;

                                    //entete

                                    ClasseFCommercial.mDO_Coord01 = reader["DO_Coord01"] == DBNull.Value ? string.Empty : reader["DO_Coord01"] as string;


                                    //if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    //if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    //if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    ////cas Hors Pipe
                                    //if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeTEMP.Add(ClasseFCommercial);

                                }

                                //Parcourir la liste temporaire et faire les différents test 

                                foreach (var obj in listeTEMP)
                                {
                                    //1-Ce n'est pas un devis ,on ajoute
                                    if (obj.mTypeDocId != 0)
                                    {
                                        if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += obj.mMontantHT;

                                        if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += obj.mMontantHT;

                                        if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += obj.mMontantHT;

                                        //cas Hors Pipe
                                        if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += obj.mMontantHT;
                                        listeFORCAST.Add(obj);
                                    }
                                    else
                                    {
                                        //tester si on a un devis et que son numero piece n'est pas egal a
                                        //la reference d'un autre doc dans la liste

                                        var isTestDE = listeTEMP.FirstOrDefault(c => c.mReference.Trim() == obj.mNumPiece && !(c.Equals(obj)));

                                        if (isTestDE == null)//On ajoute le devis car il n'a subit aucune transformation
                                        {
                                            if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += obj.mMontantHT;

                                            if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += obj.mMontantHT;

                                            if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += obj.mMontantHT;

                                            //cas Hors Pipe
                                            if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += obj.mMontantHT;
                                            listeFORCAST.Add(obj);
                                        }

                                    }

                                    //Marge totale

                                    TotalMarge += obj.mMargeBrute;


                                }


                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;

                                    item.mMontantTotalMarge = TotalMarge;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastMARGE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }

        
        public List<FCommercial> GetForecast(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur,bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr,bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                //Total des Marge

                double TotalMarge = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";
                        
                    }
                    
                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");
                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }
                    

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                

                //////////////////////FILTRES famille//////////////////////////////////
            if(IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }
                
             

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        NomMultiRevendeur = NomMultiRevendeur.Trim();

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");
                        
                            FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }

                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8","6,7") + ")  ";

                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                   // FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if(IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement des factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                 //   FiltreFacturePipe = " AND D.DO_Type in (6,7) ";
                    
                }


                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.DO_Type,D.Do_piece,D.Do_date,Fa.FA_Central,Fa.FA_CodeFamille,D.CO_No " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        " ,D.CT_Num,E.DO_DateLivr,E.DO_Statut,D.DO_Ref,fc.CT_Pays,E.DO_Coord01  from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                         " left join " + item.mBDName + ".dbo.F_COMPTET fc ON fc.CT_Num=D.CT_Num  " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev+ FiltreFamCentr+ FiltreFacturePipe+
                       " group by C.CO_Prenom,C.CO_nom,D.DO_Type,D.Do_piece,D.Do_date,Fa.FA_Central,Fa.FA_CodeFamille,D.CO_No,D.CT_Num,E.DO_DateLivr,E.DO_Statut,D.DO_Ref,fc.CT_Pays,E.DO_Coord01  " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,do_date,CO_Prenom,CO_Nom,DO_Piece,FA_Central,FA_CodeFamille,CT_Num";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mFamilleCentral = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;
                                    //Pays Revendeur
                                    ClasseFCommercial.mPaysRevendeur = reader["CT_Pays"] == DBNull.Value ? string.Empty : reader["CT_Pays"] as string;

                                    //entete

                                    ClasseFCommercial.mDO_Coord01 = reader["DO_Coord01"] == DBNull.Value ? string.Empty : reader["DO_Coord01"] as string;


                                    //if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    //if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    //if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    ////cas Hors Pipe
                                    //if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeTEMP.Add(ClasseFCommercial);

                                }

                                //Parcourir la liste temporaire et faire les différents test 

                                foreach (var obj in listeTEMP)
                                {
                                    //1-Ce n'est pas un devis ,on ajoute
                                    if (obj.mTypeDocId != 0)
                                    {
                                        if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += obj.mMontantHT;

                                        if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += obj.mMontantHT;

                                        if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += obj.mMontantHT;

                                        //cas Hors Pipe
                                        if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += obj.mMontantHT;
                                        listeFORCAST.Add(obj);
                                    }
                                    else
                                    {
                                        //tester si on a un devis et que son numero piece n'est pas egal a
                                        //la reference d'un autre doc dans la liste

                                        var isTestDE = listeTEMP.FirstOrDefault(c => c.mReference.Trim() == obj.mNumPiece && !(c.Equals(obj)));

                                        if (isTestDE == null)//On ajoute le devis car il n'a subit aucune transformation
                                        {
                                            if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += obj.mMontantHT;

                                            if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += obj.mMontantHT;

                                            if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += obj.mMontantHT;

                                            //cas Hors Pipe
                                            if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += obj.mMontantHT;
                                            listeFORCAST.Add(obj);
                                        }

                                    }

                                    //Marge totale

                                    TotalMarge += obj.mMargeBrute;


                                }


                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;

                                    item.mMontantTotalMarge = TotalMarge;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecast -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }
        
        public List<FCommercial> GetForecastBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;
                
                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace(",", "','");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace(",", "','");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        NomMultiRevendeur = NomMultiRevendeur.Trim();

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }

                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                   // if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND D.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                   //FiltreDevisAccepte = " AND D.DO_Type=0 AND D.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                 //   FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND D.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement des factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                 //   FiltreFacturePipe = " AND D.DO_Type in (6,7) ";
                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";

                }

                //cas du bundle ,on ne gère plus les famille et famille centralisatrice
                FiltreFam = string.Empty;
                FiltreFamCentr = string.Empty;

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())
              
                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.DO_Type,D.Do_piece,D.Do_date,D.CO_No " +
                        "  ,D.DO_TotalHT as DL_MontantHT " +
                        " ,E.CT_Num,E.DO_DateLivr,D.DO_Statut,D.DO_Ref,fc.CT_Pays,D.DO_Coord01    from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_COMPTET fc ON fc.CT_Num=E.CT_Num  " +
                        //" left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //" left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.DO_Type in (6,7) and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                       " group by C.CO_Prenom,C.CO_nom,D.DO_Type,D.Do_piece,D.Do_date,D.CO_No,D.DO_TotalHT,E.CT_Num,E.DO_DateLivr,D.DO_Statut,D.DO_Ref,fc.CT_Pays,D.DO_Coord01    " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,do_date,CO_Prenom,CO_Nom,DO_Piece,CT_Num";
                  //  reqFINAL = reqFINAL + "  order by PaysBD,do_date,CO_Prenom,CO_Nom,DO_Piece,FA_Central,FA_CodeFamille,CT_Num";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mFamilleCentral = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                 //ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    //Pays Revendeur
                                    ClasseFCommercial.mPaysRevendeur = reader["CT_Pays"] == DBNull.Value ? string.Empty : reader["CT_Pays"] as string;

                                    //entete
                                    ClasseFCommercial.mDO_Coord01 = reader["DO_Coord01"] == DBNull.Value ? string.Empty : reader["DO_Coord01"] as string;


                                    //if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    //if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    //if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    ////cas Hors Pipe
                                    //if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeTEMP.Add(ClasseFCommercial);

                                }

                                //Parcourir la liste temporaire et faire les différents test 

                                foreach (var obj in listeTEMP)
                                {
                                    //1-Ce n'est pas un devis ,on ajoute
                                    if (obj.mTypeDocId != 0)
                                    {
                                        if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += obj.mMontantHT;

                                        if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += obj.mMontantHT;

                                        if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += obj.mMontantHT;

                                        //cas Hors Pipe
                                        if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += obj.mMontantHT;
                                        listeFORCAST.Add(obj);
                                    }
                                    else
                                    {
                                        //tester si on a un devis et que son numero piece n'est pas egal a
                                        //la reference d'un autre doc dans la liste

                                        var isTestDE = listeTEMP.FirstOrDefault(c => c.mReference.Trim() == obj.mNumPiece && !(c.Equals(obj)));

                                        if (isTestDE == null)//On ajoute le devis car il n'a subit aucune transformation
                                        {
                                            if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += obj.mMontantHT;

                                            if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += obj.mMontantHT;

                                            if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += obj.mMontantHT;

                                            //cas Hors Pipe
                                            if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += obj.mMontantHT;
                                            listeFORCAST.Add(obj);
                                        }

                                    }

                                }


                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<FCommercial> GetForecast_ANCIEN(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsSMulFamille)//Multiple famille
                {
                    if (NomMultiFamille != string.Empty)
                    {
                        NomMultiFamille = NomMultiFamille.Replace(",", "','");

                        FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamille)
                    {
                        if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                    }


                }

                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc + ")  ";

                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }



                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.DO_Type,D.Do_piece,D.Do_date,Fa.FA_CodeFamille,D.CO_No " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        " ,D.CT_Num,E.DO_DateLivr,E.DO_Statut from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte +
                       " group by C.CO_Prenom,C.CO_nom,D.DO_Type,D.Do_piece,D.Do_date,Fa.FA_CodeFamille,D.CO_No,D.CT_Num,E.DO_DateLivr,E.DO_Statut " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,do_date,CO_Prenom,CO_Nom,DO_Piece,FA_CodeFamille,CT_Num";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecast -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<FCommercial> GetForecastSTATDateByCom(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }



                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                   // FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.do_date " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        " from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev+ FiltreFamCentr+ FiltreFacturePipe+
                       " group by C.CO_Prenom,C.CO_nom,D.Do_date " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,do_date,CO_Prenom,CO_Nom";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //   ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //   ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    ////{
                                    ////    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    ////    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    ////    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    ////    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    ////    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    ////    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    ////    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    ////    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ////    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    ////    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    ////    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    ////    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    ////    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    ////    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    ////    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    ////    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    ////};
                                    listeFORCAST.Add(ClasseFCommercial);

                                }

                            
                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATDateByCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<FCommercial> GetForecastSTATDateByComBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";
                        
                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                   // if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }



                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND D.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND D.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                   // FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND D.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                FiltreFam = string.Empty;
                FiltreFamCentr = string.Empty;
                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.Do_piece,D.do_date " +
                        "  ,D.DO_TotalHT as DL_MontantHT  " +
                        " from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                         //" left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         //" left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                       " group by C.CO_Prenom,C.CO_nom,D.Do_piece,D.Do_date,D.DO_TotalHT " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,do_date,CO_Prenom,CO_Nom";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                  //ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //   ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //   ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    ////{
                                    ////    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    ////    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    ////    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    ////    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    ////    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    ////    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    ////    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    ////    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ////    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    ////    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    ////    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    ////    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    ////    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    ////    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    ////    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    ////    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    ////};
                                    listeFORCAST.Add(ClasseFCommercial);

                                }
                                
                                //Faire des sommation par commercial et le renvoyer

                                if (listeFORCAST != null)
                                {
                                    if (listeFORCAST.Count > 0)
                                    {
                                        List<FCommercial> ListeDistCom = new List<FCommercial>();
                                        List<FCommercial> ListeDistDate = new List<FCommercial>();

                                        //recupérer les commerciaux  distinctement
                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistCom.Count == 0)
                                            {
                                                ListeDistCom.Add(item);

                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistCom.FirstOrDefault(c => c.mNomCommercial==item.mNomCommercial && c.mPrenomCommercial==item.mPrenomCommercial);

                                                if (isExist == null) ListeDistCom.Add(item);

                                            }
                                        }


                                        //recupérer les  dates distinctement
                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistDate.Count == 0)
                                            {
                                                ListeDistDate.Add(item);

                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistDate.FirstOrDefault(c => c.mDatePiece == item.mDatePiece );

                                                if (isExist == null) ListeDistDate.Add(item);

                                            }
                                        }
                                        

                                        foreach (var item in ListeDistCom)
                                        {
                                           
                                            List<FCommercial> ListeComCal = new List<FCommercial>();

                                            foreach(var obj in ListeDistDate)
                                            {
                                                double MontantCom = 0;
                                                ListeComCal = listeFORCAST.Where(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial && c.mDatePiece== obj.mDatePiece).ToList();

                                                foreach (var elt in ListeComCal)
                                                {
                                                    MontantCom += elt.mMontantHT;
                                                }

                                                FCommercial NewItem = new FCommercial();

                                                NewItem.mMontantHT = MontantCom;
                                                NewItem.mNomCommercial = item.mNomCommercial;
                                                NewItem.mPrenomCommercial = item.mPrenomCommercial;
                                                NewItem.mDatePiece = obj.mDatePiece;
                                                NewItem.mPays = item.mPays;
                                                    
                                                listeTEMP.Add(NewItem);

                                            }

                                        
                                        }


                                    }
                                }




                                ////Mettre à jour les champs pour pipe devis//Pas applicable dans le cas des bundle

                                //foreach (var item in listeFORCAST)
                                //{
                                //    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                //    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                //    if (TotalPipeDevis == 0)
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                //    }
                                //    else
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                //    }

                                //    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                //}

                               // listeTEMP = listeTEMP.OrderBy(x => x.mMontantHT).ToList();

                                return listeTEMP;


                                ////Mettre à jour les champs pour pipe devis

                                //foreach (var item in listeFORCAST)
                                //{
                                //    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                //    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                //    if (TotalPipeDevis == 0)
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                //    }
                                //    else
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                //    }

                                //    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                //}

                                //return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATDateByComBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<FCommercial> GetForecastSTATTypeDocByCom(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres


                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {

                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {

                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }


                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                 //   if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }


                

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                 //   FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                   // FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.DO_Type,D.CO_No   " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge ,E.DO_Statut" +
                        " from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev+ FiltreFamCentr+ FiltreFacturePipe+
                       " group by C.CO_Prenom,C.CO_nom,D.DO_Type,D.CO_No,E.DO_Statut " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,CO_Prenom,CO_Nom,D.DO_Type,E.DO_Statut ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //  ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTypeDocByCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<FCommercial> GetForecastSTATTypeDocByComBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                var listeTEMP = new List<FCommercial>();

                
                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres


                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux.Replace("'", "''") + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }


                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }




                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                 //   FiltreDevisSaisie = " AND D.DO_Type=0 AND D.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                //    FiltreDevisAccepte = " AND D.DO_Type=0 AND D.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND D.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                FiltreFam = string.Empty;
                FiltreFamCentr = string.Empty;

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())              
                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.Do_piece,D.DO_Type,D.CO_No   " +
                        "  ,D.DO_TotalHT as DL_MontantHT " +
                        " from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                         //" left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         //" left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                       " group by C.CO_Prenom,C.CO_nom,D.Do_piece,D.DO_Type,D.CO_No,D.DO_TotalHT " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,CO_Prenom,CO_Nom,D.DO_Type ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //  ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                  //ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                  //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                  //   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                   //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }

                                //Faire des sommation par commercial et Par type de document le renvoyer

                                if (listeFORCAST != null)
                                {
                                    if (listeFORCAST.Count > 0)
                                    {
                                        List<FCommercial> ListeDistCom = new List<FCommercial>();
                                        List<FCommercial> ListeDistTypeDoc = new List<FCommercial>();

                                        //recupérer les commerciaux  distinctement
                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistCom.Count == 0)
                                            {
                                                ListeDistCom.Add(item);

                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistCom.FirstOrDefault(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial);

                                                if (isExist == null) ListeDistCom.Add(item);

                                            }
                                        }


                                        //recupérer les  Type doc distinctement
                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistTypeDoc.Count == 0)
                                            {
                                                ListeDistTypeDoc.Add(item);

                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistTypeDoc.FirstOrDefault(c => c.mTypeDocId == item.mTypeDocId);

                                                if (isExist == null) ListeDistTypeDoc.Add(item);

                                            }
                                        }


                                        foreach (var item in ListeDistCom)
                                        {

                                            List<FCommercial> ListeComCal = new List<FCommercial>();

                                            foreach (var obj in ListeDistTypeDoc)
                                            {
                                                double MontantCom = 0;
                                                ListeComCal = listeFORCAST.Where(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial && c.mTypeDocId == obj.mTypeDocId).ToList();

                                                foreach (var elt in ListeComCal)
                                                {
                                                    MontantCom += elt.mMontantHT;
                                                }

                                                FCommercial NewItem = new FCommercial();

                                                NewItem.mMontantHT = MontantCom;
                                                NewItem.mNomCommercial = item.mNomCommercial;
                                                NewItem.mPrenomCommercial = item.mPrenomCommercial;
                                                NewItem.mTypeDocId = obj.mTypeDocId;
                                                NewItem.mPays = item.mPays;

                                                listeTEMP.Add(NewItem);

                                            }


                                        }


                                    }
                                }



                                ////Mettre à jour les champs pour pipe devis

                                //foreach (var item in listeFORCAST)
                                //{
                                //    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                //    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                //    if (TotalPipeDevis == 0)
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                //    }
                                //    else
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                //    }

                                //    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                //}

                                return listeTEMP;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTypeDocByComBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }



        //nbre de type de document par commerciaux
        public List<FCommercial> GetForecastSTATNbreTypeDocByCom(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        { 

            try
            {
                var listeFORCAST = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreRev = string.Empty;
              

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;

                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }


                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille.Replace("'", "''") + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }





                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                 //   FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,E.DO_Type,D.CO_No   " +
                        "  ,count(Distinct(E.DO_Piece)) as nbreDoc  ,E.DO_Statut" +
                        " from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev+ FiltreFamCentr+ FiltreFacturePipe+
                       " group by C.CO_Prenom,C.CO_nom,E.DO_Type,D.CO_No,E.DO_Statut " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,CO_Prenom,CO_Nom,E.DO_Type,E.DO_Statut ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //  ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    //  ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mNbreDoc = reader["nbreDoc"] == DBNull.Value ? 0 : Convert.ToInt32(reader["nbreDoc"]);

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATNbreTypeDocByCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        //Pour les BUNDLE
        public List<FCommercial> GetForecastSTATNbreTypeDocByComBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {
            
            try
            {
                var listeFORCAST = new List<FCommercial>();

                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreRev = string.Empty;


                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;

                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }


                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }





                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                 //   if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND D.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                   // FiltreDevisAccepte = " AND D.DO_Type=0 AND D.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND D.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                FiltreFamCentr = string.Empty;
                FiltreFam = string.Empty;
                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.Do_piece,E.DO_Type,D.CO_No   " +
                        "  ,count(Distinct(E.DO_Piece)) as nbreDoc ,D.DO_TotalHT as DL_MontantHT ,D.DO_Statut" +
                        " from " + item.mBDName + ".dbo.F_DOCENTETE D " + 
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                         //" left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         //" left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                       " group by C.CO_Prenom,C.CO_nom,D.Do_piece,E.DO_Type,D.CO_No,D.DO_TotalHT,D.DO_Statut " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,CO_Prenom,CO_Nom,E.DO_Type,D.DO_Statut ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //  ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                      ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mNbreDoc = reader["nbreDoc"] == DBNull.Value ? 0 : Convert.ToInt32(reader["nbreDoc"]);

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }

                                //Faire des sommation par commercial et Par type de document le renvoyer

                                if (listeFORCAST != null)
                                {
                                    if (listeFORCAST.Count > 0)
                                    {
                                        List<FCommercial> ListeDistCom = new List<FCommercial>();
                                        List<FCommercial> ListeDistTypeDoc = new List<FCommercial>();

                                        //recupérer les commerciaux  distinctement
                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistCom.Count == 0)
                                            {
                                                ListeDistCom.Add(item);

                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistCom.FirstOrDefault(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial);

                                                if (isExist == null) ListeDistCom.Add(item);

                                            }
                                        }


                                        //recupérer les  Type doc distinctement
                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistTypeDoc.Count == 0)
                                            {
                                                ListeDistTypeDoc.Add(item);

                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistTypeDoc.FirstOrDefault(c => c.mTypeDocId == item.mTypeDocId);

                                                if (isExist == null) ListeDistTypeDoc.Add(item);

                                            }
                                        }


                                        foreach (var item in ListeDistCom)
                                        {

                                            List<FCommercial> ListeComCal = new List<FCommercial>();
                                            List<FCommercial> ListeComCalNbreDoc = new List<FCommercial>();

                                            foreach (var obj in ListeDistTypeDoc)
                                            {
                                                double MontantCom = 0;
                                                int nbreDocCom = 0;

                                                ListeComCal = listeFORCAST.Where(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial && c.mTypeDocId == obj.mTypeDocId).ToList();

                                                ListeComCalNbreDoc = listeFORCAST.Where(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial ).ToList();

                                                foreach (var elt in ListeComCal)
                                                {
                                                    MontantCom += elt.mMontantHT;
                                                }

                                                foreach (var elt in ListeComCalNbreDoc)
                                                {
                                                    nbreDocCom += elt.mNbreDoc;
                                                }

                                                FCommercial NewItem = new FCommercial();

                                                NewItem.mMontantHT = MontantCom;
                                                NewItem.mNomCommercial = item.mNomCommercial;
                                                NewItem.mPrenomCommercial = item.mPrenomCommercial;
                                                NewItem.mTypeDocId = obj.mTypeDocId;
                                                NewItem.mPays = item.mPays;
                                                NewItem.mNbreDoc = nbreDocCom;

                                                listeTEMP.Add(NewItem);

                                            }


                                        }


                                    }
                                }




                                ////Mettre à jour les champs pour pipe devis

                                //foreach (var item in listeFORCAST)
                                //{
                                //    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                //    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                //    if (TotalPipeDevis == 0)
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                //    }
                                //    else
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                //    }

                                //    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                //}

                                return listeTEMP;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATNbreTypeDocByComBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        //Nbre de doc par com par date
        public List<FCommercial> GetForecastSTATDateByNbreDocCom(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        { 

            try
            {
                var listeFORCAST = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }



                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////


                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";
                        
                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA.Replace("'", "''") + "'";
                        }


                    }
                }




                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                   // FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom,D.do_date " +
                        "  ,count(Distinct(E.DO_Piece)) as nbreDoc " +
                        " from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev+FiltreFamCentr+ FiltreFacturePipe+
                       " group by C.CO_Prenom,C.CO_nom,D.Do_date " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,do_date,CO_Prenom,CO_Nom";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    // ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //   ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //   ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mNbreDoc = reader["nbreDoc"] == DBNull.Value ? 0 : Convert.ToInt32(reader["nbreDoc"]);

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATDateByNbreDocCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }




        //Top 10 familles en Montant

        public List<FCommercial> GetForecastSTATTopFamMtant(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        { 

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //Liste des 10 premiers à ramener
                List<FCommercial> LFinal = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;

                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux.Replace("'", "''") + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";
                        
                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }


                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8","6,7") + ")  ";
                   // if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8","6,7") + ")  ";
                    
                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                   // FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select  Fa.FA_CodeFamille " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        "   from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''   " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte +FiltreRev+FiltreFamCentr+ FiltreFacturePipe+
                       " group by Fa.FA_CodeFamille  " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT desc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                  //  ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                   // ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                               //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeFORCAST.Add(ClasseFCommercial);

                                }

                              

                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                //Ramener les 10 premiers élts

                                int i = 0;

                                if (listeFORCAST.Count > 0)
                                {
                                    foreach (var obj in listeFORCAST)
                                    {
                                        if(obj.mMontantHT>0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        

                                    }
                                }

                                return LFinal;

                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTopFamMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        //Top 10 familles en Qtité
        public List<FCommercial> GetForecastSTATTopFamQtite(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        { 

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //pour ramener les 10 premiers elements

                var LFinal = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;

                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");
                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }
                        
                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                   // FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                   // FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select Fa.FA_CodeFamille,sum(D.DL_Qte) as Qtite " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        "   from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte +FiltreRev+FiltreFamCentr+ FiltreFacturePipe+
                       " group by Fa.FA_CodeFamille  " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by Qtite desc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //  ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    // ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    ClasseFCommercial.mQtiteFamille = reader["Qtite"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Qtite"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;

                                    
                                    listeFORCAST.Add(ClasseFCommercial);
                                    

                                }
                                
                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                //Ramener les 10 premiers élts

                                int i = 0;

                                if (listeFORCAST.Count > 0)
                                {
                                    foreach (var obj in listeFORCAST)
                                    {
                                        if(obj.mQtiteFamille>0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        
                                    }
                                }


                                return LFinal;
                                
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTopFamQtite -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        //Top 10 familles Centralisatrice en Montant

        public List<FCommercial> GetForecastSTATTopFamilleCentrMtant(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //Liste des 10 premiers à ramener
                List<FCommercial> LFinal = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;

                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");


                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }


                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr+ "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                   // FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select  Fa.FA_Central " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        "   from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''   " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe+
                       " group by Fa.FA_Central  " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT desc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //  ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    // ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mFamilleCentral = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string;
                                   // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeFORCAST.Add(ClasseFCommercial);

                                }



                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                //Ramener les 10 premiers élts

                                int i = 0;

                                if (listeFORCAST.Count > 0)
                                {
                                    foreach (var obj in listeFORCAST)
                                    {
                                        if(obj.mMontantHT>0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }

                                       
                                    }
                                }


                                return LFinal;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTopFamilleCentrMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        //Top 10 familles Centralisatrice en Qtité
        public List<FCommercial> GetForecastSTATTopFamilleCentrQtite(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //pour ramener les 10 premiers elements

                var LFinal = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                 //   if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                 //   FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                   //FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                   // FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                 //   FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }


                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select Fa.FA_CENTRAL,sum(D.DL_Qte) as Qtite " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        "   from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe+
                       " group by Fa.FA_CENTRAL  " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by Qtite desc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //  ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    // ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                   // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mFamilleCentral = reader["FA_CENTRAL"] == DBNull.Value ? string.Empty : reader["FA_CENTRAL"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    ClasseFCommercial.mQtiteFamille = reader["Qtite"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Qtite"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeFORCAST.Add(ClasseFCommercial);


                                }

                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                //Ramener les 10 premiers élts

                                int i = 0;

                                if (listeFORCAST.Count > 0)
                                {
                                    foreach (var obj in listeFORCAST)
                                    {
                                        if(obj.mQtiteFamille>0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                       


                                    }
                                }


                                return LFinal;

                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTopFamilleCentrQtite -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }

        

        public List<FCommercial> GetForecastRankingCom(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        { 

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres


                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }



                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");


                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }


                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                 //   FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                 //   FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }


                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                   // FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        " from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte +FiltreRev+FiltreFamCentr+ FiltreFacturePipe+
                       " group by C.CO_Prenom,C.CO_nom " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT asc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                  //  ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //   ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //   ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    ////{
                                    ////    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    ////    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    ////    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    ////    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    ////    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    ////    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    ////    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    ////    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ////    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    ////    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    ////    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    ////    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    ////    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    ////    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    ////    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    ////    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    ////};
                                    listeFORCAST.Add(ClasseFCommercial);

                                }


                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastRankingCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        
        public List<FCommercial> GetForecastRankingComBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<FCommercial>();

                //Liste Temporaire
                var listeTEMP = new List<FCommercial>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres


                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }



                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }
                
                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                    
                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND D.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND D.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND D.DO_Statut in (0,2) ";
                }


                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                //Cas de BUNDLE ,on ne considere plus les filtres familles et marque 

                FiltreFam = string.Empty;
                FiltreFamCentr = string.Empty;

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,C.CO_Prenom ,C.CO_Nom " +
                        " ,D.Do_piece ,D.DO_TotalHT as DL_MontantHT  " +
                        " from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                         //" left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         //" left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                       " group by C.CO_Prenom,C.CO_nom,D.Do_piece,D.DO_TotalHT " +

                       " )";

                        
                      LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";
                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT asc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new FCommercial();

                                    ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //  ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                 //   ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //   ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //   ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    
                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    ////{
                                    ////    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    ////    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    ////    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    ////    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    ////    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    ////    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    ////    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    ////    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ////    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    ////    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    ////    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    ////    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    ////    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    ////    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    ////    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    ////    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    ////};
                                    listeFORCAST.Add(ClasseFCommercial);

                                }

                                //Faire des sommation par commercial et le renvoyer

                                if(listeFORCAST!=null)
                                {
                                    if(listeFORCAST.Count>0)
                                    {
                                        List<FCommercial> ListeDistCom = new List<FCommercial>();

                                        foreach(var item in listeFORCAST)
                                        {
                                            if (ListeDistCom.Count == 0)
                                            {
                                                ListeDistCom.Add(item);
                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistCom.FirstOrDefault(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial);

                                                if (isExist == null) ListeDistCom.Add(item);

                                            }
                                        }


                                       

                                        foreach (var item in ListeDistCom)
                                        {
                                            double MontantCom = 0;

                                            List<FCommercial> ListeComCal = new List<FCommercial>();

                                            ListeComCal = listeFORCAST.Where(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial).ToList();

                                            foreach(var elt in ListeComCal)
                                            {
                                                MontantCom += elt.mMontantHT;
                                            }

                                            item.mMontantHT = MontantCom;

                                            listeTEMP.Add(item);

                                        }
                                        
                                
                                    }
                                }




                                ////Mettre à jour les champs pour pipe devis//Pas applicable dans le cas des bundle

                                //foreach (var item in listeFORCAST)
                                //{
                                //    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                //    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                //    if (TotalPipeDevis == 0)
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                //    }
                                //    else
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                //    }

                                //    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                //}

                                listeTEMP = listeTEMP.OrderBy(x => x.mMontantHT).ToList();

                                return listeTEMP;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastRankingComBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }



        //Stat Revendeur 

        //Stat by rev by type de document========================================================================

        public List<ComRevendeur> GetForecastSTATTypeDocByRevendeur(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        { 
            
            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;
                
                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");


                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr.Replace("'", "''") + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                 //   FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        " from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev+FiltreFamCentr+ FiltreFacturePipe+
                       " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                   // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                 //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                   // ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                   // ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //  ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                  //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTypeDocByRevendeur -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<ComRevendeur> GetForecastSTATTypeDocByRevendeurBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();
                var listeTEMP = new List<ComRevendeur>();
                
                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                   // if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND D.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                  //  FiltreDevisAccepte = " AND D.DO_Type=0 AND D.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                 //   FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND D.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                //    FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }
                FiltreFamCentr = string.Empty;
                FiltreFam = string.Empty;

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select '" + item.mPays + "' as PaysBD,E.CT_Num,T.CT_Intitule,D.DO_Type,D.Do_Piece   " +
                        "  , D.DO_TotalHT as DL_MontantHT " +
                        " from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                         //" left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=E.CT_Num " +
                         //" left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                       " group by E.CT_Num,T.CT_Intitule,D.DO_Type, D.DO_Piece, D.DO_TotalHT  " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by PaysBD,E.CT_Num,T.CT_Intitule,D.DO_Type, D.DO_TotalHT ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    // ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    // ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //  ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                               //     ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;
                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                //Faire des sommation par REVENDEUR et le renvoyer

                                if (listeFORCAST != null)
                                {
                                    if (listeFORCAST.Count > 0)
                                    {
                                        List<ComRevendeur> ListeDistRev = new List<ComRevendeur>();

                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistRev.Count == 0)
                                            {
                                                ListeDistRev.Add(item);
                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistRev.FirstOrDefault(c => c.mCT_Num == item.mCT_Num && c.mCT_Intitule == item.mCT_Intitule);

                                                if (isExist == null) ListeDistRev.Add(item);

                                            }
                                        }



                                        foreach (var item in ListeDistRev)
                                        {
                                            double MontantCom = 0;

                                            List<ComRevendeur> ListeComCal = new List<ComRevendeur>();

                                            ListeComCal = listeFORCAST.Where(c => c.mCT_Num == item.mCT_Num && c.mCT_Intitule == item.mCT_Intitule).ToList();

                                            foreach (var elt in ListeComCal)
                                            {
                                                MontantCom += elt.mMontantHT;
                                            }

                                            item.mMontantHT = MontantCom;

                                            listeTEMP.Add(item);

                                        }

                                 

                                    }
                                }

                                ////Mettre à jour les champs pour pipe devis

                                //foreach (var item in listeFORCAST)
                                //{
                                //    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                //    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                //    if (TotalPipeDevis == 0)
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                //    }
                                //    else
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                //    }

                                //    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                //}

                                return listeTEMP;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTypeDocByRevendeurBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        //Top 10 des commerciaux travaillants avec les revendeurs(en %)

        public List<ComRevendeur> GetForecastSTATTopComRevendeurMtant(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        { 

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                //Liste des 10 premiers à ramener
                List<ComRevendeur> LFinal = new List<ComRevendeur>();

                //Liste Temporaire
                var listeTEMP = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux+ "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                 //   FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                   // FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select  C.CO_Prenom ,C.CO_Nom  " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        "   from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +FiltreFamCentr+ FiltreFacturePipe+
                       " group by C.CO_Prenom ,C.CO_Nom   " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT desc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //  ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    // ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                   // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeFORCAST.Add(ClasseFCommercial);

                                }



                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                //Ramener les 10 premiers élts

                                int i = 0;

                                if (listeFORCAST.Count > 0)
                                {
                                    foreach (var obj in listeFORCAST)
                                    {
                                        if(obj.mMontantHT>0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        
                                    }
                                }


                                return LFinal;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTopComRevendeurMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }

        public List<ComRevendeur> GetForecastSTATTopComRevendeurMtantBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                //Liste des 10 premiers à ramener
                List<ComRevendeur> LFinal = new List<ComRevendeur>();

                //Liste Temporaire
                var listeTEMP = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;
                string FiltreFamCentr = string.Empty;
                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");
                            
                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND D.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                   // FiltreDevisAccepte = " AND D.DO_Type=0 AND D.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND D.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                   // FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                FiltreFam = string.Empty;
                FiltreFamCentr = string.Empty;
                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select  C.CO_Prenom ,C.CO_Nom  " +
                        "  , D.DO_Piece, D.DO_TotalHT as DL_MontantHT " +
                        "   from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                         //" left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         //" left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                       " group by C.CO_Prenom ,C.CO_Nom, D.DO_Piece, D.DO_TotalHT   " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT desc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                     // ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    // ClasseFCommercial.mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                  //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeFORCAST.Add(ClasseFCommercial);

                                }


                                //Faire des sommation par REVENDEUR et le renvoyer

                                if (listeFORCAST != null)
                                {
                                    if (listeFORCAST.Count > 0)
                                    {
                                        List<ComRevendeur> ListeDistCom = new List<ComRevendeur>();

                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistCom.Count == 0)
                                            {
                                                ListeDistCom.Add(item);
                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistCom.FirstOrDefault(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial);

                                                if (isExist == null) ListeDistCom.Add(item);

                                            }
                                        }



                                        foreach (var item in ListeDistCom)
                                        {
                                            double MontantCom = 0;

                                            List<ComRevendeur> ListeComCal = new List<ComRevendeur>();

                                            ListeComCal = listeFORCAST.Where(c => c.mNomCommercial == item.mNomCommercial && c.mPrenomCommercial == item.mPrenomCommercial).ToList();

                                            foreach (var elt in ListeComCal)
                                            {
                                                MontantCom += elt.mMontantHT;
                                            }

                                            item.mMontantHT = MontantCom;

                                            listeTEMP.Add(item);

                                        }

                                        listeTEMP = listeTEMP.OrderByDescending(x => x.mMontantHT).ToList();

                                    }
                                }


                                ////Mettre à jour les champs pour pipe devis

                                //foreach (var item in listeFORCAST)
                                //{
                                //    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                //    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                //    if (TotalPipeDevis == 0)
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                //    }
                                //    else
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                //    }

                                //    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                //}

                                //Ramener les 10 premiers élts

                                int i = 0;

                                if (listeTEMP.Count > 0)
                                {
                                    foreach (var obj in listeTEMP)
                                    {
                                        if (obj.mMontantHT > 0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }

                                    }
                                }


                                return LFinal;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTopComRevendeurMtantBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }



        //Top 10 des Revendeurs (en % Montant)

        public List<ComRevendeur> GetForecastSTATTopBESTRevendeurMtant(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        { 

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                //Liste des 15 premiers à ramener
                List<ComRevendeur> LFinal = new List<ComRevendeur>();

                //Liste Temporaire
                var listeTEMP = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {

                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");
                      

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                      

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }


                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                   // if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                 //   FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                   // FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select D.CT_Num,T.CT_Intitule  " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        "   from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''   " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +FiltreFamCentr+ FiltreFacturePipe+
                       " group by D.CT_Num,T.CT_Intitule   " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT desc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //  ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;

                                    ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                 //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                 //   ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeFORCAST.Add(ClasseFCommercial);

                                }



                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                //Ramener les 15 premiers élts
                                //S'assurer qu'ils sont positifs

                                int i = 0;

                                if (listeFORCAST.Count > 0)
                                {
                                    foreach (var obj in listeFORCAST)
                                    {
                                        if(obj.mMontantHT>0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        
                                    }
                                }


                                return LFinal;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTopBESTRevendeurMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<ComRevendeur> GetForecastSTATTopBESTRevendeurMtantBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                var listeTMP = new List<ComRevendeur>();

                //Liste des 15 premiers à ramener
                List<ComRevendeur> LFinal = new List<ComRevendeur>();

                //Liste Temporaire
                var listeTEMP = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string FiltreFacturePipe = string.Empty;


                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");


                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }


                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND D.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                 //   FiltreDevisAccepte = " AND D.DO_Type=0 AND D.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND D.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                FiltreFam = string.Empty;
                FiltreFamCentr = string.Empty;

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    //foreach (var item in LAliasChoisis)
                    //{
                    //    req = @"( select D.CT_Num,T.CT_Intitule  " +
                    //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT  " +
                    //    "   from " + item.mBDName + ".dbo.F_Docligne D " +
                    //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                    //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                    //  //   " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                    //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                    //    // " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                    //    " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''   " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                    //   " group by D.CT_Num,T.CT_Intitule   " +

                    //   " )";

                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select E.CT_Num,T.CT_Intitule ,  " +
                        "  D.DO_Piece, D.DO_TotalHT as DL_MontantHT  " +
                        " from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                        //  " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=E.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe+
                        " group by E.CT_Num,T.CT_Intitule, D.DO_Piece,D.DO_TotalHT " +

                        " )";
                        
                    LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT desc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //  ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;

                                    ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //   ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeFORCAST.Add(ClasseFCommercial);

                                }

                                //Faire des sommation par REVENDEUR et le renvoyer

                                if (listeFORCAST != null)
                                {
                                    if (listeFORCAST.Count > 0)
                                    {
                                        List<ComRevendeur> ListeDistRev = new List<ComRevendeur>();

                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistRev.Count == 0)
                                            {
                                                ListeDistRev.Add(item);
                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistRev.FirstOrDefault(c => c.mCT_Num == item.mCT_Num && c.mCT_Intitule == item.mCT_Intitule);

                                                if (isExist == null) ListeDistRev.Add(item);

                                            }
                                        }

                                        

                                        foreach (var item in ListeDistRev)
                                        {
                                            double MontantCom = 0;

                                            List<ComRevendeur> ListeComCal = new List<ComRevendeur>();

                                            ListeComCal = listeFORCAST.Where(c => c.mCT_Num == item.mCT_Num && c.mCT_Intitule == item.mCT_Intitule).ToList();

                                            foreach (var elt in ListeComCal)
                                            {
                                                MontantCom += elt.mMontantHT;
                                            }

                                            item.mMontantHT = MontantCom;

                                            listeTEMP.Add(item);

                                        }

                                        listeTEMP = listeTEMP.OrderByDescending(x => x.mMontantHT).ToList();

                                    }
                                }


                                //Mettre à jour les champs pour pipe devis

                                //foreach (var item in listeFORCAST)
                                //{
                                //    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                //    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                //    if (TotalPipeDevis == 0)
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                //    }
                                //    else
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                //    }

                                //    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                //}

                                //Ramener les 15 premiers élts
                                //S'assurer qu'ils sont positifs

                                int i = 0;

                                if (listeTEMP.Count > 0)
                                {
                                    foreach (var obj in listeTEMP)
                                    {
                                        if (obj.mMontantHT > 0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }

                                    }
                                }


                                return LFinal;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATTopBESTRevendeurMtantBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }



        //Top 10 des plus mauvais Revendeurs (en % Montant)
        public List<ComRevendeur> GetForecastSTATBADRevendeurMtant(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        { 

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                //Liste des 15 premiers à ramener
                List<ComRevendeur> LFinal = new List<ComRevendeur>();

                //Liste Temporaire
                var listeTEMP = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreFacturePipe = string.Empty;

                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux+ "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }


                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                   // FiltreDevisSaisie = " AND D.DO_Type=0 AND E.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                   // FiltreDevisAccepte = " AND D.DO_Type=0 AND E.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                  //  FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND E.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }

                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select D.CT_Num,T.CT_Intitule  " +
                        "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        "   from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''   " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +FiltreFamCentr+ FiltreFacturePipe+
                       " group by D.CT_Num,T.CT_Intitule   " +

                       " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT asc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //  ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;

                                    ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //   ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeFORCAST.Add(ClasseFCommercial);

                                }



                                //Mettre à jour les champs pour pipe devis

                                foreach (var item in listeFORCAST)
                                {
                                    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                    if (TotalPipeDevis == 0)
                                    {
                                        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                    }
                                    else
                                    {
                                        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                    }

                                    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                }

                                //Ramener les 15 premiers élts

                                int i = 0;

                                if (listeFORCAST.Count > 0)
                                {
                                    foreach (var obj in listeFORCAST)
                                    {
                                        if(obj.mMontantHT>0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }

                                        }


                                    }
                                }


                                return LFinal;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATBADRevendeurMtant -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<ComRevendeur> GetForecastSTATBADRevendeurMtantBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsFacturePipe)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();
              
                
                //Liste des 15 premiers à ramener
                List<ComRevendeur> LFinal = new List<ComRevendeur>();

                //Liste Temporaire
                var listeTEMP = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;
                string FiltreFamCentr = string.Empty;

                string FiltreFacturePipe = string.Empty;

                //TotalMontantSaisieDevis (60%)
                double TotalMontantSaisieDevis = 0;

                //TotalMontantSaisieDevis (75%)
                double TotalMontantAccepteDevis = 0;

                //TotalPipeDevis (Accepte+ Saisie)

                double TotalPipeDevis = 0;

                //Total des montants Hors cas du pipe

                double TotalMontantHT = 0;

                #region Filtres

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////
                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }


                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }

                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }

                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }



                //////////////////////FILTRES Type document//////////////////////////////////
                if (!IschkTousTypeDoc)//Multiple famille
                {

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                  //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                }

                //////////////////////Filtres Devis Saisie ou Accepte (PIPE)//////////////////////

                if (IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = " AND E.DO_Type=0 AND E.DO_Statut=0 ";
                  //  FiltreDevisSaisie = " AND D.DO_Type=0 AND D.DO_Statut=0 ";
                }

                if (IsAccepteDevis)
                {
                    //On ne tient plus compte du filtre Type de document
                    FiltreTypedoc = string.Empty;

                    FiltreDevisAccepte = " AND E.DO_Type=0 AND E.DO_Statut=2 ";
                 //   FiltreDevisAccepte = " AND D.DO_Type=0 AND D.DO_Statut=2 ";
                }

                if (IsAccepteDevis && IsSaisieDevis)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis
                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = " AND E.DO_Type=0 AND E.DO_Statut in (0,2) ";
                 //   FiltreDevisSaisieAccepte = " AND D.DO_Type=0 AND D.DO_Statut in (0,2) ";
                }

                if (IsFacturePipe)
                {
                    //On ne tient plus compte du filtre Type de document,acceptedevis et saisiedevis,
                    //seulement dfes factures

                    FiltreTypedoc = string.Empty;

                    FiltreDevisSaisie = string.Empty;

                    FiltreDevisAccepte = string.Empty;

                    FiltreDevisSaisieAccepte = string.Empty;

                    FiltreFacturePipe = " AND E.DO_Type in (6,7) ";
                  //  FiltreFacturePipe = " AND D.DO_Type in (6,7) ";

                }


                FiltreFamCentr = string.Empty;
                FiltreFam = string.Empty;
                #endregion

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    //foreach (var item in LAliasChoisis)
                    //{
                    //    req = @"( select D.CT_Num,T.CT_Intitule  " +
                    //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                    //    "   from " + item.mBDName + ".dbo.F_Docligne D " +
                    //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                    //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                    //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                    //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                    //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                    //    " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and D.AR_Ref<>''   " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                    //   " group by D.CT_Num,T.CT_Intitule   " +

                    //   " )";

                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select E.CT_Num,T.CT_Intitule ,  " +
                        "  D.DO_Piece, D.DO_TotalHT as DL_MontantHT  " +
                        " from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                        //  " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=E.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>=@dateDebutGen and E.do_date<=@dateFinGen and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr + FiltreFacturePipe +
                        " group by E.CT_Num,T.CT_Intitule, D.DO_Piece,D.DO_TotalHT " +

                        " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by DL_MontantHT asc";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            command.Parameters.Add(new SqlParameter("dateDebutGen", dateDebutGen));
                            command.Parameters.Add(new SqlParameter("dateFinGen", dateFinGen));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //  ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;

                                    ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    // ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //   ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    ////  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    ////mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    //ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    //ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                  //ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    //ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));

                                    //Reference
                                    //     ClasseFCommercial.mReference = reader["DO_Ref"] == DBNull.Value ? string.Empty : reader["DO_Ref"] as string;

                                    if (IsSaisieDevis && !IsAccepteDevis) TotalMontantSaisieDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && !IsSaisieDevis) TotalMontantAccepteDevis += ClasseFCommercial.mMontantHT;

                                    if (IsAccepteDevis && IsSaisieDevis) TotalPipeDevis += ClasseFCommercial.mMontantHT;

                                    //cas Hors Pipe
                                    if (!IsAccepteDevis && !IsSaisieDevis) TotalMontantHT += ClasseFCommercial.mMontantHT;


                                    listeFORCAST.Add(ClasseFCommercial);

                                }

                                //Faire des sommation par REVENDEUR et le renvoyer

                                if (listeFORCAST != null)
                                {
                                    if (listeFORCAST.Count > 0)
                                    {
                                        List<ComRevendeur> ListeDistRev = new List<ComRevendeur>();

                                        foreach (var item in listeFORCAST)
                                        {
                                            if (ListeDistRev.Count == 0)
                                            {
                                                ListeDistRev.Add(item);
                                            }
                                            else//On a des elements,vérifier s'il existe déja
                                            {
                                                var isExist = ListeDistRev.FirstOrDefault(c => c.mCT_Num == item.mCT_Num && c.mCT_Intitule == item.mCT_Intitule);

                                                if (isExist == null) ListeDistRev.Add(item);

                                            }
                                        }



                                        foreach (var item in ListeDistRev)
                                        {
                                            double MontantCom = 0;

                                            List<ComRevendeur> ListeComCal = new List<ComRevendeur>();

                                            ListeComCal = listeFORCAST.Where(c => c.mCT_Num == item.mCT_Num && c.mCT_Intitule == item.mCT_Intitule).ToList();

                                            foreach (var elt in ListeComCal)
                                            {
                                                MontantCom += elt.mMontantHT;
                                            }

                                            item.mMontantHT = MontantCom;

                                            listeTEMP.Add(item);

                                        }

                                        listeTEMP = listeTEMP.OrderBy(x => x.mMontantHT).ToList();

                                    }
                                }

                                ////Mettre à jour les champs pour pipe devis

                                //foreach (var item in listeFORCAST)
                                //{
                                //    item.mMontantSaisieDevis = TotalMontantSaisieDevis;
                                //    item.mMontantAccepteDevis = TotalMontantAccepteDevis;
                                //    if (TotalPipeDevis == 0)
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalMontantSaisieDevis + TotalMontantAccepteDevis;
                                //    }
                                //    else
                                //    {
                                //        item.mMontantTotalPipeDevis = TotalPipeDevis;
                                //    }

                                //    item.mMontantTotalHorsCasPipe = TotalMontantHT;
                                //}

                                //Ramener les 15 premiers élts

                                int i = 0;

                                if (listeTEMP.Count > 0)
                                {
                                    foreach (var obj in listeTEMP)
                                    {
                                        if (obj.mMontantHT > 0)
                                        {
                                            LFinal.Add(obj);

                                            if (i < 9)
                                            {
                                                i += 1;
                                            }
                                            else
                                            {
                                                break;
                                            }

                                        }


                                    }
                                }


                                return LFinal;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetForecastSTATBADRevendeurMtantBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }



        #region AnalyseComparative


        public List<ComRevendeur> GetComparativeCommercial(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis,List<string>LAnnees, bool IsTousCommerciaux,  string NomCommercial, string PrenomCommercial, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr,bool IsConfigFamille, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc,  bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur,bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin,string JourMoisFinBIX)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if(!IsTousCommerciaux)
                {
                    if (NomCommercial != string.Empty && PrenomCommercial != string.Empty)
                    {
                        NomCommercial = NomCommercial.Replace("'", "''");
                        NomCommercial = NomCommercial.Replace(",", "','");

                        PrenomCommercial = PrenomCommercial.Replace("'", "''");
                        PrenomCommercial = PrenomCommercial.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomCommercial + "') and c.CO_Prenom in ('" + PrenomCommercial + "') ";

                    }

                    if (NomCommercial == string.Empty && PrenomCommercial != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomCommercial = PrenomCommercial.Replace("'", "''");
                        PrenomCommercial = PrenomCommercial.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomCommercial + "') ";

                    }

                    if (NomCommercial != string.Empty && PrenomCommercial == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomCommercial = NomCommercial.Replace("'", "''");
                        NomCommercial = NomCommercial.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomCommercial + "') ";


                    }

                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }


                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }
                
                //////////////////////FILTRES Type document//////////////////////////////////
               

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                //if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                    
                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {
                        if(IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }
                        

                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 
                        
                        foreach (var item in LAliasChoisis)
                            {
                                req = @"( select D.DO_Date ,  " +
                                "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                                " from " + item.mBDName + ".dbo.F_Docligne D " +
                                " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                                " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                                 " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                                 " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                                 " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                                " where E.do_date>='"+ periodeDeb + "' and E.do_date<='"+ periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr+
                                " group by D.DO_Date " +

                               " )";

                                LTemp.Add(req);
                            }

                        ////Revendeur
                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        //construire la requete

                        foreach (var obj in LTemp)
                            {
                                reqFINAL += obj + " UNION ALL ";

                            }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by D.DO_Date ";

                            if (mConnection == null) return null;
                            mConnection.ConnectionString = chaineconnexion;
                            mConnection.Open();

                            using (var command = mConnection.CreateCommand())
                            {
                                try
                                {
                                    command.CommandText = reqFINAL;

                                    //limite du temps de reponse 5 minute
                                    command.CommandTimeout = 300;

                                    //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                                    //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                                    using (var reader = command.ExecuteReader())
                                    {
                                        while (reader.Read())
                                        {
                                            var ClasseFCommercial = new ComRevendeur();

                                            //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                            //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                            //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                            //// ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                            //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                            // ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                            //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                           // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                            //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                            //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                         //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                           // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                              ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                            ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                     //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));



                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                        }


                                        return listeFORCAST;
                                    }
                                }
                                finally
                                {
                                    mConnection.Close();
                                }
                            }
                    //    }
                    

            }

                
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeCommercial -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<ComRevendeur> GetComparativeCommercialBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis,List<string>LAnnees, bool IsTousCommerciaux,  string NomCommercial, string PrenomCommercial, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr,bool IsConfigFamille, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc,  bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur,bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin,string JourMoisFinBIX)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if(!IsTousCommerciaux)
                {
                    if (NomCommercial != string.Empty && PrenomCommercial != string.Empty)
                    {
                        NomCommercial = NomCommercial.Replace("'", "''");
                        NomCommercial = NomCommercial.Replace(",", "','");

                        PrenomCommercial = PrenomCommercial.Replace("'", "''");
                        PrenomCommercial = PrenomCommercial.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomCommercial + "') and c.CO_Prenom in ('" + PrenomCommercial + "') ";

                    }

                    if (NomCommercial == string.Empty && PrenomCommercial != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomCommercial = PrenomCommercial.Replace("'", "''");
                        PrenomCommercial = PrenomCommercial.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomCommercial + "') ";

                    }

                    if (NomCommercial != string.Empty && PrenomCommercial == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomCommercial = NomCommercial.Replace("'", "''");
                        NomCommercial = NomCommercial.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomCommercial + "') ";


                    }

                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }


                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }
                
                //////////////////////FILTRES Type document//////////////////////////////////
               

                    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                //if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
                    
                FiltreFam = string.Empty;
                FiltreFamCentr = string.Empty;
                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {
                        if(IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }
                        

                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 
                        
                        foreach (var item in LAliasChoisis)
                            {
                                req = @"( select D.DO_Date ,  " +
                                "  D.DO_Piece, D.DO_TotalHT as DL_MontantHT  " +
                                " from " + item.mBDName + ".dbo.F_DOCENTETE D " + 
                                " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                                " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                               //  " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                                 " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=E.CT_Num " +
                            //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                                " where E.do_date>='"+ periodeDeb + "' and E.do_date<='"+ periodeFin + "' and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr+
                                " group by D.DO_Date, D.DO_Piece,D.DO_TotalHT " +

                               " )";

                                LTemp.Add(req);
                            }

                        ////Revendeur
                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        //construire la requete

                        foreach (var obj in LTemp)
                            {
                                reqFINAL += obj + " UNION ALL ";

                            }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by D.DO_Date ";

                            if (mConnection == null) return null;
                            mConnection.ConnectionString = chaineconnexion;
                            mConnection.Open();

                            using (var command = mConnection.CreateCommand())
                            {
                                try
                                {
                                    command.CommandText = reqFINAL;

                                    //limite du temps de reponse 5 minute
                                    command.CommandTimeout = 300;

                                    //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                                    //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                                    using (var reader = command.ExecuteReader())
                                    {
                                        while (reader.Read())
                                        {
                                            var ClasseFCommercial = new ComRevendeur();

                                            //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                            //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                            //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                            //// ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                            //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                            // ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                            //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                           // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                            //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                            //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                         //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                           // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                              ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                            ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                     //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));



                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                        }


                                        return listeFORCAST;
                                    }
                                }
                                finally
                                {
                                    mConnection.Close();
                                }
                            }
                    //    }
                    

            }

                
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeCommercial -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }

        public List<ComRevendeur> GetComparativeMarque(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees,bool IsTousMarque,string NomFamCentr, bool IschkTousCom, bool IsSMulCom, string NomMultiCommerciaux, string PrenomMultiCommerciaux,string NomCommerciauxDE,string NomCommerciauxA, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE,string NomFamilleA,string NomMultiFamille, string ListTypeDoc, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }
                ///////////////////////FAMILLE CENTR///////////////////////////////////////
                
                if(!IsTousMarque)
                {
                    if (NomFamCentr != string.Empty)
                    {
                        NomFamCentr = NomFamCentr.Replace("'", "''");
                        NomFamCentr = NomFamCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomFamCentr + "')  ";
                    }
                }

                
                    
                //////////////////////FILTRES famille//////////////////////////////////
              
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                        NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                        NomFamilleA = NomFamilleA.Replace("'", "''");
                        NomFamilleDE = NomFamilleDE.Replace("'", "''");
                        if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                


                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }

                //////////////////////FILTRES Type document//////////////////////////////////


                if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
              //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {

                        if (IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }



                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 

                        foreach (var item in LAliasChoisis)
                        {
                            req = @"( select D.DO_Date ,  " +
                            "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                            " from " + item.mBDName + ".dbo.F_Docligne D " +
                            " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                            " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                             " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                             " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                             " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                            " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +FiltreFamCentr+
                            " group by D.DO_Date " +

                           " )";

                            LTemp.Add(req);
                        }

                        ////Revendeur
                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        //construire la requete

                        foreach (var obj in LTemp)
                        {
                            reqFINAL += obj + " UNION ALL ";

                        }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by D.DO_Date ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                            //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    //// ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    // ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));



                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                    //    }


                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeMarque -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<ComRevendeur> GetComparativeFamille(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees,bool IschkTousCom, bool IsSMulCom, string NomCommerciauxDE,string NomCommerciauxA, string NomMultiCommerciaux,string PrenomMultiCommerciaux, bool IschkTousFamille, string NomFamille,  string ListTypeDoc, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres
                

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {

                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");
                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }

              

                //////////////////////FILTRES famille//////////////////////////////////
              
                    if(!IschkTousFamille)
                {
                    if (NomFamille != null)
                    {
                        NomFamille = NomFamille.Replace("'", "''");
                        NomFamille = NomFamille.Replace(",", "','");

                        FiltreFam = " and Fa.FA_CodeFamille in ('" + NomFamille + "')  ";
                    }
                }
                      
                    
                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }

                //////////////////////FILTRES Type document//////////////////////////////////


                if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
               // if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {

                        if (IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }



                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 

                        foreach (var item in LAliasChoisis)
                        {
                            req = @"( select D.DO_Date ,  " +
                            "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                            " from " + item.mBDName + ".dbo.F_Docligne D " +
                            " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                            " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                             " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                             " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                             " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                            " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                            " group by D.DO_Date " +

                           " )";

                            LTemp.Add(req);
                        }

                        ////Revendeur
                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        //construire la requete

                        foreach (var obj in LTemp)
                        {
                            reqFINAL += obj + " UNION ALL ";

                        }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by D.DO_Date ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                            //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    //// ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    // ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));



                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                    //    }


                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeFamille -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }

         
    public List<ComRevendeur> GetComparativeRevendeurBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees,bool IsTousRevendeur,string NomRevendeur, bool IschkTousCom, bool IsSMulCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsConfigFamille, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
        {
           
            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres



                //////////////////////FILTRES Revendeur//////////////////////////////////
                
                if(!IsTousRevendeur)
                {
                    if (NomRevendeur != string.Empty)
                    {
                        NomRevendeur = NomRevendeur.Replace("'", "''");

                        NomRevendeur = NomRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomRevendeur + "')  ";
                    }

                }



                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");
                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }


                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }
                



                //////////////////////FILTRES Type document//////////////////////////////////


                if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
             //   if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                FiltreFam = string.Empty;
                FiltreFamCentr = string.Empty;

                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {

                        if (IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }



                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 

                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select D.DO_Date ,  " +
                        //    "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                        //    " group by D.DO_Date " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}


                        foreach (var item in LAliasChoisis)
                        {
                            req = @"( select D.DO_Date ,  " +
                            "  D.DO_Piece, D.DO_TotalHT as DL_MontantHT  " +
                            " from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                            " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                            " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                             //  " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                             " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=E.CT_Num " +
                            //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                            " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                            " group by D.DO_Date, D.DO_Piece,D.DO_TotalHT " +

                           " )";

                            LTemp.Add(req);
                        }


                        //////Revendeur
                        ////foreach (var item in LAliasChoisis)
                        ////{
                        ////    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        ////    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        ////    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        ////    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        ////    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        ////     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        ////     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        ////     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        ////    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        ////   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        ////   " )";

                        ////    LTemp.Add(req);
                        ////}

                        //construire la requete

                        foreach (var obj in LTemp)
                        {
                            reqFINAL += obj + " UNION ALL ";

                        }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by D.DO_Date ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                            //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    //// ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    // ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));



                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                    //    }


                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeRevendeur -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }

    public List<ComRevendeur> GetComparativeRevendeur(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees, bool IsTousRevendeur, string NomRevendeur, bool IschkTousCom, bool IsSMulCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsConfigFamille, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
    {

        try
        {
            var listeFORCAST = new List<ComRevendeur>();

            string req = string.Empty;
            string reqFINAL = string.Empty;

            string FiltreCom = string.Empty;
            string FiltreFam = string.Empty;
            string FiltreRev = string.Empty;

            string FiltreFamCentr = string.Empty;

            string FiltreTypedoc = string.Empty;

            string FiltreDevisSaisie = string.Empty;

            string FiltreDevisAccepte = string.Empty;

            string FiltreDevisSaisieAccepte = string.Empty;

            string periodeDeb = string.Empty;

            string periodeFin = string.Empty;


            #region Filtres



            //////////////////////FILTRES Revendeur//////////////////////////////////

            if (!IsTousRevendeur)
            {
                if (NomRevendeur != string.Empty)
                {
                    NomRevendeur = NomRevendeur.Replace("'", "''");

                    NomRevendeur = NomRevendeur.Replace(",", "','");

                    FiltreRev = " and D.CT_Num in ('" + NomRevendeur + "')  ";
                }

            }



            //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

            if (IsSMulCom)//Multiple commerciaux
            {

                if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                {
                    NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                    PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                    FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                }

                if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                {
                    //filtre que sur les prenoms
                    PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                    FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                }

                if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                {
                    //filtre que sur les NOMS
                    NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                    FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                }


            }
            else
            {//DE_A
                if (!IschkTousCom)
                {
                    if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                    {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                    }


                }


            }


            ///////////////////////FAMILLE CENTR///////////////////////////////////////

            if (IsSMulFamilleCentr)//Multiple famille centr
            {
                if (NomMultiFamilleCentr != string.Empty)
                {
                    NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                    FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                }


            }
            else
            {//DE_A
                if (!IschkTousFamilleCentr)
                {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                }


            }



            //////////////////////FILTRES famille//////////////////////////////////
            if (IsConfigFamille)
            {
                if (IsSMulFamille)//Multiple famille
                {
                    if (NomMultiFamille != string.Empty)
                    {
                        NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                        FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamille)
                    {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                    }


                }
            }




            //////////////////////FILTRES Type document//////////////////////////////////


            if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
           //     if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

            #endregion

            //List<String> LTemp = new List<String>();

            var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

            using (var mConnection = mProvider.CreateConnection())
            {

                foreach (var elt in LAnnees)
                {

                    if (IsFinFevrierDeb)
                    {
                        int year = Int32.Parse(elt);

                        if (DateTime.IsLeapYear(year))
                        {
                            string a = JourMoisDebBIX + "-" + elt.ToString();
                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }

                    }
                    else
                    {
                        string a = dateDebutGen + "-" + elt.ToString();

                        periodeDeb = DateTime.Parse(a).ToShortDateString();
                    }


                    if (IsFinFevrierFin)
                    {
                        int year = Int32.Parse(elt);

                        if (DateTime.IsLeapYear(year))
                        {
                            string a = JourMoisFinBIX + "-" + elt.ToString();
                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }

                    }
                    else
                    {
                        string a = dateFinGen + "-" + elt.ToString();

                        periodeFin = DateTime.Parse(a).ToShortDateString();
                    }



                    List<String> LTemp = new List<String>();

                    #region  choix Revendeur 

                    foreach (var item in LAliasChoisis)
                    {
                        req = @"( select D.DO_Date ,  " +
                        "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                        " from " + item.mBDName + ".dbo.F_Docligne D " +
                        " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                         " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                         " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                         " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                        " group by D.DO_Date " +

                       " )";

                        LTemp.Add(req);
                    }

                    ////Revendeur
                    //foreach (var item in LAliasChoisis)
                    //{
                    //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                    //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                    //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                    //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                    //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                    //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                    //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                    //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                    //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                    //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                    //   " )";

                    //    LTemp.Add(req);
                    //}

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    #endregion

                }

                if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                reqFINAL = reqFINAL + "  order by D.DO_Date ";

                if (mConnection == null) return null;
                mConnection.ConnectionString = chaineconnexion;
                mConnection.Open();

                using (var command = mConnection.CreateCommand())
                {
                    try
                    {
                        command.CommandText = reqFINAL;

                        //limite du temps de reponse 5 minute
                        command.CommandTimeout = 300;

                        //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                        //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var ClasseFCommercial = new ComRevendeur();

                                //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                //// ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                //ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                // ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));



                                //var ClasseFCommercial = new FCommercial()
                                //{
                                //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                //};

                                listeFORCAST.Add(ClasseFCommercial);
                            }


                            return listeFORCAST;
                        }
                    }
                    finally
                    {
                        mConnection.Close();
                    }
                }
                //    }


            }


        }
        catch (Exception ex)
        {
            var msg = "DAOForecastCom -> GetComparativeRevendeur -> TypeErreur: " + ex.Message;
            CAlias.Log(msg);
            MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return null;

        }

    }

        //Comparatif Multiple BUNDLE========================================================

        public List<ComRevendeur> GetComparativeMultiCommercialAxeBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees, bool IschkTousCom, bool IsSMulCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsConfigFamille, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
        {
            
            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres



                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux.Replace("'", "''") + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux.Replace("'", "''") + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux.Replace("'", "''") + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux.Replace("'", "''") + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE.Replace("'", "''") + "' and C.CO_Nom <='" + NomCommerciauxA.Replace("'", "''") + "'";
                        }


                    }


                }


                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr.Replace("'", "''") + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE.Replace("'", "''") + "' and Fa.FA_Central <='" + NomFamilleCentrA.Replace("'", "''") + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille.Replace("'", "''") + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE.Replace("'", "''") + "' and Fa.FA_CodeFamille <='" + NomFamilleA.Replace("'", "''") + "'";
                        }


                    }
                }

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES Type document//////////////////////////////////


                if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
               // if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {

                        if (IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }



                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 
                        
                        foreach (var item in LAliasChoisis)
                        {
                            req = @"( select C.CO_Prenom ,C.CO_Nom,D.DO_Date ,  " +
                              "  D.DO_Piece, D.DO_TotalHT as DL_MontantHT  " +
                              " from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                              " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                              " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                               //  " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                               " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=E.CT_Num " +
                              //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                              " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                              " group by C.CO_Prenom ,C.CO_Nom,D.DO_Date, D.DO_Piece,D.DO_TotalHT " +

                             " )";


                            // req = @"( select C.CO_Prenom ,C.CO_Nom,D.DO_Date ,  " +
                            // "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                            // " from " + item.mBDName + ".dbo.F_Docligne D " +
                            // " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                            // " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                            //  " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                            //  " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                            //  " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                            // " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                            // " group by C.CO_Prenom ,C.CO_Nom,D.DO_Date " +

                            //" )";

                            LTemp.Add(req);
                        }

                        ////Revendeur
                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        //construire la requete

                        foreach (var obj in LTemp)
                        {
                            reqFINAL += obj + " UNION ALL ";

                        }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by C.CO_Prenom ,C.CO_Nom,D.DO_Date ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                            //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    //// ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));



                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                    //    }


                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeMultiCommercialAxeBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }



        public List<ComRevendeur> GetComparativeMultiCommercialAxe(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees, bool IschkTousCom, bool IsSMulCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsConfigFamille, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres

                

                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }


                    }


                }


                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }
                    
                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");

                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES Type document//////////////////////////////////


                if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
               // if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {

                        if (IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }



                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 
                        
                        foreach (var item in LAliasChoisis)
                        {
                            req = @"( select C.CO_Prenom ,C.CO_Nom,D.DO_Date ,  " +
                            "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                            " from " + item.mBDName + ".dbo.F_Docligne D " +
                            " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                            " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                             " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                             " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                             " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                            " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                            " group by C.CO_Prenom ,C.CO_Nom,D.DO_Date " +

                           " )";

                            LTemp.Add(req);
                        }

                        ////Revendeur
                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        //construire la requete

                        foreach (var obj in LTemp)
                        {
                            reqFINAL += obj + " UNION ALL ";

                        }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by C.CO_Prenom ,C.CO_Nom,D.DO_Date ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                            //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    //// ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                     ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));



                                    //var ClasseFCommercial = new FCommercial()
                                    //{
                                    //    mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string,
                                    //    mRevendeur = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string,
                                    //    mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                    //    mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                    //    mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                    //    mPaysCommerciale =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    mPays =  reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string,
                                    //    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //    mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]),
                                    //    mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"])),
                                    //    mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString()),
                                    //    mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]),
                                    //    mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]),
                                    //    mStatutId= reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]),
                                    //    mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString()),
                                    //    mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]))

                                    //};

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                    //    }


                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeMultiCommercialAxe -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<ComRevendeur> GetComparativeMultiMarqueAxe(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees, bool IschkTousCom, bool IsSMulCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsConfigFamille, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres



                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");


                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }
                        
                    }


                }
                
                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr.Trim() + "')  ";
                    }
                    
                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");

                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");

                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES Type document//////////////////////////////////


                if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
              //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {

                        if (IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }



                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 

                        foreach (var item in LAliasChoisis)
                        {
                            req = @"( select Fa.FA_Central,D.DO_Date ,  " +
                            "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                            " from " + item.mBDName + ".dbo.F_Docligne D " +
                            " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                            " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                             " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                             " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                             " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                            " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                            " group by fa.FA_Central,D.DO_Date " +

                           " )";

                            LTemp.Add(req);
                        }

                        ////Revendeur
                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        //construire la requete

                        foreach (var obj in LTemp)
                        {
                            reqFINAL += obj + " UNION ALL ";

                        }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by fa.FA_Central ,D.DO_Date ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                            //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    ClasseFCommercial.mFamilleCentral = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string;
                                    //  ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    

                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                    //    }


                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeMultiMarqueAxe -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }

        
        public List<ComRevendeur> GetComparativeMultiRevendeurAxe(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees, bool IschkTousCom, bool IsSMulCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsConfigFamille, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres



                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";
                        
                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");

                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }

                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA+ "'";
                        }


                    }
                }

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }
                
                //////////////////////FILTRES Type document//////////////////////////////////


                if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
               // if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {

                        if (IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }



                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 

                        foreach (var item in LAliasChoisis)
                        {
                            req = @"( select D.CT_Num,T.CT_Intitule,D.DO_Date ,  " +
                            "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                            " from " + item.mBDName + ".dbo.F_Docligne D " +
                            " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                            " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                             " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                             " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                             " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                            " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                            " group by D.CT_Num,T.CT_Intitule,D.DO_Date " +

                           " )";

                            LTemp.Add(req);
                        }

                        ////Revendeur
                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        //construire la requete

                        foreach (var obj in LTemp)
                        {
                            reqFINAL += obj + " UNION ALL ";

                        }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by D.CT_Num,T.CT_Intitule ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                            //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                  //  ClasseFCommercial.mFamilleCentral = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string;
                                    //  ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                    //    }


                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeMultiRevendeurAxe -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        public List<ComRevendeur> GetComparativeMultiRevendeurAxeBUNDLE(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees, bool IschkTousCom, bool IsSMulCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentr, bool IsConfigFamille, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres



                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";

                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");

                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }

                    }


                }

                ///////////////////////FAMILLE CENTR///////////////////////////////////////

                if (IsSMulFamilleCentr)//Multiple famille centr
                {
                    if (NomMultiFamilleCentr != string.Empty)
                    {
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace("'", "''");
                        NomMultiFamilleCentr = NomMultiFamilleCentr.Replace(",", "','");

                        FiltreFamCentr = " and Fa.FA_Central in ('" + NomMultiFamilleCentr + "')  ";
                    }

                }
                else
                {//DE_A
                    if (!IschkTousFamilleCentr)
                    {
                        NomFamilleCentrDE = NomFamilleCentrDE.Replace("'", "''");
                        NomFamilleCentrA = NomFamilleCentrA.Replace("'", "''");
                        if (NomFamilleCentrDE != string.Empty && NomFamilleCentrA != string.Empty) FiltreFamCentr = " and Fa.FA_Central>='" + NomFamilleCentrDE + "' and Fa.FA_Central <='" + NomFamilleCentrA + "'";
                    }


                }



                //////////////////////FILTRES famille//////////////////////////////////
                if (IsConfigFamille)
                {
                    if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                            NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                            NomFamilleDE = NomFamilleDE.Replace("'", "''");
                            NomFamilleA = NomFamilleA.Replace("'", "''");
                            if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                }

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and E.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and E.CT_Num >='" + NomRevendeurDE + "' and E.CT_Num <='" + NomRevendeurA + "'";
                    }


                }

                //////////////////////FILTRES Type document//////////////////////////////////


                if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
            //    if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";


                FiltreFamCentr = string.Empty;
                FiltreFam = string.Empty;

                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {

                        if (IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }



                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 

                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select D.CT_Num,T.CT_Intitule,D.DO_Date ,  " +
                        //    "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                        //    " group by D.CT_Num,T.CT_Intitule,D.DO_Date " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        
                        foreach (var item in LAliasChoisis)
                        {
                            req = @"( select E.CT_Num,T.CT_Intitule,D.DO_Date ,  " +
                            "  D.DO_Piece, D.DO_TotalHT as DL_MontantHT  " +
                            " from " + item.mBDName + ".dbo.F_DOCENTETE D " +
                            " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                            " left join " + item.mBDName + ".dbo.F_Docligne E on D.DO_Piece=E.DO_Piece " +
                            //  " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                            " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=E.CT_Num " +
                            //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                            " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and E.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                            " group by E.CT_Num,T.CT_Intitule,D.DO_Date, D.DO_Piece,D.DO_TotalHT " +

                            " )";

                            LTemp.Add(req);
                        }


                        //////Revendeur
                        ////foreach (var item in LAliasChoisis)
                        ////{
                        ////    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        ////    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        ////    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        ////    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        ////    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        ////     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        ////     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        ////     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        ////    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        ////   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        ////   " )";

                        ////    LTemp.Add(req);
                        ////}

                        //construire la requete

                        foreach (var obj in LTemp)
                        {
                            reqFINAL += obj + " UNION ALL ";

                        }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by E.CT_Num,T.CT_Intitule ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                            //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                    //  ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                    //  ClasseFCommercial.mFamilleCentral = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string;
                                    //  ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                    //    }


                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeMultiRevendeurAxeBUNDLE -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }

        
        public List<ComRevendeur> GetComparativeMultiFamilleAxe(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, List<string> LAnnees, bool IschkTousCom, bool IsSMulCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, string ListTypeDoc, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsFinFevrierDeb, string JourMoisDebBIX, bool IsFinFevrierFin, string JourMoisFinBIX)
        {

            try
            {
                var listeFORCAST = new List<ComRevendeur>();

                string req = string.Empty;
                string reqFINAL = string.Empty;

                string FiltreCom = string.Empty;
                string FiltreFam = string.Empty;
                string FiltreRev = string.Empty;

                string FiltreFamCentr = string.Empty;

                string FiltreTypedoc = string.Empty;

                string FiltreDevisSaisie = string.Empty;

                string FiltreDevisAccepte = string.Empty;

                string FiltreDevisSaisieAccepte = string.Empty;

                string periodeDeb = string.Empty;

                string periodeFin = string.Empty;


                #region Filtres



                //////////////////////FILTRES COMMERCIAUX//////////////////////////////////

                if (IsSMulCom)//Multiple commerciaux
                {

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux == string.Empty && PrenomMultiCommerciaux != string.Empty)
                    {
                        //filtre que sur les prenoms
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace("'", "''");
                        PrenomMultiCommerciaux = PrenomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and c.CO_Prenom in ('" + PrenomMultiCommerciaux + "') ";

                    }

                    if (NomMultiCommerciaux != string.Empty && PrenomMultiCommerciaux == string.Empty)
                    {
                        //filtre que sur les NOMS
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace("'", "''");
                        NomMultiCommerciaux = NomMultiCommerciaux.Replace(",", "','");

                        FiltreCom = " and C.CO_Nom in ('" + NomMultiCommerciaux + "') ";


                    }


                }
                else
                {//DE_A
                    if (!IschkTousCom)
                    {
                        if (NomCommerciauxDE != string.Empty && NomCommerciauxA != string.Empty)
                        {
                            NomCommerciauxDE = NomCommerciauxDE.Replace("'", "''");
                            NomCommerciauxA = NomCommerciauxA.Replace("'", "''");
                            FiltreCom = " and C.CO_Nom>='" + NomCommerciauxDE + "' and C.CO_Nom <='" + NomCommerciauxA + "'";
                        }

                    }


                }
                

                //////////////////////FILTRES famille//////////////////////////////////

                 if (IsSMulFamille)//Multiple famille
                    {
                        if (NomMultiFamille != string.Empty)
                        {
                            NomMultiFamille = NomMultiFamille.Replace("'", "''");
                        NomMultiFamille = NomMultiFamille.Replace(",", "','");

                            FiltreFam = " and Fa.FA_CodeFamille in ('" + NomMultiFamille + "')  ";
                        }


                    }
                    else
                    {//DE_A
                        if (!IschkTousFamille)
                        {
                        NomFamilleDE = NomFamilleDE.Replace("'", "''");
                        NomFamilleA = NomFamilleA.Replace("'", "''");
                        if (NomFamilleDE != string.Empty && NomFamilleA != string.Empty) FiltreFam = " and Fa.FA_CodeFamille>='" + NomFamilleDE + "' and Fa.FA_CodeFamille <='" + NomFamilleA + "'";
                        }


                    }
                

                //////////////////////FILTRES Revendeur//////////////////////////////////
                if (IsSMulRevendeur)//Multiple Revendeur
                {
                    if (NomMultiRevendeur != string.Empty)
                    {
                        NomMultiRevendeur = NomMultiRevendeur.Replace("'", "''");

                        NomMultiRevendeur = NomMultiRevendeur.Replace(",", "','");

                        FiltreRev = " and D.CT_Num in ('" + NomMultiRevendeur + "')  ";
                    }


                }
                else
                {//DE_A
                    if (!IschkTousRevendeur)
                    {
                        NomRevendeurDE = NomRevendeurDE.Replace("'", "''");
                        NomRevendeurA = NomRevendeurA.Replace("'", "''");

                        FiltreRev = " and D.CT_Num >='" + NomRevendeurDE + "' and D.CT_Num <='" + NomRevendeurA + "'";
                    }


                }


                //////////////////////FILTRES Type document//////////////////////////////////


                if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND E.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";
              //  if (ListTypeDoc != string.Empty) FiltreTypedoc = "  AND D.DO_Type in (" + ListTypeDoc.Replace("8", "6,7") + ")  ";

                #endregion

                //List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");

                using (var mConnection = mProvider.CreateConnection())
                {

                    foreach (var elt in LAnnees)
                    {

                        if (IsFinFevrierDeb)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisDebBIX + "-" + elt.ToString();
                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateDebutGen + "-" + elt.ToString();

                                periodeDeb = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateDebutGen + "-" + elt.ToString();

                            periodeDeb = DateTime.Parse(a).ToShortDateString();
                        }


                        if (IsFinFevrierFin)
                        {
                            int year = Int32.Parse(elt);

                            if (DateTime.IsLeapYear(year))
                            {
                                string a = JourMoisFinBIX + "-" + elt.ToString();
                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }
                            else
                            {
                                string a = dateFinGen + "-" + elt.ToString();

                                periodeFin = DateTime.Parse(a).ToShortDateString();
                            }

                        }
                        else
                        {
                            string a = dateFinGen + "-" + elt.ToString();

                            periodeFin = DateTime.Parse(a).ToShortDateString();
                        }



                        List<String> LTemp = new List<String>();

                        #region  choix Revendeur 

                        foreach (var item in LAliasChoisis)
                        {
                            req = @"( select Fa.FA_CodeFamille,D.DO_Date ,  " +
                            "  SUM(D.DL_MontantHT) as DL_MontantHT  " +
                            " from " + item.mBDName + ".dbo.F_Docligne D " +
                            " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                            " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                             " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                             " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                             " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                            " where E.do_date>='" + periodeDeb + "' and E.do_date<='" + periodeFin + "' and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev + FiltreFamCentr +
                            " group by fa.FA_CodeFamille,D.DO_Date " +

                           " )";

                            LTemp.Add(req);
                        }

                        ////Revendeur
                        //foreach (var item in LAliasChoisis)
                        //{
                        //    req = @"( select '" + item.mPays + "' as PaysBD,D.CT_Num,T.CT_Intitule,D.DO_Type   " +
                        //    "  ,SUM(D.DL_MontantHT) as DL_MontantHT ,SUM( (case when D.do_type in (0,1,2,3,6,7,8) then D.DL_MontantHT - D.DL_QTE*D.DL_CMUP else case when D.DO_TYPE  in (4,5)  THEN D.DL_QTE*D.DL_CMUP - D.DL_MontantHT else 0 end end)) as Marge " +
                        //    " from " + item.mBDName + ".dbo.F_Docligne D " +
                        //    " left join " + item.mBDName + ".dbo.F_COLLABORATEUR C on D.co_no=C.Co_No " +
                        //    " left join " + item.mBDName + ".dbo.F_DOCENTETE E on D.DO_Piece=E.DO_Piece " +
                        //     " left join " + item.mBDName + ".dbo.F_ARTICLE A on D.AR_Ref=A.AR_Ref " +
                        //     " left join " + item.mBDName + ".dbo.F_COMPTET T on T.CT_Num=D.CT_Num " +
                        //     " left join " + item.mBDName + ".dbo.F_FAMILLE fa ON Fa.FA_CODEFAMILLE=A.FA_CODEFAMILLE " +
                        //    " where E.do_date>=@periodeDeb and E.do_date<=@periodeFin and E.do_domaine=0 and D.AR_Ref<>''  " + FiltreCom + FiltreFam + FiltreTypedoc + FiltreDevisSaisie + FiltreDevisAccepte + FiltreDevisSaisieAccepte + FiltreRev +
                        //   " group by D.CT_Num,T.CT_Intitule,D.DO_Type " +

                        //   " )";

                        //    LTemp.Add(req);
                        //}

                        //construire la requete

                        foreach (var obj in LTemp)
                        {
                            reqFINAL += obj + " UNION ALL ";

                        }

                        #endregion

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by fa.FA_CodeFamille ,D.DO_Date ";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            //command.Parameters.Add(new SqlParameter("periodeDeb", periodeDeb));
                            //command.Parameters.Add(new SqlParameter("periodeFin", periodeFin));


                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var ClasseFCommercial = new ComRevendeur();

                                    //ClasseFCommercial.mNumPiece = reader["Do_piece"] == DBNull.Value ? string.Empty : reader["Do_piece"] as string;
                                    //ClasseFCommercial.mCT_Num = reader["CT_Num"] == DBNull.Value ? string.Empty : reader["CT_Num"] as string;
                                    //ClasseFCommercial.mCT_Intitule = reader["CT_Intitule"] == DBNull.Value ? string.Empty : reader["CT_Intitule"] as string;
                                      ClasseFCommercial.mFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string;
                                  //  ClasseFCommercial.mFamilleCentral = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string;
                                    //  ClasseFCommercial.mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string;
                                    //ClasseFCommercial.mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string;
                                    //// ClasseFCommercial.mPaysCommerciale = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    // ClasseFCommercial.mPays = reader["PaysBD"] == DBNull.Value ? string.Empty : reader["PaysBD"] as string;
                                    //  mAr_Ref = reader["AR_Ref"] == DBNull.Value ? string.Empty : reader["AR_Ref"] as string,
                                    //mDl_Design = reader["DL_Design"] == DBNull.Value ? string.Empty : reader["DL_Design"] as string,
                                    //   ClasseFCommercial.mTypeDocId = reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]);
                                    // ClasseFCommercial.mTypeDoc = GetSigleDoc(reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));
                                    ClasseFCommercial.mDatePiece = reader["Do_date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["Do_date"].ToString());
                                    ClasseFCommercial.mMontantHT = reader["DL_MontantHT"] == DBNull.Value ? 0 : Convert.ToDouble(reader["DL_MontantHT"]);
                                    //  ClasseFCommercial.mMargeBrute = reader["Marge"] == DBNull.Value ? 0 : Convert.ToDouble(reader["Marge"]);
                                    //  ClasseFCommercial.mStatutId = reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]);
                                    ////   ClasseFCommercial.mDateLivraison = reader["DO_DateLivr"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_DateLivr"].ToString());
                                    //  ClasseFCommercial.mDatePiece = reader["DO_Date"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["DO_Date"].ToString());
                                    //ClasseFCommercial.mStatut = GetSigleStatut(reader["DO_Statut"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Statut"]), reader["DO_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["DO_Type"]));


                                    listeFORCAST.Add(ClasseFCommercial);
                                }


                                return listeFORCAST;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                    //    }


                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetComparativeMultiFamilleAxe -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }

        }


        #endregion






        //======================================================================================================//
        //Charger tous les commerciaux
        public List<ComClass> LoadComAll(string chaineconnexion, List<CAlias> LAllAlias)
        {
            var listeComAll = new List<ComClass>();
            try
            {
                string req = string.Empty;
                string reqFINAL = string.Empty;

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAllAlias)
                    {
                        req = @"( select '" + item.mPays + "' as Pays,'" + item.mId.ToString() + "' as IdPays,C.CO_Prenom ,C.CO_Nom,C.CO_No " +
                        "  from " + item.mBDName + ".dbo.F_COLLABORATEUR C " +

                        " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by CO_Nom,CO_Prenom";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var CCommercial = new ComClass()
                                    {

                                        mNomCommercial = reader["CO_nom"] == DBNull.Value ? string.Empty : reader["CO_nom"] as string,
                                        mPrenomCommercial = reader["CO_Prenom"] == DBNull.Value ? string.Empty : reader["CO_Prenom"] as string,
                                        mPays = reader["Pays"] == DBNull.Value ? string.Empty : reader["Pays"] as string,
                                        mNumCommercial = reader["CO_No"] == DBNull.Value ? 0 : Convert.ToInt32(reader["CO_No"]),
                                        mIdPays = reader["IdPays"] == DBNull.Value ? 0 : Convert.ToInt32(reader["IdPays"]),


                                    };

                                    listeComAll.Add(CCommercial);
                                }

                                return listeComAll;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> LoadComAll -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }


        //Charger toutes les familles centralisatrices
        public List<ComFamille> LoadFamilleCentralAll(string chaineconnexion, List<CAlias> LAllAlias)
        {
            var listeFamAll = new List<ComFamille>();
            try
            {
                string req = string.Empty;
                string reqFINAL = string.Empty;

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAllAlias)
                    {
                        req = @"( select '" + item.mPays + "' as Pays,'" + item.mId.ToString() + "' as IdPays,F.FA_CodeFamille,F.FA_Type,F.FA_Intitule,F.FA_Central " +
                        "  from " + item.mBDName + ".dbo.F_FAMILLE F Where FA_Type=2 " +

                        " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by FA_CodeFamille,FA_Intitule";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var CFam = new ComFamille()
                                    {
                                        mFa_CodeFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                        mFa_Central = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string,
                                        mFa_Intitule = reader["FA_Intitule"] == DBNull.Value ? string.Empty : reader["FA_Intitule"] as string,
                                        mPays = reader["Pays"] == DBNull.Value ? string.Empty : reader["Pays"] as string,
                                        mFa_Type = reader["FA_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["FA_Type"]),
                                        mIdPays = reader["IdPays"] == DBNull.Value ? 0 : Convert.ToInt32(reader["IdPays"]),


                                    };

                                    //tester s'il n'existe pas déjà(eviter les doublons)

                                  var test = listeFamAll.FirstOrDefault(n => n.mFa_CodeFamille == CFam.mFa_CodeFamille && n.mIdPays == CFam.mIdPays);
                                   // var test = listeFamAll.FirstOrDefault(n => n.mFa_CodeFamille == CFam.mFa_CodeFamille );

                                    if (test == null) listeFamAll.Add(CFam);
                                }

                                return listeFamAll;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> LoadFamilleCentralAll -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }


        //Charger toutes les familles
        public List<ComFamille> LoadFamilleAll(string chaineconnexion, List<CAlias> LAllAlias)
        {
            var listeFamAll = new List<ComFamille>();
            try
            {
                string req = string.Empty;
                string reqFINAL = string.Empty;

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAllAlias)
                    {
                        req = @"( select '" + item.mPays + "' as Pays,'" + item.mId.ToString() + "' as IdPays,F.FA_CodeFamille,F.FA_Type,F.FA_Intitule,F.FA_Central " +
                        "  from " + item.mBDName + ".dbo.F_FAMILLE F WHERE F.FA_Type<>2 " +

                        " )";

                        LTemp.Add(req);
                    }

                    //construire la requete

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by FA_CodeFamille,FA_Intitule";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var CFam = new ComFamille()
                                    {
                                        mFa_CodeFamille = reader["FA_CodeFamille"] == DBNull.Value ? string.Empty : reader["FA_CodeFamille"] as string,
                                        mFa_Central = reader["FA_Central"] == DBNull.Value ? string.Empty : reader["FA_Central"] as string,
                                        mFa_Intitule = reader["FA_Intitule"] == DBNull.Value ? string.Empty : reader["FA_Intitule"] as string,
                                        mPays = reader["Pays"] == DBNull.Value ? string.Empty : reader["Pays"] as string,
                                        mFa_Type = reader["FA_Type"] == DBNull.Value ? 0 : Convert.ToInt32(reader["FA_Type"]),
                                        mIdPays = reader["IdPays"] == DBNull.Value ? 0 : Convert.ToInt32(reader["IdPays"]),
                                        

                                    };

                                    //tester s'il n'existe pas déjà(eviter les doublons)

                                    var test = listeFamAll.FirstOrDefault(n => n.mFa_CodeFamille == CFam.mFa_CodeFamille && n.mIdPays == CFam.mIdPays);
                                  // var test = listeFamAll.FirstOrDefault(n => n.mFa_CodeFamille == CFam.mFa_CodeFamille );

                                    if (test == null) listeFamAll.Add(CFam);
                                }

                                return listeFamAll;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> LoadFamilleAll -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }


        //Charger tous les revendeurs
        public List<ComRevendeur> LoadRevendeurAll(string chaineconnexion, List<CAlias> LAllAlias)
        {
            var listeRevAll = new List<ComRevendeur>();
            try
            {
                string req = string.Empty;
                string reqFINAL = string.Empty;

                List<String> LTemp = new List<String>();

                var mProvider = DbProviderFactories.GetFactory("System.Data.SqlClient");
                using (var mConnection = mProvider.CreateConnection())

                {
                    foreach (var item in LAllAlias)
                    {
                        req = @"( select '" + item.mPays + "' as Pays,'" + item.mId.ToString() + "' as IdPays,C.cbmarq,C.CT_NUM,C.CT_INTITULE,Cg_numPrinc " +
                        "  from " + item.mBDName + ".dbo.F_comptet C " +
                        "  LEFT outer join " + item.mBDName + ".dbo.F_COMPTEG on (cg_num=cg_numPrinc)  where ct_type=0 " +

                        " )";

                        LTemp.Add(req);
                    }

                    //construire la requete  CT_DateCreate

                    foreach (var obj in LTemp)
                    {
                        reqFINAL += obj + " UNION ALL ";

                    }

                    if (reqFINAL != string.Empty) reqFINAL = reqFINAL.Substring(0, reqFINAL.Length - 10);

                    reqFINAL = reqFINAL + "  order by CT_NUM,CT_INTITULE";

                    if (mConnection == null) return null;
                    mConnection.ConnectionString = chaineconnexion;
                    mConnection.Open();

                    using (var command = mConnection.CreateCommand())
                    {
                        try
                        {
                            command.CommandText = reqFINAL;

                            //limite du temps de reponse 5 minute
                            command.CommandTimeout = 300;

                            using (var reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var CRev = new ComRevendeur()
                                    {
                                        mCT_Num = reader["CT_NUM"] == DBNull.Value ? string.Empty : reader["CT_NUM"] as string,
                                        mCT_Intitule = reader["CT_INTITULE"] == DBNull.Value ? string.Empty : reader["CT_INTITULE"] as string,
                                        mPays = reader["Pays"] == DBNull.Value ? string.Empty : reader["Pays"] as string,
                                        mCollectif = reader["Pays"] == DBNull.Value ? string.Empty : reader["Pays"] as string,
                                     //   mDateCreate= reader["CT_DateCreate"] == DBNull.Value ? new DateTime() : DateTime.Parse(reader["CT_DateCreate"].ToString()),
                                        mIdPays = reader["IdPays"] == DBNull.Value ? 0 : Convert.ToInt32(reader["IdPays"]),


                                    };

                                    //tester s'il n'existe pas déjà(eviter les doublons)

                                var test = listeRevAll.FirstOrDefault(n => n.mCT_Num == CRev.mCT_Num && n.mIdPays == CRev.mIdPays);
                                 //   var test = listeRevAll.FirstOrDefault(n => n.mCT_Num == CRev.mCT_Num );

                                    if (test == null) listeRevAll.Add(CRev);
                                }

                                return listeRevAll;
                            }
                        }
                        finally
                        {
                            mConnection.Close();
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> LoadRevendeurAll -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }

        }
        
        //ramener une liste d'apres les Id pays choisis
        public List<ComFamille> GetListByListIdFam(List<String> Lcritere, List<ComFamille> ListElt)
        {
            List<ComFamille> LRetour = new List<ComFamille>();
            try
            {

                if (Lcritere.Count > 0)
                {
                    foreach (var item in Lcritere)
                    {
                        List<ComFamille> LId = new List<ComFamille>();
                        LId = ListElt.Where(c => c.mIdPays == Convert.ToInt32(item.ToString())).ToList();

                        LRetour.AddRange(LId);
                    }

                }

                LRetour = LRetour.OrderBy(x => x.mFa_CodeFamille).ToList();

                return LRetour;
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetListByListIdFam -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }
        }



        //ramener une liste d'apres les Id pays choisis et liste des fam central choisies
        public List<ComFamille> GetListByListIdPaysFamCentrFam(List<String> Lcritere, List<ComFamille> ListElt, List<String> ListeFamCentr)
        {
            List<ComFamille> LRetourTMP = new List<ComFamille>();
            List<ComFamille> LRetour = new List<ComFamille>();
            try
            {
                if (Lcritere.Count > 0)
                {
                    foreach (var item in Lcritere)
                    {
                        List<ComFamille> LId = new List<ComFamille>();
                        LId = ListElt.Where(c => c.mIdPays == Convert.ToInt32(item.ToString())).ToList();

                        LRetourTMP.AddRange(LId);
                    }

                }

                if(ListeFamCentr.Count>0 && LRetourTMP.Count>0)
                {
                    foreach (var item in ListeFamCentr)
                    {
                        List<ComFamille> LFamille = new List<ComFamille>();
                        LFamille = LRetourTMP.Where(c => c.mFa_Central.Equals(item)).ToList();

                        LRetour.AddRange(LFamille);
                    }

                }

              if(LRetour.Count>0)  LRetour = LRetour.OrderBy(x => x.mFa_CodeFamille).ToList();

                return LRetour;
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetListByListIdPaysFamCentrFam -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }
        }

      
        //ramener une liste d'apres les Id commerciaux choisis
        public List<ComClass> GetListByListIdCom(List<String> Lcritere, List<ComClass> ListElt)
        {
            List<ComClass> LRetour = new List<ComClass>();
            try
            {

                if (Lcritere.Count > 0)
                {
                    foreach (var item in Lcritere)
                    {
                        List<ComClass> LId = new List<ComClass>();
                        LId = ListElt.Where(c => c.mIdPays == Convert.ToInt32(item.ToString())).ToList();

                        LRetour.AddRange(LId);
                    }

                }

                LRetour = LRetour.OrderBy(x => x.mNomCommercial).ToList();

                return LRetour;
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetListByListIdCom -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }
        }


        //ramener une liste d'apres les Id Revendeur choisis
        public List<ComRevendeur> GetListByListIdRev(List<String> Lcritere, List<ComRevendeur> ListElt)
        {
            List<ComRevendeur> LRetour = new List<ComRevendeur>();
            try
            {

                if (Lcritere.Count > 0)
                {
                    foreach (var item in Lcritere)
                    {
                        List<ComRevendeur> LId = new List<ComRevendeur>();
                        LId = ListElt.Where(c => c.mIdPays == Convert.ToInt32(item.ToString())).ToList();

                        LRetour.AddRange(LId);
                    }

                }

                LRetour = LRetour.OrderBy(x => x.mCT_Intitule).ToList();

                return LRetour;
            }
            catch (Exception ex)
            {
                var msg = "DAOForecastCom -> GetListByListIdRev -> TypeErreur: " + ex.Message;
                CAlias.Log(msg);
                MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "FORECASTCOM", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;

            }
        }
        

        //Raùmener une liste suivant les bornes choisies
        public List<ComClass> GetListByDEACom(string Debut, string Fin, List<ComClass> ListElt)
        {
            List<ComClass> LId = new List<ComClass>();
            try
            {
                //   LId = ListElt.Where(n => n.mNomCommercial >= Debut && n.mNomCommercial <= Fin).tolist();

                return LId;
            }
            catch (Exception ex)
            {
                return LId;
            }

        }

    }
}
