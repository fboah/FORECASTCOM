using ForecastCom.DAO;
using ForecastCom.Models;
using ForecastCom.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.Services
{
   public  class FormatageForecastCom
    {
        private readonly DAOForecastCom dao = new DAOForecastCom();

        public bool EcrireDansFichier(List<string> msgTowrite, string CheminFichier)
        {
            bool test = false;

            try
            {
                if (msgTowrite != null)
                {
                    //fs va créer le fichier .txt
                    var fs = new FileStream(CheminFichier, FileMode.Append, FileAccess.Write, FileShare.None);

                    var swFromFileStreamDefaultEnc = new StreamWriter(fs, Encoding.Default);
                    foreach (var str in msgTowrite)
                        swFromFileStreamDefaultEnc.Write(str + "\r\n");
                    swFromFileStreamDefaultEnc.Flush();
                    swFromFileStreamDefaultEnc.Close();

                    test = true;
                }


            }
            catch (Exception ex)
            {
                // MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "Reporting HPI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Services -> FormatageForecastCom ->EcrireDansFichier -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                test = false;
            }

            return test;
        }


       

        public bool FormatForecast(string Chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis,Microsoft.Office.Interop.Excel.Application procxls, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentral,bool IsFacturePipe)
        {
            bool IsOK = false;

            List<FCommercial> ListFC = new List<FCommercial>();

            try
            {
                if (LAliasChoisis.Count > 0)

                {
                    ListFC = dao.GetForecast(Chaineconnexion,  dateDebutGen,  dateFinGen,  LAliasChoisis,  IsSMulCom,  IschkTousCom,  NomCommerciauxDE,  NomCommerciauxA,  NomMultiCommerciaux,  PrenomMultiCommerciaux,  IsSMulFamille,  IschkTousFamille,  NomFamilleDE,  NomFamilleA,  NomMultiFamille,  IschkTousTypeDoc,  ListTypeDoc,  IsSaisieDevis,  IsAccepteDevis, IsSMulRevendeur, IschkTousRevendeur, NomRevendeurDE, NomRevendeurA, NomMultiRevendeur,  IsConfigFamille,  IsSMulFamilleCentr,  IschkTousFamilleCentr,  NomFamilleCentrDE,  NomFamilleCentrA,  NomMultiFamilleCentral, IsFacturePipe);
                    

                        int i = 2;//Ligne 
                    int j = 1;//Colonne
                    
                    if (ListFC.Count > 0)
                    {
                        //On a des PRS,on écrit dans la feuille excel DATA
                        foreach (var item in ListFC)
                        {
                            //Numero piece(ligne=2 ;colonne=1)
                            EcrireDansExcel(procxls, i, j, item.mNumPiece);

                            //Revendeur
                            EcrireDansExcel(procxls, i, j + 1, item.mRevendeur);

                            //Famille
                            EcrireDansExcel(procxls, i, j + 2, item.mFamille);

                            //Nom et prenom
                            
                            EcrireDansExcel(procxls, i, j + 3, item.mPrenomCommercial + " " + item.mNomCommercial);

                            //Pays
                            EcrireDansExcel(procxls, i, j + 4, item.mPays);
                           
                     
                            //Type Doc
                            EcrireDansExcel(procxls, i, j + 5, item.mTypeDoc);

                            //Statut Doc
                            EcrireDansExcel(procxls, i, j + 6, item.mStatut);


                            //Date Pièce
                            DateTime datpiece = DateTime.Parse(item.mDatePiece.ToShortDateString());
                            EcrireDansExcelDATE(procxls, i, j + 7, datpiece);
                            
                            //Date Livraison
                            //DateTime datLivraison = DateTime.Parse(item.mDateLivraison.ToShortDateString());
                            //EcrireDansExcelDATE(procxls, i, j + 8, datLivraison);


                            //Montant HT
                            EcrireDansExcel(procxls, i, j + 8,( Math.Round(item.mMontantHT)).ToString());

                            //MB
                            EcrireDansExcel(procxls, i, j + 9, (Math.Round(item.mMargeBrute)).ToString());

                         
                            i += 1;
                        }

                        IsOK = true;

                    }

                    if (ListFC.Count == 0)
                    {
                        IsOK = true;
                    }
                    
                }
                return IsOK;

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "Reporting HPI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Services -> FormatageForecastCom ->FormatForecast -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return IsOK;

            }

        }

        public bool FormatForecastBUNDLE(string Chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis, Microsoft.Office.Interop.Excel.Application procxls, bool IsSMulCom, bool IschkTousCom, string NomCommerciauxDE, string NomCommerciauxA, string NomMultiCommerciaux, string PrenomMultiCommerciaux, bool IsSMulFamille, bool IschkTousFamille, string NomFamilleDE, string NomFamilleA, string NomMultiFamille, bool IschkTousTypeDoc, string ListTypeDoc, bool IsSaisieDevis, bool IsAccepteDevis, bool IsSMulRevendeur, bool IschkTousRevendeur, string NomRevendeurDE, string NomRevendeurA, string NomMultiRevendeur, bool IsConfigFamille, bool IsSMulFamilleCentr, bool IschkTousFamilleCentr, string NomFamilleCentrDE, string NomFamilleCentrA, string NomMultiFamilleCentral, bool IsFacturePipe)
        {
            bool IsOK = false;

            List<FCommercial> ListFC = new List<FCommercial>();

            try
            {
                if (LAliasChoisis.Count > 0)

                {
                    ListFC = dao.GetForecastBUNDLE(Chaineconnexion, dateDebutGen, dateFinGen, LAliasChoisis, IsSMulCom, IschkTousCom, NomCommerciauxDE, NomCommerciauxA, NomMultiCommerciaux, PrenomMultiCommerciaux, IsSMulFamille, IschkTousFamille, NomFamilleDE, NomFamilleA, NomMultiFamille, IschkTousTypeDoc, ListTypeDoc, IsSaisieDevis, IsAccepteDevis, IsSMulRevendeur, IschkTousRevendeur, NomRevendeurDE, NomRevendeurA, NomMultiRevendeur, IsConfigFamille, IsSMulFamilleCentr, IschkTousFamilleCentr, NomFamilleCentrDE, NomFamilleCentrA, NomMultiFamilleCentral, IsFacturePipe);


                    int i = 2;//Ligne 
                    int j = 1;//Colonne

                    if (ListFC.Count > 0)
                    {
                        //On a des PRS,on écrit dans la feuille excel DATA
                        foreach (var item in ListFC)
                        {
                            //Numero piece(ligne=2 ;colonne=1)
                            EcrireDansExcel(procxls, i, j, item.mNumPiece);

                            //Revendeur
                            EcrireDansExcel(procxls, i, j + 1, item.mRevendeur);

                            //Famille
                            EcrireDansExcel(procxls, i, j + 2, item.mFamille);

                            //Nom et prenom

                            EcrireDansExcel(procxls, i, j + 3, item.mPrenomCommercial + " " + item.mNomCommercial);

                            //Pays
                            EcrireDansExcel(procxls, i, j + 4, item.mPays);


                            //Type Doc
                            EcrireDansExcel(procxls, i, j + 5, item.mTypeDoc);

                            //Statut Doc
                            EcrireDansExcel(procxls, i, j + 6, item.mStatut);


                            //Date Pièce
                            DateTime datpiece = DateTime.Parse(item.mDatePiece.ToShortDateString());
                            EcrireDansExcelDATE(procxls, i, j + 7, datpiece);

                            //Date Livraison
                            //DateTime datLivraison = DateTime.Parse(item.mDateLivraison.ToShortDateString());
                            //EcrireDansExcelDATE(procxls, i, j + 8, datLivraison);


                            //Montant HT
                            EcrireDansExcel(procxls, i, j + 8, (Math.Round(item.mMontantHT)).ToString());

                            //MB
                            EcrireDansExcel(procxls, i, j + 9, (Math.Round(item.mMargeBrute)).ToString());

                            ////Reference
                            //EcrireDansExcel(procxls, i, j + 10, item.mReference);
                       

                            i += 1;
                        }

                        IsOK = true;

                    }

                    if (ListFC.Count == 0)
                    {
                        IsOK = true;
                    }

                }
                return IsOK;

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "Reporting HPI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Services -> FormatageForecastCom ->FormatForecastBUNDLE -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return IsOK;

            }

        }


        public bool FormatForecastECRIRE(List<FCommercial>ListFC, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            bool IsOK = false;
            
            try
            {
                    int ligne = 2;//Ligne 
                    int col = 1;//Colonne

                    if (ListFC.Count > 0)
                    {
                        //On a des PRS,on écrit dans la feuille excel DATA
                        foreach (var item in ListFC)
                        {
                            //Numero piece(ligne=2 ;colonne=1)
                          //  EcrireDansExcel(procxls, i, j, item.mNumPiece);
                        worksheet.Cells[ligne, col] = item.mNumPiece;

                        //Revendeur
                       // EcrireDansExcel(procxls, i, j + 1, item.mRevendeur);
                        worksheet.Cells[ligne, col+1] = item.mRevendeur;

                        //Famille
                       // EcrireDansExcel(procxls, i, j + 2, item.mFamille);
                        worksheet.Cells[ligne, col + 2] = item.mFamille;

                        //Nom et prenom

                      //  EcrireDansExcel(procxls, i, j + 3, item.mPrenomCommercial + " " + item.mNomCommercial);
                        worksheet.Cells[ligne, col + 3] = item.mPrenomCommercial + " " + item.mNomCommercial;


                        //Pays
                       // EcrireDansExcel(procxls, i, j + 4, item.mPays);
                        worksheet.Cells[ligne, col + 4] = item.mPays;


                        //Type Doc
                     //   EcrireDansExcel(procxls, i, j + 5, item.mTypeDoc);
                        worksheet.Cells[ligne, col + 5] = item.mTypeDoc;

                        //Statut Doc
                       // EcrireDansExcel(procxls, i, j + 6, item.mStatut);
                        worksheet.Cells[ligne, col + 6] = item.mStatut;


                        //Date Pièce
                        DateTime datpiece = DateTime.Parse(item.mDatePiece.ToShortDateString());
                         //   EcrireDansExcelDATE(procxls, i, j + 7, datpiece);
                        worksheet.Cells[ligne, col + 7] = datpiece;

                        //Date Livraison
                        //DateTime datLivraison = DateTime.Parse(item.mDateLivraison.ToShortDateString());
                        //EcrireDansExcelDATE(procxls, i, j + 8, datLivraison);


                        //Montant HT
                      //  EcrireDansExcel(procxls, i, j + 8, (Math.Round(item.mMontantHT)).ToString());
                        worksheet.Cells[ligne, col + 8] = (Math.Round(item.mMontantHT)).ToString();

                        //MB
                     //   EcrireDansExcel(procxls, i, j + 9, (Math.Round(item.mMargeBrute)).ToString());
                        worksheet.Cells[ligne, col + 9] = (Math.Round(item.mMargeBrute)).ToString();

                        //Pays Revendeur
                        worksheet.Cells[ligne, col + 10] = item.mPaysRevendeur;

                        //Reference
              
                        worksheet.Cells[ligne, col + 11] = item.mReference;

                        //entete1
                        worksheet.Cells[ligne, col + 12] = item.mDO_Coord01;


                        ligne += 1;
                        }

                        IsOK = true;

                    

                    if (ListFC.Count == 0)
                    {
                        IsOK = true;
                    }

                }
                return IsOK;

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "Reporting HPI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Services -> FormatageForecastCom ->FormatForecast -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return IsOK;

            }

        }

        public bool FormatForecastBUNDLEECRIRE(List<FCommercial> ListFC, Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            bool IsOK = false;
            
            try
            {
                    int ligne = 2;//Ligne 
                    int col = 1;//Colonne

                    if (ListFC.Count > 0)
                    {
                        //On a des PRS,on écrit dans la feuille excel DATA
                        foreach (var item in ListFC)
                        {
                            //Numero piece(ligne=2 ;colonne=1)
                         //   EcrireDansExcel(procxls, i, j, item.mNumPiece);
                        worksheet.Cells[ligne, col] = item.mNumPiece;

                        //Revendeur
                      //  EcrireDansExcel(procxls, i, j + 1, item.mRevendeur);
                        worksheet.Cells[ligne, col+1] = item.mRevendeur;
                        
                        //Famille
                       // EcrireDansExcel(procxls, i, j + 2, item.mFamille);
                        worksheet.Cells[ligne, col + 2] = item.mFamille;

                        //Nom et prenom

                       // EcrireDansExcel(procxls, i, j + 3, item.mPrenomCommercial + " " + item.mNomCommercial);
                        worksheet.Cells[ligne, col + 3] = item.mPrenomCommercial + " " + item.mNomCommercial;

                        //Pays
                        //  EcrireDansExcel(procxls, i, j + 4, item.mPays);
                        worksheet.Cells[ligne, col + 4] = item.mPays;
                        
                        //Type Doc
                       // EcrireDansExcel(procxls, i, j + 5, item.mTypeDoc);
                        worksheet.Cells[ligne, col + 5] = item.mTypeDoc;
                        
                        //Statut Doc
                      //  EcrireDansExcel(procxls, i, j + 6, item.mStatut);
                        worksheet.Cells[ligne, col + 6] = item.mStatut;


                        //Date Pièce
                        DateTime datpiece = DateTime.Parse(item.mDatePiece.ToShortDateString());
                        //  EcrireDansExcelDATE(procxls, i, j + 7, datpiece);
                        worksheet.Cells[ligne, col + 7] = datpiece;

                        //Date Livraison
                        //DateTime datLivraison = DateTime.Parse(item.mDateLivraison.ToShortDateString());
                        //EcrireDansExcelDATE(procxls, i, j + 8, datLivraison);
                        
                        //Montant HT
                       // EcrireDansExcel(procxls, i, j + 8, (Math.Round(item.mMontantHT)).ToString());
                        worksheet.Cells[ligne, col + 8] = Math.Round(item.mMontantHT).ToString();

                        //MARGE
                        // EcrireDansExcel(procxls, i, j + 8, (Math.Round(item.mMontantHT)).ToString());
                        worksheet.Cells[ligne, col + 9] = Math.Round(item.mMargeBrute).ToString();

                        //Pays Revendeur
                        //    EcrireDansExcel(procxls, i, j + 9, (Math.Round(item.mMargeBrute)).ToString());
                        worksheet.Cells[ligne, col + 10] = item.mPaysRevendeur;

                        //Référence
                        worksheet.Cells[ligne, col + 11] = item.mReference;

                        //entete1
                        worksheet.Cells[ligne, col + 12] = item.mDO_Coord01;


                        ligne += 1;
                        }

                        IsOK = true;
                    
                }
                return IsOK;

            }
            catch (Exception ex)
            {
                //MessageBox.Show("Une erreur est survenue! Veuillez contacter votre Administrateur!", "Reporting HPI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                var msg = "Services -> FormatageForecastCom ->FormatForecastBUNDLE -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);
                return IsOK;

            }

        }





        public bool EcrireDansExcel(Microsoft.Office.Interop.Excel.Application procxcel, int Ligne, int Col, string donne)
        {
            bool IsOK = false;
            try
            {

                procxcel.Cells[Ligne, Col].Value = donne;
                IsOK = true;
                return IsOK;
            }
            catch (Exception ex)
            {
                var msg = "Services -> FormatageForecastCom ->EcrireDansExcel -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

                return IsOK;
            }


        }

        public bool EcrireDansExcelDATE(Microsoft.Office.Interop.Excel.Application procxcel, int Ligne, int Col, DateTime donne)
        {
            bool IsOK = false;
            try
            {

                procxcel.Cells[Ligne, Col].Value = donne.Date;
                IsOK = true;
                return IsOK;
            }
            catch (Exception ex)
            {
                var msg = "Services -> FormatageForecastCom ->EcrireDansExcelDATE -> TypeErreur: " + ex.Message; ;
                CAlias.Log(msg);

                return IsOK;
            }


        }

    }
}
