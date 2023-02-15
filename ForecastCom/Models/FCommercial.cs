using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.Models
{
    public class FCommercial
    {
        public string mNumPiece { get; set; }
        public string mRevendeur { get; set; }
        public string mFamille { get; set; }
        public string mFamilleCentral { get; set; }
        public string mNomCommercial { get; set; }
        public string mPrenomCommercial { get; set; }
        public string mPays { get; set; }
        //public string mAr_Ref { get; set; }
        //public string mDl_Design { get; set; }
        public string mTypeDoc { get; set; }
        public int mTypeDocId { get; set; }
        public DateTime mDatePiece { get; set; }
        public double mMontantHT { get; set; }
        public double mMargeBrute { get; set; }
        public string mPaysCommerciale { get; set; }
        public int mStatutId { get; set; }
        public string mStatut { get; set; }
        public DateTime mDateLivraison { get; set; }
        public double mMontantSaisieDevis { get; set; }//60%
        public double mMontantAccepteDevis { get; set; }//75%
        public double mMontantTotalPipeDevis { get; set; }//Somme des deux
        public double mMontantTotalHorsCasPipe { get; set; }
        public double mMontantTotalMarge { get; set; }
        public string mReference { get; set; }//Reference document
        public string mPaysRevendeur { get; set; }//Reference document
        public int mNbreDoc { get; set; }
        public double mQtiteFamille { get; set; }

        public string mDO_Coord01 { get; set; }


        public FCommercial()
        {

            mNumPiece = string.Empty;
            mRevendeur = string.Empty;
            mFamille = string.Empty;
            mFamilleCentral = string.Empty;
            mNomCommercial = string.Empty;
            mPrenomCommercial = string.Empty;
            mPays = string.Empty;
            //mAr_Ref = string.Empty;
            //mDl_Design = string.Empty;
            mTypeDoc = string.Empty;
            mTypeDocId = -1;
            mDatePiece = new DateTime();
            mMontantHT = 0;
            mMargeBrute = 0;
            mPaysCommerciale = string.Empty;
            mStatutId = -1;
            mDateLivraison= new DateTime();
            mStatut = string.Empty;
            mMontantSaisieDevis = 0;
            mMontantAccepteDevis = 0;
            mMontantTotalPipeDevis = 0;
            mMontantTotalHorsCasPipe = 0;
            mReference = string.Empty;
            mNbreDoc = 0;
            mQtiteFamille = 0;
            mMontantTotalMarge = 0;
            mPaysRevendeur = string.Empty;
            mDO_Coord01 = string.Empty;

    }








    }
}
