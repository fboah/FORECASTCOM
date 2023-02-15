using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.Models
{
    public class ComRevendeur
    {
        public string mCT_Num { get; set; }
        public string mPays { get; set; }
        public int mIdPays { get; set; }
        public string mCT_Intitule { get; set; }
        public string mCollectif { get; set; }
      //  public DateTime mDateCreate { get; set; }
        public string mFamille { get; set; }
        public string mFamilleCentral { get; set; }
        public string mNomCommercial { get; set; }
        public string mPrenomCommercial { get; set; }
        public string mTypeDoc { get; set; }
        public int mTypeDocId { get; set; }
        public DateTime mDatePiece { get; set; }
        public double mMontantHT { get; set; }
        public double mMargeBrute { get; set; }
        public double mMontantSaisieDevis { get; set; }//60%
        public double mMontantAccepteDevis { get; set; }//75%
        public double mMontantTotalPipeDevis { get; set; }//Somme des deux
        public double mMontantTotalHorsCasPipe { get; set; }
        public int mStatutId { get; set; }
        public string mStatut { get; set; }


        public ComRevendeur()
        {
            mCT_Num = string.Empty;
            mPays = string.Empty;
            mIdPays = 0;
            mCT_Intitule = string.Empty;
            mCollectif = string.Empty;
         //   mDateCreate = new DateTime();
            mFamille = string.Empty;
            mFamilleCentral = string.Empty;
            mNomCommercial = string.Empty;
            mPrenomCommercial = string.Empty;
            mTypeDoc = string.Empty;
            mTypeDocId = 0;
            mDatePiece = new DateTime();
            mMontantHT = 0;
            mMargeBrute = 0;
            mMontantSaisieDevis = 0;
            mMontantAccepteDevis = 0;
            mMontantTotalPipeDevis = 0;
        mMontantTotalHorsCasPipe = 0;
            mStatutId = 0;
        mStatut = string.Empty;
    }


    }
}
