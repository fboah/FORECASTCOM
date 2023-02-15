using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.Models
{
    public class ComFiltre
    {
        public string mPeriode { get; set; }
        public string mCommerciaux { get; set; }
        public string mFamille { get; set; }
        public string mFamilleCentral { get; set; }
        public string mRevendeur { get; set; }
        public string mTypeDoc { get; set; }
        public string mSite { get; set; }
        public bool mIsPipe { get; set; }




        public ComFiltre()
        {
            mPeriode = string.Empty;
            mCommerciaux = string.Empty;
            mFamille = string.Empty;
            mRevendeur = string.Empty;
            mTypeDoc = string.Empty;
            mSite = string.Empty;
            mIsPipe = false;
            mFamilleCentral = string.Empty;



        }
    }
}
