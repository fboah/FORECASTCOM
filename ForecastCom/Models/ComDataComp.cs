using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.Models
{
    public class ComDataComp
    {
        public int mAnnee { get; set; }
        public string mMois { get; set; }
        public double mMontant { get; set; }
        public string mCommercial { get; set; }
        public string mMarque { get; set; }
        public string mFamille { get; set; }
        public string mRevendeur { get; set; }


        public ComDataComp()
        {
            mMois = string.Empty;
            mCommercial = string.Empty;
            mMarque = string.Empty;
            mFamille = string.Empty;
            mRevendeur = string.Empty;
            mMontant = 0;
            mAnnee = 0;


        }

    }
}
