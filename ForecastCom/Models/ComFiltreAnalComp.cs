using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.Models
{
    public class ComFiltreAnalComp
    {

        public string mAxeAnalyse { get; set; }
        public string mPeriode { get; set; }
        public string mPeriodeAnalyse { get; set; }
        public string mPeriodeAnalyseDefinition { get; set; }
        public string mElementAnalyse { get; set; }
        public string mCommercialCritere { get; set; }
        public string mFamilleCritere { get; set; }
        public string mMarqueCritere { get; set; }
        public string mRevendeurCritere { get; set; }
        public string mTypeDoc { get; set; }
        public string mSite { get; set; }
       
        

        public ComFiltreAnalComp()
        {
            mPeriode = string.Empty;
          mPeriodeAnalyse = string.Empty;
            mPeriodeAnalyseDefinition = string.Empty;
            mElementAnalyse = string.Empty;
            mCommercialCritere = string.Empty;
            mFamilleCritere = string.Empty;
            mMarqueCritere = string.Empty;
            mRevendeurCritere = string.Empty;
            mTypeDoc = string.Empty;
            mSite = string.Empty;




        }

}
}
