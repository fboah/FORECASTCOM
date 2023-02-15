using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.Models
{
    public class ComFamille
    {
        public string mFa_CodeFamille { get; set; }
        public string mFa_Central { get; set; }
        public string mPays { get; set; }
        public int mIdPays { get; set; }
        public int mFa_Type { get; set; }
        public string mFa_Intitule { get; set; }


        public ComFamille()
        {
            mFa_CodeFamille = string.Empty;
            mPays = string.Empty;
            mIdPays = 0;
            mFa_Type = 0;
            mFa_Intitule = string.Empty;
            mFa_Central = string.Empty;

        }
    }
}
