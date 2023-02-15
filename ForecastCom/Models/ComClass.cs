using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.Models
{
    //Classe pour gérer le filtre sur les commerciaux(Nom ,prenoms ,Pays ,Id Pays )
    public class ComClass
    {
        public string mNomCommercial { get; set; }
        public string mPrenomCommercial { get; set; }
        public string mPays { get; set; }
        public int mIdPays { get; set; }
        public int mNumCommercial { get; set; }
        

        public ComClass()
        {
            mNomCommercial = string.Empty;
            mPrenomCommercial = string.Empty;
            mPays = string.Empty;
            mIdPays = 0;
            mNumCommercial = 0;
          

        }


    }
}
