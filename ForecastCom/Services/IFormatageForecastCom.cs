using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.Services
{
    public interface IFormatageForecastCom
    {
         bool EcrireDansFichier(List<string> msgTowrite, string CheminFichier);


    }
}
