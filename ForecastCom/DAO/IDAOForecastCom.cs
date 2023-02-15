using ForecastCom.Models;
using ForecastCom.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForecastCom.DAO
{
    public interface IDAOForecastCom
    {
        List<FCommercial> GetForecast(string chaineconnexion, string dateDebutGen, string dateFinGen, List<CAlias> LAliasChoisis);

    }
}
