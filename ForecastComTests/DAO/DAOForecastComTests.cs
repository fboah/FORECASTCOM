using Microsoft.VisualStudio.TestTools.UnitTesting;
using ForecastCom.DAO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ForecastCom.Utils;
using ForecastCom.Models;

namespace ForecastCom.DAO.Tests
{
    [TestClass()]
    public class DAOForecastComTests
    {
        [TestMethod()]
        public void GetForecastTest()
        {

            DAOForecastCom dao = new DAOForecastCom();

            var LAliasChoisis = new List<CAlias>();

            var LResult = new List<FCommercial>();

            var obj1 = new CAlias();

            string chaine = @"Initial Catalog=AITEKCI;Data Source=BOAH\SAGE200;Integrated Security=SSPI";

            obj1.mId = 1;
            obj1.mBDName = "AITEKCI";
            obj1.mVille = "ABidjan";
            obj1.mAliasName = @"Initial Catalog=AITEKCI;Data Source=BOAH\SAGE200;Integrated Security=SSPI";
            obj1.IsAbidjan = true;

            var obj2 = new CAlias();

            obj2.mId = 2;
            obj2.mBDName = "BIJOU";
            obj2.mVille = "Dakar";
            obj2.mAliasName = @"Initial Catalog=BIJOU;Data Source=BOAH\SAGE200;user=sa;password=2017Aitek";
            obj2.IsAbidjan = false;

            var obj3 = new CAlias();

            obj3.mId = 3;
            obj3.mBDName = "AITEKMALI";
            obj3.mVille = "Bamako";
            obj3.mAliasName = @"Initial Catalog=AITEKMALI;Data Source=BOAH\SAGE100;Integrated Security=SSPI";
            obj3.IsAbidjan = false;

            LAliasChoisis.Add(obj1);
            LAliasChoisis.Add(obj2);

            string datedeb = "18/03/2015";
            string datefin = "18/03/2015";

         //   LResult = dao.GetForecast(chaine, datedeb, datefin, LAliasChoisis);

        }

        [TestMethod()]
        public void LoadComAllTest()
        {
            DAOForecastCom dao = new DAOForecastCom();

            var LAliasChoisis = new List<CAlias>();

            var LResult = new List<ComClass>();

            var obj1 = new CAlias();

            string chaine = @"Initial Catalog=AITEKCI;Data Source=BOAH\SAGE200;Integrated Security=SSPI";

            obj1.mId = 1;
            obj1.mBDName = "AITEKCI";
            obj1.mVille = "ABidjan";
            obj1.mAliasName = @"Initial Catalog=AITEKCI;Data Source=BOAH\SAGE200;Integrated Security=SSPI";
            obj1.IsAbidjan = true;
            obj1.mPays = "CI";

            var obj2 = new CAlias();

            obj2.mId = 2;
            obj2.mBDName = "BIJOU";
            obj2.mVille = "Dakar";
            obj2.mAliasName = @"Initial Catalog=BIJOU;Data Source=BOAH\SAGE200;user=sa;password=2017Aitek";
            obj2.IsAbidjan = false;
            obj2.mPays = "BIJOULAND";

            var obj3 = new CAlias();

            obj3.mId = 3;
            obj3.mBDName = "AITEKMALI";
            obj3.mVille = "Bamako";
            obj3.mAliasName = @"Initial Catalog=AITEKMALI;Data Source=BOAH\SAGE100;Integrated Security=SSPI";
            obj3.IsAbidjan = false;

            LAliasChoisis.Add(obj1);
            LAliasChoisis.Add(obj2);
           //LAliasChoisis.Add(obj3);

            
            LResult = dao.LoadComAll(chaine, LAliasChoisis);

            //  List<ComClass> LK = new List<ComClass>();

            //var  mLK = LResult.Where(c => c.mIdPays == 1 &&  c.mIdPays == 2).ToList();

           




        }
    }
}