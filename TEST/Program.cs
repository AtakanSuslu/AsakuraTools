using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using Modelleyici;
using System.Data.SqlClient;
namespace TEST
{
    class Program
    {
        static void Main(string[] args)
        {
            SQLTEST();
        }
        static void SQLTEST()
        {
            var con = new SqlConnection();
            con.Insert(new Models.sbptTest() { ID = 1, Isim = "atakan" });
            var k = new List<string>();
        }
        static void ExcelHucreCek()
        {
            ExcelBL bl = new ExcelBL("TEST.xls", IslemTipi.OKUMA);
            var test = bl.Hucre("Sayfa1", "A1:A3");
            bl.Kapat();
            //com.Degisken<ExcelBL>(ref k);
            //bl.EkleBaslik(bl.GetType(),"Sayfa1");
        }
    }
}
