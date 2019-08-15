using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using Modelleyici;
using System.Data.SqlClient;
using TEST.Models;

namespace TEST
{
    class Program
    {
        static void Main(string[] args)
        {
            //SqlSelectTest();
            //SQLInserTest();
            //SQLUpdateTest();
            //SQLDeleteTest();
            UntitledTest();
        }
        static void SqlSelectTest()
        {
            using (var con=new SqlConnection(""))
            {
                /////////////
                var y = new List<string>();
                var us = con.Select<tUser>(ref y);
                foreach (var u in us)
                {
                    Console.WriteLine($"ID: {u.ID} Name: {u.Name}");
                }

                ////////////////
                var com = con.CreateCommand();
                com.CommandText = "select top 1 * from tUser";
                //Veri tabanındaki model ile eşleşmeyen alanlar
                y = new List<string>();
                var user=com.Degisken<tUser>(ref y);
                Console.WriteLine($"ID: {user.ID} Name: {user.Name}");

                ///////////////
                y = new List<string>();
                com.CommandText = "select * from tUser where ID>@ID";
                com.Parameters.AddWithValue("@ID", 5);
                var users = com.Liste<tUser>(ref y);
                foreach (var u in users)
                {
                    Console.WriteLine($"ID: {u.ID} Name: {u.Name}");
                }

                /////////////
                y = new List<string>();
                com.CommandText = "select * from tUser";
                com.Parameters.AddWithValue("@ID", 5);
                var DicUsers = com.Liste();
                foreach (var u in DicUsers)
                {
                    Console.WriteLine($"ID: {u["ID"]} Name: {u["Name"]}");
                }
                
            }
        }
        static void SQLInserTest()
        {
            var user = new tUser()
            {
                ID = 1,
                Name = "Atakan",
                Password = "password",
                UserName = "Asakura"
            };
            using (var con=new SqlConnection(""))
            {
                var UserID=con.Insert(user);
            }
        }
        static void SQLUpdateTest()
        {
            var user = new tUser()
            {
                ID = 1,
                Name = "Atakan Süslü",
                Password = "password",
                UserName = "Asakura"
            };
            using (var con = new SqlConnection(""))
            {
               var EfectedRowsCount=con.Update(user);
            }
        }
        static void SQLDeleteTest()
        {
            var user = new tUser()
            {
                ID = 1
            };
            using (var con = new SqlConnection(""))
            {
                var EfectedRowsCount = con.Delete(user);
            }
        }
        static void ExcelHucreCek()
        {
            ExcelBL bl = new ExcelBL("TEST.xls", IslemTipi.OKUMA);
            var test = bl.Hucre("Sayfa1", "A1:A3");
            bl.Kapat();
            //com.Degisken<ExcelBL>(ref k);
            //bl.EkleBaslik(bl.GetType(),"Sayfa1");
        }

        static void UntitledTest()
        {
            var user = new tUser
            {
                ID = 1,
                Name = "Atakan",
                Password = "1234",
                UserName = "Asakura"
            };

            var user2 = new tUser();

            user2.Modelle(user);

        }
    }
}
