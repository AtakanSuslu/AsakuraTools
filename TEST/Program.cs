using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using Modelleyici;
using System.Data.SqlClient;
using TEST.Models;
using Newtonsoft.Json;

namespace TEST
{
    class Program
    {
        static void Main(string[] args)
        {
            SqlSelectTest();
            SQLInserTest();
            SQLUpdateTest();
            SQLDeleteTest();
        }
        static void ExcelYaz()
        {
            ExcelBL bl = new ExcelBL("text.xls", IslemTipi.YAZMA);
            bl.Kapat();
        }
        static void SqlSelectTest()
        {
            using (var con = new SqlConnection(""))
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
                var user = com.Degisken<tUser>(ref y);
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
            using (var con = new SqlConnection(""))
            {
                var UserID = con.Insert(user);
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
                var EfectedRowsCount = con.Update(user);
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
        static void serializeobject()
        {
            var k = new k1()
            {
                //a1 = 1,
                a2 = "1",
                a3 = new List<int>() { 1, 23, 4 },
                a4 = new List<string>() { "asd", "ghtdfh", "324234", "asd" },
                a5 = new int[] { 12, 3, 5, 6457, 456 },
                a6 = new string[] { "asda", "asdasda", "a" },
                a7 = new k2()
                {
                    a1 = 1,
                    a2 = "1",
                    a3 = new List<int>() { 1, 23, 4 },
                    a4 = new List<string>() { "asd", "ghtdfh", "324234", "asd" },
                    a5 = new int[] { 12, 3, 5, 6457, 456 },
                    a6 = new string[] { "asda", "asdasda", "a" }
                },
                a8 = new List<k2>() { new k2()
                {
                    a1 = 1,
                    a2 = "1",
                    a3 = new List<int>() { 1, 23, 4 },
                    a4 = new List<string>() { "asd", "ghtdfh", "324234", "asd" },
                    a5 = new int[] { 12, 3, 5, 6457, 456 },
                    a6 = new string[] { "asda", "asdasda", "a" },
                },new k2()
                {
                    a1 = 1,
                    a2 = "1",
                    a3 = new List<int>() { 1, 23, 4 },
                    a4 = new List<string>() { "asd", "ghtdfh", "324234", "asd" },
                    a5 = new int[] { 12, 3, 5, 6457, 456 },
                    a6 = new string[] { "asda", "asdasda", "a" },
                },new k2()
                {
                    a1 = 1,
                    a2 = "1",
                    a3 = new List<int>() { 1, 23, 4 },
                    a4 = new List<string>() { "asd", "ghtdfh", "324234", "asd" },
                    a5 = new int[] { 12, 3, 5, 6457, 456 },
                    a6 = new string[] { "asda", "asdasda", "a" },
                }}
            };
            var kk = new k1() { a2="asd"};
            var aaa = Cevir.JsonSerializeObject(k,IgnoreIfNull:true);
            var a1aa = Cevir.JsonSerializeObject(kk, IgnoreIfNull: true);
            var bb=JsonConvert.SerializeObject(k);
            var a = JsonConvert.DeserializeObject<k1>(bb);
            GC.Collect();
        }

        public class k1
        {
            public int? a1 { get; set; }
            public string a2 { get; set; }
            public List<int> a3 { get; set; }
            public List<string> a4 { get; set; }
            public int[] a5 { get; set; }
            public string[] a6 { get; set; }
            public k2 a7 { get; set; }
            public List<k2> a8 { get; set; }
        }
        public class k2
        {
            public int a1 { get; set; }
            public string a2 { get; set; }
            public List<int> a3 { get; set; }
            public List<string> a4 { get; set; }
            public int[] a5 { get; set; }
            public string[] a6 { get; set; }
        }
    }
}
