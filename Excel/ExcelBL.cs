using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Modelleyici;

namespace Excel
{
    public enum IslemTipi
    {
        OKUMA, YAZMA
    }
    public class ExcelBL
    {
        OleDbCommand com = new OleDbCommand();
        OleDbConnection con;
        public IslemTipi IslemTipi { get; set; }
        public string DosyaYolu { get; set; }
        public ExcelBL(string DosyaYolu, IslemTipi IslemTipi)
        {
            this.DosyaYolu = DosyaYolu;
            this.IslemTipi = IslemTipi;
            var Ozellikler = "";
            switch (IslemTipi)
            {
                case IslemTipi.OKUMA:
                    Ozellikler = "Excel 12.0;HDR=YES;IMEX=1;";
                    //Ozellikler = "Excel 12.0;HDR=NO;";
                    break;
                case IslemTipi.YAZMA:
                    Ozellikler = "Excel 12.0;HDR=NO;";
                    if (!File.Exists(DosyaYolu))
                        File.Create(DosyaYolu);
                    break;
            }

            con = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={DosyaYolu};Extended Properties=\"{Ozellikler}\"");
            con.Open();
            com.Connection = con;
        }
        public List<Dictionary<string,dynamic>> Tablo(string TabloIsmi, string Range = "", bool Trim = true, bool WhiteSpace = false)
        {
            com.CommandText = $"select * from [{TabloIsmi}${Range}]";
            return com.Liste();
        }
        public List<Dictionary<string, dynamic>> ORKUN(string TabloIsmi, string Range = "", bool Trim = true, bool WhiteSpace = false)
        {
            com.CommandText = $"select * from [{TabloIsmi}${Range}]";
            if (com.Connection.State == System.Data.ConnectionState.Closed)
                com.Connection.Open();
            var rdr = com.ExecuteReader();
            var Sonuc = new List<Dictionary<string, dynamic>>();
            var KolonSayisi = rdr.FieldCount;
            while (rdr.Read())
            {
                var Kayit = new Dictionary<string, dynamic>();
                for (int i = 0; i < KolonSayisi; i++)
                {
                    var Tip = rdr.GetFieldType(i);
                    var Kolon = rdr.GetName(i);
                    var oDeger = rdr.GetValue(i);
                    object Deger = null;
                    if (Kolon== "Müş.Sip.Fiş Kodları")
                    {
                        Type t = typeof(string);
                        var gereksiz = 0;
                        if (int.TryParse(Deger.ToString(), out gereksiz))
                            t = typeof(int);
                        if (oDeger != null && oDeger != DBNull.Value)
                            Deger = Convert.ChangeType(oDeger, t);
                    }
                    else
                    {
                        if (oDeger != null && oDeger != DBNull.Value)
                            Deger = Convert.ChangeType(oDeger, Tip);
                    }
                
                  
                    
                    Kayit.Add(Kolon.ToLower(),Deger);
                }
                Sonuc.Add(Kayit);
            }
            rdr.Close();
            return Sonuc;

        }
        /// <summary>
        /// Girilen sutunlar arasındaki kayıtları çek
        /// </summary>
        /// <typeparam name="T">Model Tipi</typeparam>
        /// <param name="TabloIsmi">Kayıtların çekilmek istendiği tablo ismi. Boş geçilirse modelin adı yazılır.</param>
        /// <param name="Range">Çekilmek istenen kayıtlar hangi sutunlar arasında. (A1:F3)</param>
        /// <param name="Trim">Boşluklar silinsin mi</param>
        /// <returns></returns>
        public List<T> Tablo<T>(string TabloIsmi="", string Range = "", bool Trim = true,bool WhiteSpace=false)
        {
            if (string.IsNullOrEmpty(TabloIsmi))
                TabloIsmi = typeof(T).Name;
            List<T> Sonuc = new List<T>();
            com.CommandText = $"select * from [{TabloIsmi}${Range}]";
            var rdr = com.ExecuteReader();
            while (rdr.Read())
            {
                var Kontrol = false;
                var Kayit = Activator.CreateInstance<T>();
                foreach (var prop in Kayit.GetType().GetProperties())
                {
                    object Deger = null;
                    if (prop.CustomAttributes.Count(x=>x.AttributeType.Name.Equals("atrTabloDisi"))>0)
                        continue;
                    try
                    {
                        var SutunIsmi = prop.Name;
                        if (WhiteSpace)
                            SutunIsmi.Replace("_", " ");
                        Deger = Trim ? rdr[SutunIsmi].ToString().Trim() : rdr[SutunIsmi].ToString();
                    }
                    catch (Exception e)
                    {

                    }


                    if (Deger != null)
                    {
                        prop.SetValue(Kayit, Convert.ChangeType(Deger, prop.PropertyType));
                        Kontrol = true;
                    }
                }
                if (Kontrol)
                {
                    int bos = 0;
                    foreach (var prop in Kayit.GetType().GetProperties())
                    {
                        if (prop.GetValue(Kayit)!=null&&prop.GetValue(Kayit).Equals(""))
                        {
                            bos++;
                            //prop.SetValue(Kayit, "11");
                        }
                    }
                    if ((TabloIsmi.Equals("Özellik Detay - stbOzDetay") && bos > 2) || (!TabloIsmi.Equals("Özellik Detay - stbOzDetay") && bos > 0))
                    {

                    }
                   
                    Sonuc.Add(Kayit);
                }

            }
            rdr.Close();
            return Sonuc;
        }
        /// <summary>
        /// Girilen sutunlar arasındaki tek hucre içeriğini çek
        /// </summary>
        /// <param name="TabloIsmi">Hucrenin çekilmek istendiği tablo ismi.</param>
        /// <param name="Range">Çekilmek istenen hucrenin sutunu. (A1:A1)</param>
        /// <param name="Trim">Boşluklar silinsin mi?</param>
        /// <returns></returns>
        public string Hucre(string TabloIsmi, string Range, bool Trim = false)
        {

            var Sonuc = "";
            com.CommandText = $"select * from [{TabloIsmi}${Range}]";
            var rdr = com.ExecuteReader();
            if (rdr.Read())
                Sonuc = Trim ? rdr[0].ToString().Trim() : rdr[0].ToString();
            rdr.Close();
            return Sonuc;
        }
        /// <summary>
        /// Girilen sutunlar arasındaki kaydı çek
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="TabloIsmi">Kaydın çekilmek istendiği tablo ismi. Boş geçilirse modelin adı yazılır.</param>
        /// <param name="Range">Çekilmek istenen kayıt hangi sutunlar arasında. (A1:A3)</param>
        /// <returns></returns>
        public T Kayit<T>(string Range, string TabloIsmi = "")
        {
            var Sonuc = Activator.CreateInstance<T>();
            if (TabloIsmi.Equals(""))
                TabloIsmi = Sonuc.GetType().Name;
            com.CommandText = $"select * from [{TabloIsmi}${Range}]";
            var rdr = com.ExecuteReader();
            if (rdr.Read())
                foreach (var prop in Sonuc.GetType().GetProperties())
                    prop.SetValue(Sonuc, rdr[prop.Name].ToString());
            rdr.Close();
            return Sonuc;
        }
        /// <summary>
        /// İstenen hücreyi günceller
        /// </summary>
        /// <param name="TabloIsmi">Guncellenmek istenen tablo ismi.</param>
        /// <param name="Range">Guncellenmek istenen hucrenin sutunu. (A1:A1)</param>
        /// <param name="Deger">Yeni değer</param>
        public void GuncelleHucre(string TabloIsmi, string Range, string Deger)
        {
            //HDR NO
            com.CommandText = $"update [{TabloIsmi}${Range}] set F1='{Deger}'";
            com.ExecuteNonQuery();
        }
        /// <summary>
        /// İstenen hücreyi günceller
        /// </summary>
        /// <param name="sql"></param>
        public void GuncelleHucre(string sql)
        {
            //HDR YES
            com.CommandText = sql;
            com.ExecuteNonQuery();
        }
        /// <summary>
        /// Liste halindeki verilen modelin alanlarına göre excele kayıt atar.
        /// </summary>
        /// <typeparam name="T">Kayıtların model tipi.</typeparam>
        /// <param name="Tablo">Eklenicek Kayıtların listesi.</param>
        /// <param name="TabloIsmi">Kayıtların eklenmek istendiği tablo ismi. Boş geçilirse modelin adı yazılır.</param>
        public void Ekle<T>(List<T> Tablo, string TabloIsmi = "")
        {
            //EkleBaslik(Tablo.FirstOrDefault().GetType(), TabloIsmi);
            if (TabloIsmi.Equals(""))
                TabloIsmi = Tablo.FirstOrDefault().GetType().Name;
            foreach (var Kayit in Tablo)
            {
                var Alanlar = "";
                var Degerler = "";
                foreach (var prop in Kayit.GetType().GetProperties())
                {
                    Alanlar += prop.Name + ",";
                    Degerler += "\'" + prop.GetValue(Kayit).ToString() + "\',";
                }
                Alanlar = Alanlar.Substring(0, Alanlar.Length - 1);
                Degerler = Degerler.Substring(0, Degerler.Length - 1);
                com.CommandText = $"insert into [{TabloIsmi}$] ({Alanlar}) values ({Degerler})";
                com.ExecuteNonQuery();
            }

        }
        /// <summary>
        /// Verilen modelin alanlarına göre excele kayıt atar.
        /// </summary>
        /// <typeparam name="T">Model Tipi</typeparam>
        /// <param name="Kayit">Veri tabanına eklenicek kayit.</param>
        /// <param name="TabloIsmi">Kayıtların eklenmek istendiği tablo ismi. Boş geçilirse modelin adı yazılır.</param>
        public void Ekle<T>(T Kayit, string TabloIsmi = "")
        {
            if (TabloIsmi.Equals(""))
                TabloIsmi = Kayit.GetType().Name;

            var Alanlar = "";
            var Degerler = "";
            foreach (var prop in Kayit.GetType().GetProperties())
            {
                Alanlar += prop.Name + ",";
                Degerler += "'" + prop.GetValue(Kayit).ToString() + "',";
            }
            Alanlar = Alanlar.Substring(0, Alanlar.Length - 1);
            Degerler = Degerler.Substring(0, Degerler.Length - 1);
            com.CommandText = $"insert into [{TabloIsmi}$] ({Alanlar}) values ({Degerler})";
            com.ExecuteNonQuery();
        }
        /// <summary>
        /// Excel Dosyasına Başlık Girişi Yap. Eklenmek istenen modelin alanları ilk satır olarak girilir. 
        /// </summary>
        /// <param name="Kayit">Eklenicek kayıtların tipi.</param>
        /// <param name="TabloIsmi">Kayıtların eklenmek istendiği tablo ismi. Boş geçilirse modelin adı yazılır.</param>
        public void EkleBaslik(Type Kayit, string TabloIsmi = "")
        {
            
            var _con = con.ConnectionString;
            con.Close();
            var Ozellikler  = "Excel 12.0;HDR=NO;";
            con.ConnectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={DosyaYolu};Extended Properties=\"{Ozellikler}\"";
            con.Open();
            com.Connection = con;

            if (TabloIsmi.Equals(""))
                TabloIsmi = Kayit.GetType().Name;

            var Alanlar = "";
            var Degerler = "";
            int i = 1;
            foreach (var prop in Kayit.GetProperties())
            {
                Alanlar += $"[F{i++}],";
                Degerler += $"'{ prop.Name}',";
            }
            Alanlar = Alanlar.Substring(0, Alanlar.Length - 1);
            Degerler = Degerler.Substring(0, Degerler.Length - 1);
            com.CommandText = $"insert into [{TabloIsmi}$] (F1,F2,F3)values(1,2,3)";
            com.ExecuteNonQuery();
            con.Close();

            con.ConnectionString = _con;
            con.Open();
            com.Connection = con;
        }
        /// <summary>
        /// Tüm nesneleri serbest bırak
        /// </summary>
        public void Kapat()
        {
            com.Dispose();
            con.Close();
            con.Dispose();
        }
    }
}
