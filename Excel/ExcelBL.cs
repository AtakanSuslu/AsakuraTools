using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.IO;
using Modelleyici;
using System.Reflection;

namespace Excel
{
    public enum IslemTipi
    {
        OKUMA, YAZMA
    }
    public class ExcelBL
    {
        OleDbCommand com;
        OleDbConnection con;
        public IslemTipi IslemTipi { get; set; }
        public string DosyaYolu { get; set; }
        public ExcelBL(string DosyaYolu, IslemTipi IslemTipi)
        {
            this.DosyaYolu = DosyaYolu;
            this.IslemTipi = IslemTipi;
            ConnectionOlustur();
            com = con.CreateCommand();
        }
        private void ConnectionOlustur()
        {
            var Ozellikler = "";
            switch (IslemTipi)
            {
                case IslemTipi.OKUMA:
                    Ozellikler = "Excel 12.0;HDR=YES;IMEX=1;";
                    //Ozellikler = "Excel 12.0;HDR=NO;";
                    break;
                case IslemTipi.YAZMA:
                    Ozellikler = "Excel 12.0 Xml;HDR=YES;";
                    break;
            }
            con = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={DosyaYolu};Extended Properties=\"{Ozellikler}\"");
            con.Open();
        }
        public void TabloOlustur<T>(T Object, bool DosyavarsaSil = false)
        {
            TabloOlustur(typeof(T), DosyavarsaSil: DosyavarsaSil);
            Ekle(Object);
        }
        public void TabloOlustur<T>(List<T> Object, bool DosyavarsaSil = false)
        {
            TabloOlustur(typeof(T), DosyavarsaSil: DosyavarsaSil);
            Ekle(Object);
        }
        public void TabloOlustur(Type _Type, bool DosyavarsaSil = false)
        {
            if (DosyavarsaSil)
            {
                DosyaSil();
                ConnectionOlustur();
                com = con.CreateCommand();
            }
            var TabloIsmi = _Type.Name;
            com.CommandText = string.Format("create table {0} (@@)", TabloIsmi);

            var Sutunlar = new List<string>();
            foreach (var prop in _Type.GetProperties())
            {
                var Type = prop.PropertyType;
                var Isim = prop.Name;
                var Sutun = string.Format("{0} {1}", Isim, GetExcelDataType(Type));
                Sutunlar.Add(Sutun);
            }
            com.CommandText = com.CommandText.Replace("@@", String.Join(",", Sutunlar));
            com.ExecuteNonQuery();

        }
        private void TabloSil(Type _Type, bool BaslikSil = false)
        {
            var TabloIsmi = _Type.Name;
            var Sutunlar = new List<string>();
            foreach (var prop in _Type.GetProperties())
                Sutunlar.Add(string.Format("{0}=''", prop.Name));
            com.CommandText = $"update [{TabloIsmi}$] set {string.Join(",", Sutunlar)}";
            com.ExecuteNonQuery();
            com.Parameters.Clear();
            if (BaslikSil)
            {
                com.CommandText = $"drop table [{TabloIsmi}$]";
                com.ExecuteNonQuery();
            }
        }
        public void DosyaSil()
        {
            Kapat();
            File.Delete(DosyaYolu);
        }
        private void TabloSil<T>(T Object, bool BaslikSil = false)
        {
            TabloSil(typeof(T), BaslikSil: BaslikSil);
        }
        private void TabloSil<T>(List<T> Object, bool BaslikSil = false)
        {
            TabloSil(typeof(T), BaslikSil: BaslikSil);
        }
        public List<Dictionary<string, dynamic>> Tablo(string TabloIsmi, string Range = "", bool Trim = true, bool WhiteSpace = false)
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
                    if (Kolon == "Müş.Sip.Fiş Kodları")
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



                    Kayit.Add(Kolon.ToLower(), Deger);
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
        public List<T> Tablo<T>(string TabloIsmi = "", string Range = "", bool ConvertType = true)
        {
            if (string.IsNullOrEmpty(TabloIsmi))
                TabloIsmi = typeof(T).Name;
            com.CommandText = $"select * from [{TabloIsmi}${Range}]";
            var y = new List<string>();
            var Sonuc = com.Liste<T>(ref y, ConvertType: ConvertType);
            for (int i = 0; i < Sonuc.Count; i++)
            {
                int j = 0;
                var propCount = typeof(T).GetProperties().Count();
                foreach (PropertyInfo item in typeof(T).GetProperties())
                {
                    var val = item.GetValue(Sonuc[i]);
                    var val2 = item.GetValue(Activator.CreateInstance<T>());
                    if (object.Equals(val, val2))
                        j++;
                }
                if (j == propCount)
                    Sonuc.RemoveRange(i, Sonuc.Count - i);
            }
            return Sonuc;
        }
        /// <summary>
        /// Girilen sutunlar arasındaki kayıtları çek
        /// Kullanılmıyor
        /// </summary>
        /// <typeparam name="T">Model Tipi</typeparam>
        /// <param name="TabloIsmi">Kayıtların çekilmek istendiği tablo ismi. Boş geçilirse modelin adı yazılır.</param>
        /// <param name="Range">Çekilmek istenen kayıtlar hangi sutunlar arasında. (A1:F3)</param>
        /// <param name="Trim">Boşluklar silinsin mi</param>
        /// <returns></returns>
        public List<T> TabloEski<T>(string TabloIsmi = "", string Range = "", bool Trim = true, bool WhiteSpace = false)
        {
            if (string.IsNullOrEmpty(TabloIsmi))
                TabloIsmi = typeof(T).Name;
            List<T> Sonuc = new List<T>();
            com.CommandText = $"select * from [{TabloIsmi}${Range}]";
            var y = new List<string>();
            var sonuc = com.Liste<T>(ref y);
            con.Open();
            var rdr = com.ExecuteReader();
            while (rdr.Read())
            {
                var Kontrol = false;
                var Kayit = Activator.CreateInstance<T>();
                foreach (var prop in Kayit.GetType().GetProperties())
                {
                    object Deger = null;
                    if (prop.CustomAttributes.Count(x => x.AttributeType.Name.Equals("atrTabloDisi")) > 0)
                        continue;
                    try
                    {
                        var SutunIsmi = prop.Name;
                        if (WhiteSpace)
                            SutunIsmi.Replace("_", " ");
                        var test = rdr.GetOrdinal(SutunIsmi);
                        Deger = rdr[SutunIsmi];

                        if (Deger != null)
                            Deger = Trim ? Deger.ToString().Trim() : Deger.ToString();

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
                        if (prop.GetValue(Kayit) != null && prop.GetValue(Kayit).Equals(""))
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
        /// HDR NO
        /// </summary>
        /// <param name="TabloIsmi">Guncellenmek istenen tablo ismi.</param>
        /// <param name="Range">Guncellenmek istenen hucrenin sutunu. (A1:A1)</param>
        /// <param name="Deger">Yeni değer</param>
        public void GuncelleHucre(string TabloIsmi, string Range, string Deger)
        {
            com.CommandText = $"update [{TabloIsmi}${Range}] set F1='{Deger}'";
            com.ExecuteNonQuery();
        }
        /// <summary>
        /// İstenen hücreyi günceller
        /// HDR YES
        /// </summary>
        /// <param name="sql"></param>
        public void GuncelleHucre(string sql)
        {
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
            Tablo.ForEach(x => Ekle(x));
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
                var oDeger = prop.GetValue(Kayit);
                var Deger = "";
                if (oDeger != null)
                    Deger = oDeger.ToString();
                Alanlar += prop.Name + ",";
                com.Parameters.AddWithValue(string.Format("@{0}", prop.Name), Deger);
                Degerler += "?,";
                //Degerler += "\'" + prop.GetValue(Kayit).ToString() + "\',";
            }
            Alanlar = Alanlar.Substring(0, Alanlar.Length - 1);
            Degerler = Degerler.Substring(0, Degerler.Length - 1);
            com.CommandText = $"insert into [{TabloIsmi}$] ({Alanlar}) values ({Degerler})";
            com.ExecuteNonQuery();
            com.Parameters.Clear();
        }
        /// <summary>
        /// Excel Dosyasına Başlık Girişi Yap. Eklenmek istenen modelin alanları ilk satır olarak girilir. 
        /// </summary>
        /// <param name="Kayit">Eklenicek kayıtların tipi.</param>
        /// <param name="TabloIsmi">Kayıtların eklenmek istendiği tablo ismi. Boş geçilirse modelin adı yazılır.</param>
        public void EkleBaslik<T>(T Kayit, string TabloIsmi = "")
        {
            if (TabloIsmi.Equals(""))
                TabloIsmi = Kayit.GetType().Name;
            int i = 1;
            var Alanlar = new List<string>();
            var Degerler = new List<string>();
            foreach (var prop in typeof(T).GetProperties())
            {
                Alanlar.Add($"[F{i++}]");
                Degerler.Add($"?");
                com.Parameters.AddWithValue($"@{prop.Name}", prop.Name);
            }
            com.CommandText = $"insert into [{TabloIsmi}$] ({string.Join(",", Alanlar)})values({string.Join(",", Degerler)})";
            var rr = com.ExecuteNonQuery();
            com.Parameters.Clear();
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

        private string GetExcelDataType(Type type)
        {
            var name = type.Name;
            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                name = type.GetGenericArguments()[0].Name;
            }

            switch (name)
            {
                case "Guid":
                    return "text";
                case "Boolean":
                    return "number";
                case "Byte":
                    return "number";
                case "Int16":
                    return "number";
                case "Int32":
                    return "number";
                case "Int64":
                    return "number";
                case "Decimal":
                    return "number";
                case "Single":
                    return "number";
                case "Double":
                    return "number";
                case "DateTime":
                    return "datetime";
                case "String":
                case "Char[]":
                    return "Memo";
                case "Char":
                    return "text";
                case "Byte[]":
                    return "text";
                case "Object":
                    return "text";
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
    }
}
