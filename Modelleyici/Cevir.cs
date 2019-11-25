using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Modelleyici
{
    #region #1-18.04.2018
    //metotları extension metotlara dönüştürdüm 
    //Dönüş tipi belli olmayan veriler için dictionary döndüren extension metot ekledim
    //Gelen datareaderı işlemler bitince kapattım
    //Command desteği
    #endregion
    #region #2-25.04.2018
    //Veri kaynağından null gelen değerler için fix
    //Veri tabanından null gelen veriler= DBNull.value
    #endregion
    #region #3-19.07.2018
    //Tablodan verilen modele göre kayıt seti çekerken reflection yöntemi yerine dictionary yöntemi ilen çalışan metotdan reflection yaptım -- eğer dictionaryde key varsa modele geçir
    //Tablodan veri çekerken sutun ismini büyük küçük harfe göre duyarlı olabilicek geliştirme 
    //Modele göre kayıt getirirken sorgudan çekip te modele aktarılamayan sutunları gösterme geliştirmesi
    #endregion
    #region #4-12.03.2019
    //Değişken fonksiyonunda veritabanından değer dönmediği zaman çıkan null hata düzeltmesi
    #endregion
    #region #5-13.03.2019
    //rdr (sqldatareader) işlem bittikten sonra kapatıldı
    #endregion
    #region #6-15.03.2019
    //Connection nesnesinden generic classlar ile insert update delete extension metodları yazıldı
    //Liste extension metodunu Connection sınıfından çağıracak Select Extensionu yazıldı
    #endregion
    #region #7-03.04.2019
    //Connection nesnesinden generic classlar ile Değişken(tek kayıt) extension metodları yazıldı
    #endregion
    #region #8-16.06.2019
    //NotMappedAttribute attribute ile normalde veritabanında olmayan alanlardan sorun çıkması engellendi
    #endregion
    public static class Cevir
    {
        #region DataReader
        /// <summary>
        /// Veri tabanındaki kaydı istenen tipe göre döndürür.
        /// </summary>
        /// <typeparam name="T">Geri Dönüş Tipi</typeparam>
        /// <param name="rdr">Veri tabanı Datareader nesnesi</param>
        /// <returns></returns>
        public static T Degisken<T>(this DbDataReader rdr, ref List<string> YakalanamayanAlanlar, bool BuyukKucukHarfDuyarli = true) where T : class
        {
            var DicKayit = rdr.Degisken(BuyukKucukHarfDuyarli);
            var Kayit = Activator.CreateInstance<T>();
            if (DicKayit == null)
                return null;
            var Propeties = Kayit.GetType().GetProperties();
            foreach (var prop in Propeties)
            {
                if (prop.GetCustomAttributes(typeof(NotMappedAttribute), false).Count() > 0) continue;
                dynamic val;
                if (DicKayit.TryGetValue(BuyukKucukHarfDuyarli ? prop.Name : prop.Name.ToLower(), out val))
                    prop.SetValue(Kayit, val);
            }
            YakalanamayanAlanlar.AddRange(DicKayit.Keys.ToList().Where(x => Propeties.FirstOrDefault(y => (BuyukKucukHarfDuyarli ? y.Name : y.Name.ToLower()).Equals(x)) == null).ToArray());
            return Kayit;
        }
        /// <summary>
        /// Veri tabanındaki kaydı dictionary şeklinde döndürür
        /// </summary>
        /// <param name="rdr">Veri tabanı Datareader nesnesi</param>
        /// <returns></returns>
        public static Dictionary<string, dynamic> Degisken(this DbDataReader rdr, bool BuyukKucukHarfDuyarli = true)
        {

            if (!rdr.Read())
                return null;
            var KolonSayisi = rdr.FieldCount;
            var Sonuc = new Dictionary<string, dynamic>();
            for (int i = 0; i < KolonSayisi; i++)
            {
                var Tip = rdr.GetFieldType(i);
                var Kolon = rdr.GetName(i);
                var oDeger = rdr.GetValue(i);
                object Deger = null;
                if (oDeger != null && oDeger != DBNull.Value)
                    Deger = Convert.ChangeType(oDeger, Tip);
                Sonuc.Add(BuyukKucukHarfDuyarli ? Kolon : Kolon.ToLower(), Deger);
            }
            rdr.Close();
            return Sonuc;
        }
        /// <summary>
        /// Veri tabanındaki kayıtları istenen tipe göre liste şeklinde döndürür
        /// </summary>
        /// <typeparam name="T">Geri Dönüş Tipi</typeparam>
        /// <param name="rdr">Veri tabanı Datareader nesnesi</param>
        /// <returns></returns>
        public static List<T> Liste<T>(this DbDataReader rdr, ref List<string> YakalanamayanAlanlar, bool BuyukKucukHarfDuyarli = true)
        {
            var Kayitlar = rdr.Liste(BuyukKucukHarfDuyarli);
            var Sonuc = new List<T>();
            var Propeties = typeof(T).GetProperties();
            foreach (var Kayit in Kayitlar)
            {
                var model = Activator.CreateInstance<T>();
                foreach (var prop in Propeties)
                {
                    if (prop.GetCustomAttributes(typeof(NotMappedAttribute), false).Count() > 0) continue;
                    dynamic val;
                    if (Kayit.TryGetValue(BuyukKucukHarfDuyarli ? prop.Name : prop.Name.ToLower(), out val))
                        prop.SetValue(model, val);
                }
                Sonuc.Add(model);
            }
            if (Kayitlar.Count.Equals(0))
                return Sonuc;
            YakalanamayanAlanlar.AddRange(Kayitlar[0].Keys.ToList().Where(x => Propeties.FirstOrDefault(y => (BuyukKucukHarfDuyarli ? y.Name : y.Name.ToLower()).Equals(x)) == null).ToArray());
            rdr.Close();
            return Sonuc;
        }
        /// <summary>
        /// Veri tabanındaki kayıtları dictionart list şeklinde döndürür
        /// </summary>
        /// <param name="rdr">Veri tabanı Datareader nesnesi</param>
        /// <returns></returns>
        public static List<Dictionary<string, dynamic>> Liste(this DbDataReader rdr, bool BuyukKucukHarfDuyarli = true)
        {
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
                    if (oDeger != null && oDeger != DBNull.Value)
                        Deger = Convert.ChangeType(oDeger, Tip);
                    Kayit.Add(BuyukKucukHarfDuyarli ? Kolon : Kolon.ToLower(), Deger);
                }
                Sonuc.Add(Kayit);
            }
            rdr.Close();
            return Sonuc;
        }
        #endregion

        #region Command
        /// <summary>
        /// Veri tabanındaki kaydı istenen tipe göre döndürür.
        /// </summary>
        /// <typeparam name="T">Geri Dönüş Tipi</typeparam>
        /// <param name="rdr">Veri tabanı Datareader nesnesi</param>
        /// <returns></returns>
        public static T Degisken<T>(this DbCommand com, ref List<string> YakalanamayanAlanlar, bool BuyukKucukHarfDuyarli = true) where T : class
        {
            if (com.Connection.State == ConnectionState.Closed)
                com.Connection.Open();
            var rdr = com.ExecuteReader();
            var Sonuc = rdr.Degisken<T>(ref YakalanamayanAlanlar, BuyukKucukHarfDuyarli);
            rdr.Close();
            com.Connection.Close();
            return Sonuc;
        }
        /// <summary>
        /// Veri tabanındaki kaydı dictionary şeklinde döndürür
        /// </summary>
        /// <param name="rdr">Veri tabanı Datareader nesnesi</param>
        /// <returns></returns>
        public static Dictionary<string, dynamic> Degisken(this DbCommand com, bool BuyukKucukHarfDuyarli = true)
        {
            if (com.Connection.State == ConnectionState.Closed)
                com.Connection.Open();
            var rdr = com.ExecuteReader();
            var Sonuc = rdr.Degisken(BuyukKucukHarfDuyarli);
            rdr.Close();
            com.Connection.Close();
            return Sonuc;
        }
        /// <summary>
        /// Veri tabanındaki kayıtları istenen tipe göre liste şeklinde döndürür
        /// </summary>
        /// <typeparam name="T">Geri Dönüş Tipi</typeparam>
        /// <param name="rdr">Veri tabanı Datareader nesnesi</param>
        /// <returns></returns>
        public static List<T> Liste<T>(this DbCommand com, ref List<string> YakalanamayanAlanlar, bool BuyukKucukHarfDuyarli = true)
        {
            if (com.Connection.State == ConnectionState.Closed)
                com.Connection.Open();
            var rdr = com.ExecuteReader();
            var Sonuc = rdr.Liste<T>(ref YakalanamayanAlanlar, BuyukKucukHarfDuyarli);
            rdr.Close();
            com.Connection.Close();
            return Sonuc;
        }
        /// <summary>
        /// Veri tabanındaki kayıtları dictionary list şeklinde döndürür
        /// </summary>
        /// <param name="rdr">Veri tabanı Datareader nesnesi</param>
        /// <returns></returns>
        public static List<Dictionary<string, dynamic>> Liste(this DbCommand com, bool BuyukKucukHarfDuyarli = true)
        {
            if (com.Connection.State == ConnectionState.Closed)
                com.Connection.Open();
            var rdr = com.ExecuteReader();
            var Sonuc = rdr.Liste(BuyukKucukHarfDuyarli);
            rdr.Close();
            com.Connection.Close();
            return Sonuc;
        }
        #endregion

        #region Connection
        public static Dictionary<string, dynamic> Degisken<T>(this SqlConnection con, T Kayit, bool BuyukKucukHarfDuyarli = true)
        {
            Dictionary<string, dynamic> Sonuc;
            if (con.State.Equals(ConnectionState.Closed))
                con.Open();
            var TabloIsmi = typeof(T).Name.ToString();
            var com = con.CreateCommand();
            com.CommandText = string.Format("select *sutunlar from {0}", TabloIsmi);
            var Propeties = typeof(T).GetProperties();
            var SutunIsimleri = "";


            foreach (var prop in Propeties)
            {
                if (prop.GetCustomAttributes(typeof(NotMappedAttribute), false).Count() > 0) continue;
                SutunIsimleri += "," + prop.Name;
            }
            SutunIsimleri = SutunIsimleri.Substring(1);
            com.CommandText = com.CommandText.Replace("*sutunlar", SutunIsimleri);

            var key = Kayit.GetType().GetProperties().FirstOrDefault(x => Attribute.IsDefined(x, typeof(KeyAttribute)));

            if (key == null)
                Sonuc = com.Degisken(BuyukKucukHarfDuyarli);
            //else if (Convert.ToInt32(key.GetValue(Kayit)) <= 0)
            //    Sonuc = com.Degisken(BuyukKucukHarfDuyarli);
            else
            {
                //Key özellikli bir alan varsa ve değeri 0 dan büyük ise sadece o kaydı cek
                com.CommandText += string.Format(" where {0}=@{0}", key.Name);
                com.Parameters.AddWithValue("@" + key.Name, Convert.ToInt32(key.GetValue(Kayit)));
                Sonuc = com.Degisken(BuyukKucukHarfDuyarli);
            }
            com.Dispose();
            con.Close();
            return Sonuc;
        }

        private static DbType ConvertTypeToDBtype(Type t)
        {
            var typeMap = new Dictionary<Type, DbType>();
            typeMap[typeof(byte)] = DbType.Byte;
            typeMap[typeof(sbyte)] = DbType.SByte;
            typeMap[typeof(short)] = DbType.Int16;
            typeMap[typeof(ushort)] = DbType.UInt16;
            typeMap[typeof(int)] = DbType.Int32;
            typeMap[typeof(uint)] = DbType.UInt32;
            typeMap[typeof(long)] = DbType.Int64;
            typeMap[typeof(ulong)] = DbType.UInt64;
            typeMap[typeof(float)] = DbType.Single;
            typeMap[typeof(double)] = DbType.Double;
            typeMap[typeof(decimal)] = DbType.Decimal;
            typeMap[typeof(bool)] = DbType.Boolean;
            typeMap[typeof(string)] = DbType.String;
            typeMap[typeof(char)] = DbType.StringFixedLength;
            typeMap[typeof(Guid)] = DbType.Guid;
            typeMap[typeof(DateTime)] = DbType.DateTime;
            typeMap[typeof(DateTimeOffset)] = DbType.DateTimeOffset;
            typeMap[typeof(byte[])] = DbType.Binary;
            typeMap[typeof(byte?)] = DbType.Byte;
            typeMap[typeof(sbyte?)] = DbType.SByte;
            typeMap[typeof(short?)] = DbType.Int16;
            typeMap[typeof(ushort?)] = DbType.UInt16;
            typeMap[typeof(int?)] = DbType.Int32;
            typeMap[typeof(uint?)] = DbType.UInt32;
            typeMap[typeof(long?)] = DbType.Int64;
            typeMap[typeof(ulong?)] = DbType.UInt64;
            typeMap[typeof(float?)] = DbType.Single;
            typeMap[typeof(double?)] = DbType.Double;
            typeMap[typeof(decimal?)] = DbType.Decimal;
            typeMap[typeof(bool?)] = DbType.Boolean;
            typeMap[typeof(char?)] = DbType.StringFixedLength;
            typeMap[typeof(Guid?)] = DbType.Guid;
            typeMap[typeof(DateTime?)] = DbType.DateTime;
            typeMap[typeof(DateTimeOffset?)] = DbType.DateTimeOffset;
            //typeMap[typeof(System.Data.Linq.Binary)] = DbType.Binary;
            return typeMap[t];
        }
        public static int Insert<T>(this DbConnection con, T Kayit, bool IdentityInsert = false)
        {
            var TabloIsmi = Kayit.GetType().Name.ToString();
            if (con.State == ConnectionState.Closed)
                con.Open();
            var com = con.CreateCommand();
            com.CommandText = string.Format("insert into {0} (*sutunlar) values (*degerler); select @@identity", TabloIsmi);
            var Propeties = Kayit.GetType().GetProperties();
            var SutunIsimleri = "";
            var ParametreIsimleri = "";
            var Props = Propeties.Where(x => !Attribute.IsDefined(x, typeof(KeyAttribute)) || IdentityInsert);
            foreach (var prop in Props)
            {
                if (prop.GetCustomAttributes(typeof(NotMappedAttribute), false).Count() > 0) continue;
                SutunIsimleri += "," + prop.Name;
                ParametreIsimleri += ",@" + prop.Name;
                var prm = com.CreateParameter();
                prm.Value = Convert.ChangeType(prop.GetValue(Kayit), prop.PropertyType);
                prm.DbType = ConvertTypeToDBtype(prop.PropertyType);
                prm.ParameterName = "@" + prop.Name;
                com.Parameters.Add(prm);
            }
            SutunIsimleri = SutunIsimleri.Substring(1);
            ParametreIsimleri = ParametreIsimleri.Substring(1);
            com.CommandText = com.CommandText.Replace("*sutunlar", SutunIsimleri).Replace("*degerler", ParametreIsimleri);
            var ID = Convert.ToInt32(com.ExecuteScalar());
            com.Dispose();
            con.Close();
            return ID;
        }
        public static string Insert<T>(this DbConnection con, List<T> KayitList, bool IdentityInsert = false)
        {
            var Result = "";
            foreach (var Kayit in KayitList)
            {
                var ID = con.Insert(Kayit, IdentityInsert: IdentityInsert);
                Result += $",{ID}";
            }
            return Result.Substring(1);
        }
        public static int Update<T>(this DbConnection con, T Kayit)
        {
            var TabloIsmi = Kayit.GetType().Name.ToString();
            if (con.State == ConnectionState.Closed)
                con.Open();
            var com = con.CreateCommand();
            com.CommandText = string.Format("update {0} set *upt where *kosul", TabloIsmi);
            var Propeties = Kayit.GetType().GetProperties();


            var KeyProp = Propeties.FirstOrDefault(x => Attribute.IsDefined(x, typeof(KeyAttribute)));
            if (KeyProp == null)
                return 0;

            #region Koşul Kısmı

            var KosulStr = $"{KeyProp.Name}=@{KeyProp.Name}";
            var prm = com.CreateParameter();
            prm.Value = Convert.ChangeType(KeyProp.GetValue(Kayit), KeyProp.PropertyType);
            prm.DbType = ConvertTypeToDBtype(KeyProp.PropertyType);
            prm.ParameterName = "@" + KeyProp.Name;
            com.Parameters.Add(prm);

            #endregion

            #region Guncellenek Alanlar

            var GuncelleStr = "";
            var Props = Propeties.Where(x => !Attribute.IsDefined(x, typeof(KeyAttribute)));
            foreach (var prop in Props)
            {
                if (prop.GetCustomAttributes(typeof(NotMappedAttribute), false).Count() > 0) continue;
                GuncelleStr += $",{prop.Name}=@{prop.Name}";
                prm = com.CreateParameter();
                prm.Value = Convert.ChangeType(prop.GetValue(Kayit), prop.PropertyType);
                prm.DbType = ConvertTypeToDBtype(prop.PropertyType);
                prm.ParameterName = "@" + prop.Name;
                com.Parameters.Add(prm);
            }
            GuncelleStr = GuncelleStr.Substring(1);
            #endregion

            com.CommandText = com.CommandText.Replace("*upt", GuncelleStr).Replace("*kosul", KosulStr);
            var Sonuc = com.ExecuteNonQuery();
            com.Dispose();
            con.Close();
            return Sonuc;
        }
        public static int Delete<T>(this DbConnection con, T Kayit)
        {
            var TabloIsmi = Kayit.GetType().Name.ToString();
            if (con.State == ConnectionState.Closed)
                con.Open();
            var com = con.CreateCommand();
            com.CommandText = string.Format("delete from {0} where *sutun=*deger", TabloIsmi);
            var Propeties = Kayit.GetType().GetProperties();
            var SutunIsimleri = "";
            var ParametreIsimleri = "";
            var prop = Propeties.FirstOrDefault(x => Attribute.IsDefined(x, typeof(KeyAttribute)));
            if (prop == null)
                return 0;

            SutunIsimleri += "" + prop.Name;
            ParametreIsimleri += "@" + prop.Name;
            var prm = com.CreateParameter();
            prm.Value = Convert.ChangeType(prop.GetValue(Kayit), prop.PropertyType);
            prm.DbType = ConvertTypeToDBtype(prop.PropertyType);
            prm.ParameterName = "@" + prop.Name;
            com.Parameters.Add(prm);
            com.CommandText = com.CommandText.Replace("*sutun", SutunIsimleri).Replace("*deger", ParametreIsimleri);
            var Sonuc = com.ExecuteNonQuery();
            com.Dispose();
            con.Close();
            return Sonuc;
        }
        public static List<T> Select<T>(this DbConnection con, ref List<string> YakalanamayanAlanlar, bool BuyukKucukHarfDuyarli = true)
        {
            var TabloIsmi = typeof(T).Name.ToString();
            var com = con.CreateCommand();
            com.CommandText = string.Format("select *sutunlar from {0}", TabloIsmi);
            var Propeties = typeof(T).GetProperties();
            var SutunIsimleri = "";
            foreach (var prop in Propeties)
            {
                if (prop.GetCustomAttributes(typeof(NotMappedAttribute), false).Count() > 0) continue;
                SutunIsimleri += "," + prop.Name;
            }
            SutunIsimleri = SutunIsimleri.Substring(1);
            com.CommandText = com.CommandText.Replace("*sutunlar", SutunIsimleri);
            return com.Liste<T>(ref YakalanamayanAlanlar, BuyukKucukHarfDuyarli);
        }
        public static void Degisken<T>(this SqlConnection con, ref T Kayit, bool BuyukKucukHarfDuyarli = true)
        {
            var Dic = con.Degisken(Kayit, BuyukKucukHarfDuyarli);
            if (Dic == null) return;
            Kayit = Dic.Modelle<T>();
        }
        #endregion

        public static T Modelle<T>(this Dictionary<string, dynamic> dic)
        {
            var Sonuc = Activator.CreateInstance<T>();
            foreach (var prop in Sonuc.GetType().GetProperties())
            {
                if (prop.GetCustomAttributes(typeof(NotMappedAttribute), false).Count() > 0) continue;
                if (dic.ContainsKey(prop.Name) && dic[prop.Name] != null)
                    prop.SetValue(Sonuc, Convert.ChangeType(dic[prop.Name], prop.PropertyType));
            }
            return Sonuc;
        }

        public static void Modelle<T>(this T a, T b)
        {
            foreach (var p in a.GetType().GetProperties())
                p.SetValue(a, p.GetValue(b));
        }
    }
}
