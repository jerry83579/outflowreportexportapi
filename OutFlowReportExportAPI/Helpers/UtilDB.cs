using System;
using System.Data;
using System.Linq;
using Dapper;
using System.Data.OleDb;
/// <summary>
/// 資料相關處理
/// 1. DB相關事宜; 2.資料轉成Json格式; 3.資料轉成XML格式; 4.
/// </summary>
public class UtilDB
{
    public UtilDB() { }

    #region DB連線相關處理

    #region DB連線設定
    /// <summary>
    /// 預設FullDBConnction
    /// </summary>
    public static string GetFullDBConnctionString
    {
        get
        {
            return GetFullDBConnction("dbConnectionString", "DataBaseName");
            // getDBConnectionString + ";Database=" + getDBNameString;
        }
    }

    /// <summary>
    /// 預設AccessFullDBConnction
    /// </summary>
    public static string GetFullDBConnctionStringByAccess
    {
        get
        {
            return GetDBConnection("accessConnectionString");// getFullDBConnction("dbConnectionString", "DataBaseName");
            // getDBConnectionString + ";Database=" + getDBNameString;
        }
    }
    /// <summary>
    /// 指定DBConnection
    /// </summary>
    /// <param name="tmpValue"></param>
    /// <returns></returns>
    private static string GetDBConnection(string tmpValue)
    {
        return System.Configuration.ConfigurationManager.ConnectionStrings[tmpValue].ConnectionString;
    }
    /// <summary>
    /// 指定DBName
    /// </summary>
    /// <param name="tmpValue"></param>
    /// <returns></returns>
    private static string GetDBName(string tmpValue)
    {
        return System.Configuration.ConfigurationManager.AppSettings[tmpValue];
        // return System.Configuration.ConfigurationManager.ConnectionStrings[tmpValue].ConnectionString;
    }
    /// <summary>
    /// 指定FullDBConnction
    /// </summary>
    /// <param name="DBConnection"></param>
    /// <param name="DBName"></param>
    /// <returns></returns>
    public static string GetFullDBConnction(string DBConnection, string DBName)
    {
        return GetDBConnection(DBConnection) + ";Database=" + GetDBName(DBName);
    }
    #endregion

    #region MSSQL
    /// <summary>
    /// MSSQL
    /// 進行SQL搜尋功能
    /// 執行SQL且帶入判斷條件內所需要的變數，回傳List的泛型資料
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpSQL">SQL語法</param>
    /// <param name="tmpParameters">判斷條件所需變數</param>
    /// <returns></returns>
    public static System.Collections.Generic.List<T> GetDataList<T>(string tmpSQL)
    {
        return GetDataList<T>(tmpSQL, null);
    }

    /// <summary>
    /// MSSQL
    /// 進行SQL搜尋功能
    /// 執行SQL且帶入判斷條件內所需要的變數，回傳List的泛型資料
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpSQL">SQL語法</param>
    /// <param name="tmpParameters">判斷條件所需變數</param>
    /// <returns></returns>
    public static System.Collections.Generic.List<T> GetDataList<T>(string tmpSQL, object tmpParameters)
    {
        return GetDataList<T>(tmpSQL, tmpParameters, GetFullDBConnctionString);
    }

    /// <summary>
    /// MSSQL
    /// 進行SQL搜尋功能
    /// 執行SQL且帶入判斷條件內所需要的變數，回傳List的泛型資料
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpSQL">SQL語法</param>
    /// <param name="tmpParameters">判斷條件所需變數</param>
    /// <param name="tmpFullDBConnctionString">連線字串</param>
    /// <returns></returns>
    public static System.Collections.Generic.List<T> GetDataList<T>(string tmpSQL, object tmpParameters, string tmpFullDBConnctionString)
    {
        try
        {
            using (var tmpCon = new System.Data.SqlClient.SqlConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Query<T>(tmpSQL, tmpParameters).AsParallel<T>().ToList<T>();
            }
        }
        catch (Exception ex)
        {
            return default(System.Collections.Generic.List<T>);
        }
    }

    /// <summary>
    /// MSSQL
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters"></param>
    /// <returns></returns>
    public static int RunSQLCommand(string tmpSQL, object tmpParameters)
    {
        try
        {
            using (var tmpCon = new System.Data.SqlClient.SqlConnection(GetFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Execute(tmpSQL, tmpParameters);
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }


    public static string RunSQLCommandStr(string tmpSQL, object tmpParameters)
    {
        try
        {
            using (var tmpCon = new System.Data.SqlClient.SqlConnection(GetFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Execute(tmpSQL, tmpParameters).ToString();
            }
        }
        catch (Exception ex)
        {
            return ex.ToString();
        }
    }

    /// <summary>
    /// MSSQL
    /// 同一段語法 帶入不同參數使用
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters">List的格式</param>
    /// <returns></returns>
    public static int RunSQLCommand(string tmpSQL, System.Collections.Generic.List<object> tmpParameters)
    {
        try
        {
            using (var tmpCon = new System.Data.SqlClient.SqlConnection(GetFullDBConnctionString))
            {
                tmpCon.Open();
                int tmpRun = 0;
                foreach (var tmpRow in tmpParameters)
                    if (tmpCon.Execute(tmpSQL, tmpRow) > 0)
                        tmpRun++;
                return tmpRun;
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }

    /// <summary>
    /// MSSQL
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters"></param>
    /// <returns></returns>
    public static int RunSQLCommand(string tmpSQL, object tmpParameters, string tmpFullDBConnctionString)
    {
        try
        {
            using (var tmpCon = new System.Data.SqlClient.SqlConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Execute(tmpSQL, tmpParameters);
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }

    /// <summary>
    /// MSSQL
    /// 同一段語法 帶入不同參數使用
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters">List的格式</param>
    /// <returns></returns>
    public static int RunSQLCommand(string tmpSQL, System.Collections.Generic.List<object> tmpParameters, string tmpFullDBConnctionString)
    {
        try
        {
            using (var tmpCon = new System.Data.SqlClient.SqlConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                int tmpRun = 0;
                foreach (var tmpRow in tmpParameters)
                    if (tmpCon.Execute(tmpSQL, tmpRow) > 0)
                        tmpRun++;
                return tmpRun;
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }
    #endregion

    #region MySQL
    /*
    /// <summary>
    /// MySQL
    /// 進行SQL搜尋功能
    /// 執行SQL且帶入判斷條件內所需要的變數，回傳List的泛型資料
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpSQL">SQL語法</param>
    /// <param name="tmpParameters">判斷條件所需變數</param>
    /// <returns></returns>
    public static System.Collections.Generic.List<T> getDataListByMySQL<T>(string tmpSQL)
    {
        return getDataList<T>(tmpSQL, null);
    }

    /// <summary>
    /// MySQL
    /// 進行SQL搜尋功能
    /// 執行SQL且帶入判斷條件內所需要的變數，回傳List的泛型資料
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpSQL">SQL語法</param>
    /// <param name="tmpParameters">判斷條件所需變數</param>
    /// <returns></returns>
    public static System.Collections.Generic.List<T> getDataListByMySQL<T>(string tmpSQL, object tmpParameters)
    {
        return getDataList<T>(tmpSQL, tmpParameters, getFullDBConnctionString);
    }

    /// <summary>
    /// MySQL
    /// 進行SQL搜尋功能
    /// 執行SQL且帶入判斷條件內所需要的變數，回傳List的泛型資料
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpSQL">SQL語法</param>
    /// <param name="tmpParameters">判斷條件所需變數</param>
    /// <param name="tmpFullDBConnctionString">連線字串</param>
    /// <returns></returns>
    public static System.Collections.Generic.List<T> getDataListByMySQL<T>(string tmpSQL, object tmpParameters, string tmpFullDBConnctionString)
    {
        try
        {

            using (var tmpCon = new MySql.Data.MySqlClient.MySqlConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Query<T>(tmpSQL, tmpParameters).AsParallel<T>().ToList<T>();
            }
        }
        catch (Exception ex)
        {
            return default(System.Collections.Generic.List<T>);
        }
    }

    /// <summary>
    /// MySQL
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters"></param>
    /// <returns></returns>
    public static int RunSQLCommandByMySQL(string tmpSQL, object tmpParameters)
    {
        try
        {
            using (var tmpCon = new MySql.Data.MySqlClient.MySqlConnection(getFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Execute(tmpSQL, tmpParameters);
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }

    /// <summary>
    /// MySQL
    /// 同一段語法 帶入不同參數使用
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters">List的格式</param>
    /// <returns></returns>
    public static int RunSQLCommandByMySQL(string tmpSQL, System.Collections.Generic.List<object> tmpParameters)
    {
        try
        {
            using (var tmpCon = new MySql.Data.MySqlClient.MySqlConnection(getFullDBConnctionString))
            {
                tmpCon.Open();
                int tmpRun = 0;
                (from tmpData2 in tmpParameters.AsParallel()
                 select tmpData2
                ).AsParallel().ToList<object>().ForEach(x => {
                    if (tmpCon.Execute(tmpSQL, x) > 0)
                        tmpRun++;
                });
                return tmpRun;
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }
    
    /// <summary>
    /// MySQL
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters"></param>
    /// <returns></returns>
    public static int RunSQLCommandByMySQL(string tmpSQL, object tmpParameters, string tmpFullDBConnctionString)
    {
        try
        {
            using (var tmpCon = new MySql.Data.MySqlClient.MySqlConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Execute(tmpSQL, tmpParameters);
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }

    /// <summary>
    /// MySQL
    /// 同一段語法 帶入不同參數使用
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters">List的格式</param>
    /// <returns></returns>
    public static int RunSQLCommandByMySQL(string tmpSQL, System.Collections.Generic.List<object> tmpParameters, string tmpFullDBConnctionString)
    {
        try
        {
            using (var tmpCon = new MySql.Data.MySqlClient.MySqlConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                int tmpRun = 0;
                (from tmpData2 in tmpParameters.AsParallel()
                 select tmpData2
                ).AsParallel().ToList<object>().ForEach(x =>
                {
                    if (tmpCon.Execute(tmpSQL, x) > 0)
                        tmpRun++;
                });
                return tmpRun;
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }
    */
    #endregion

    #region Access SQL

    /// <summary>
    /// AccessSQL
    /// 進行SQL搜尋功能
    /// 執行SQL且帶入判斷條件內所需要的變數，回傳List的泛型資料
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpSQL">SQL語法</param>
    /// <param name="tmpParameters">判斷條件所需變數</param>
    /// <returns></returns>
    public static System.Collections.Generic.List<T> getDataListByAccess<T>(string tmpSQL)
    {
        return getDataListByAccess<T>(tmpSQL, null);
    }

    /// <summary>
    /// AccessSQL
    /// 進行SQL搜尋功能
    /// 執行SQL且帶入判斷條件內所需要的變數，回傳List的泛型資料
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpSQL">SQL語法</param>
    /// <param name="tmpParameters">判斷條件所需變數</param>
    /// <returns></returns>
    public static System.Collections.Generic.List<T> getDataListByAccess<T>(string tmpSQL, object tmpParameters)
    {
        return getDataListByAccess<T>(tmpSQL, tmpParameters, GetFullDBConnctionStringByAccess);
    }

    /// <summary>
    /// AccessSQL
    /// 進行SQL搜尋功能
    /// 執行SQL且帶入判斷條件內所需要的變數，回傳List的泛型資料
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpSQL">SQL語法</param>
    /// <param name="tmpParameters">判斷條件所需變數</param>
    /// <param name="tmpFullDBConnctionString">連線字串</param>
    /// <returns></returns>
    public static System.Collections.Generic.List<T> getDataListByAccess<T>(string tmpSQL, object tmpParameters, string tmpFullDBConnctionString)
    {
        try
        {
            using (var tmpCon = new System.Data.OleDb.OleDbConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Query<T>(tmpSQL, tmpParameters).AsParallel<T>().ToList<T>();
            }
        }
        catch (Exception ex)
        {
            return default(System.Collections.Generic.List<T>);
        }
    }

    /// <summary>
    /// AccessSQL
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters"></param>
    /// <returns></returns>
    public static int RunSQLCommandByAccess(string tmpSQL, object tmpParameters)
    {
        try
        {
            using (var tmpCon = new System.Data.OleDb.OleDbConnection(GetFullDBConnctionStringByAccess))
            {
                tmpCon.Open();
                return tmpCon.Execute(tmpSQL, tmpParameters);
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }

    /// <summary>
    /// AccessSQL
    /// 同一段語法 帶入不同參數使用
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters">List的格式</param>
    /// <returns></returns>
    public static int RunSQLCommandByAccess(string tmpSQL, System.Collections.Generic.List<object> tmpParameters)
    {
        try
        {
            using (var tmpCon = new System.Data.OleDb.OleDbConnection(GetFullDBConnctionStringByAccess))
            {
                tmpCon.Open();
                int tmpRun = 0;
                foreach (var tmpRow in tmpParameters)
                    if (tmpCon.Execute(tmpSQL, tmpRow) > 0)
                        tmpRun++;
                return tmpRun;
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }

    /// <summary>
    /// AccessSQL
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters"></param>
    /// <returns></returns>
    public static int RunSQLCommandByAccess(string tmpSQL, object tmpParameters, string tmpFullDBConnctionString)
    {
        try
        {
            using (var tmpCon = new System.Data.OleDb.OleDbConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Execute(tmpSQL, tmpParameters, null, 300, null);
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }
    /// <summary>
    /// AccessSQL
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters"></param>
    /// <returns></returns>
    public static int RunSQLCommandByAccess(string tmpSQL, string tmpFullDBConnctionString)
    {
        try
        {
            using (var tmpCon = new System.Data.OleDb.OleDbConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                return tmpCon.Execute(tmpSQL);
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }


    /// <summary>
    /// AccessSQL
    /// 同一段語法 帶入不同參數使用
    /// 執行Insert、Update、Delete語法
    /// </summary>
    /// <param name="tmpSQL"></param>
    /// <param name="tmpParameters">List的格式</param>
    /// <returns></returns>
    public static int RunSQLCommandByAccess(string tmpSQL, System.Collections.Generic.List<object> tmpParameters, string tmpFullDBConnctionString)
    {
        try
        {
            using (var tmpCon = new System.Data.OleDb.OleDbConnection(tmpFullDBConnctionString))
            {
                tmpCon.Open();
                int tmpRun = 0;
                foreach (var tmpRow in tmpParameters)
                    if (tmpCon.Execute(tmpSQL, tmpRow) > 0)
                        tmpRun++;
                return tmpRun;
            }
        }
        catch (Exception ex)
        {
            return 0;
        }
    }

    #endregion

    #endregion

    #region Json相關處理內容

    /// <summary>
    /// T轉Json資料(包含DataTable也行)，如有意外發生，則回傳空字串
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpData"></param>
    /// <returns></returns>
    public static string getObjectToJson<T>(T tmpData)
    {
        try
        {
            if (tmpData == null)
                return "{}";
            return Newtonsoft.Json.JsonConvert.SerializeObject(tmpData, Newtonsoft.Json.Formatting.None);
        }
        catch (Exception ex)
        {
            return "";
        }
    }
    /// <summary>
    /// 將Json資料轉成物件，如有意外發生，則回傳default(T)
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpData"></param>
    /// <returns></returns>
    public static T getJsonToObject<T>(string tmpData)
    {
        try
        {
            return (T)Newtonsoft.Json.JsonConvert.DeserializeObject(tmpData);
        }
        catch (Exception ex)
        {
            return default(T);
        }
    }

    #endregion

    #region XML相關處理內容
    /// <summary>
    /// Class 轉 XML
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpObject"></param>
    /// <param name="E"></param>
    /// <returns></returns>
    public static string getObjectToXML<T>(T tmpObject, System.Text.Encoding E)
    {
        using (System.IO.MemoryStream tmpMemory = new System.IO.MemoryStream())
        using (System.IO.StreamWriter tmpWriter = new System.IO.StreamWriter(tmpMemory, E))
        {
            System.Xml.Serialization.XmlSerializer tmpSer = new System.Xml.Serialization.XmlSerializer(tmpObject.GetType());
            tmpSer.Serialize(tmpWriter, tmpObject);
            return E.GetString(tmpMemory.ToArray());
        }
    }

    /// <summary>
    /// XML 轉 Class
    /// </summary>
    /// <typeparam name="T">泛型</typeparam>
    /// <param name="tmpValue"></param>
    /// <param name="E"></param>
    /// <returns></returns>
    public static T getXMLToObject<T>(string tmpValue, System.Text.Encoding E)
    {
        System.Xml.XmlDocument tmpXmlDocument = new System.Xml.XmlDocument();
        try
        {
            byte[] tmpEncodeByte = E.GetBytes(tmpValue);
            using (System.IO.MemoryStream tmpMemory = new System.IO.MemoryStream(tmpEncodeByte))
            {
                tmpMemory.Flush();
                tmpMemory.Position = 0;
                tmpXmlDocument.Load(tmpMemory);
            }
            using (System.Xml.XmlNodeReader tmpReader = new System.Xml.XmlNodeReader(tmpXmlDocument.DocumentElement))
            {
                System.Xml.Serialization.XmlSerializer tmpSer = new System.Xml.Serialization.XmlSerializer(typeof(T));
                object tmpObject = tmpSer.Deserialize(tmpReader);
                return (T)tmpObject;
            }
        }
        catch (Exception ex)
        {
            return default(T);
        }
    }
    #endregion

}

//public static class DataTableExtensions
//{
//    /// <summary>
//    /// 將DataTable 轉換成 List<dynamic>
//    /// reverse 反轉：控制返回结果中是只存在 FilterField 指定的字段,還是排除.
//    /// [flase 返回FilterField 指定的字段]|[true 返回结果剔除 FilterField 指定的字段]
//    /// FilterField  字段過濾，FilterField 為空 忽略 reverse 參數；返回DataTable中的全部數
//    /// </summary>
//    /// <param name="table">DataTable</param>
//    /// <param name="reverse">
//    /// 反轉：控制返回结果中是只存在 FilterField 指定的字段,還是排除.
//    /// [flase 返回FilterField 指定的字段]|[true 返回结果剔除 FilterField 指定的字段]
//    ///</param>
//    /// <param name="FilterField">字段過濾，FilterField 為空 忽略 reverse 參數；返回DataTable中的全部數據</param>
//    /// <returns>List<dynamic></returns>
//    public static List<dynamic> ToDynamicList(this DataTable table, bool reverse = true, params string[] FilterField)
//    {
//        var modelList = new List<dynamic>();
//        foreach (DataRow row in table.Rows)
//        {
//            dynamic model = new ExpandoObject();
//            var dict = (IDictionary<string, object>)model;
//            foreach (DataColumn column in table.Columns)
//            {
//                if (FilterField.Length != 0)
//                {
//                    if (reverse == true)
//                    {
//                        if (!FilterField.Contains(column.ColumnName))
//                        {
//                            dict[column.ColumnName] = row[column];
//                        }
//                    }
//                    else
//                    {
//                        if (FilterField.Contains(column.ColumnName))
//                        {
//                            dict[column.ColumnName] = row[column];
//                        }
//                    }
//                }
//                else
//                {
//                    dict[column.ColumnName] = row[column];
//                }
//            }
//            modelList.Add(model);
//        }
//        return modelList;
//    }
//}
