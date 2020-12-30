using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Helpers;
using System.Web.Script.Serialization;

namespace IFCC.WEB.Services
{
    public static class DataHelper
    {

        private static DataSet _ds = new DataSet();

        #region GetRequestData
        public static DataSet GetRequestData(HttpContextBase HttpContext)
        {
            var resolveRequest = HttpContext.Request;
            Dictionary<string, object> model = new Dictionary<string, object>();
            resolveRequest.InputStream.Seek(0, SeekOrigin.Begin);
            string jsonString = new StreamReader(resolveRequest.InputStream).ReadToEnd();
            if (jsonString != null)
            {
                jsonString = Encoding.UTF8.GetString(Convert.FromBase64String(jsonString));
                int dataCRC = 0;
                int dataLength = jsonString.Length;
                if (!string.IsNullOrEmpty(resolveRequest.Headers["Context-Data-CRC"]))
                {
                    dataCRC = Convert.ToInt32(resolveRequest.Headers["Context-Data-CRC"]);
                }
                if (dataCRC != 0 && dataLength != 0 && dataCRC != dataLength)
                {
                    throw new Exception("Requested data is invalid.<br>Please check again !");
                }

                JavaScriptSerializer serializer = new JavaScriptSerializer();
                model = (Dictionary<string, object>)serializer.Deserialize(jsonString, typeof(Dictionary<string, object>));
            }
            _ds = ConvertDictionaryToDataSet(model);
            return _ds.Copy();
        }
        #endregion

        #region GetUserLoginID
        public static string GetUserLoginID(HttpContextBase HttpContext)
        {
            string value = string.Empty;
            DataSet ds = GetRequestData(HttpContext);
            if (ds != null && ds.Tables.Count > 0)
            {
                if (ds.Tables["user"] != null && ds.Tables["user"].Rows.Count > 0)
                {
                    if (ds.Tables["user"].Rows[0]["userLoginID"] != null)
                    {
                        value = ds.Tables["user"].Rows[0]["userLoginID"].ToString();
                    }
                }
            }
            return value;
        } 
        #endregion

        #region GetRequestValue
        public static string GetRequestValue(HttpContextBase HttpContext, string columnName)
        {
            string value = string.Empty;
            DataSet ds = GetRequestData(HttpContext);
            if (ds != null && ds.Tables.Count > 0)
            {
                if (ds.Tables["master"] != null && ds.Tables["master"].Rows.Count > 0)
                {
                    if (ds.Tables["master"].Rows[0][columnName] != null)
                    {
                        value = ds.Tables["master"].Rows[0][columnName].ToString();
                    }
                }
                else if (ds.Tables["master"] == null && ds.Tables[0] != null && ds.Tables[0].Rows.Count > 0)
                {
                    if (ds.Tables[0].Rows[0][columnName] != null)
                    {
                        value = ds.Tables[0].Rows[0][columnName].ToString();
                    }
                }
            }
            return value;
        }

        public static string GetRequestValue(DataSet ds, string columnName)
        {
            string value = string.Empty;
            if (ds != null && ds.Tables.Count > 0)
            {
                if (ds.Tables["master"] != null && ds.Tables["master"].Rows.Count > 0)
                {
                    if (ds.Tables["master"].Rows[0][columnName] != null)
                    {
                        value = ds.Tables["master"].Rows[0][columnName].ToString();
                    }
                }
            }
            return value;
        }

        public static string GetRequestValue(string columnName)
        {
            string value = string.Empty;
            if (_ds != null && _ds.Tables.Count > 0)
            {
                if (_ds.Tables["master"] != null && _ds.Tables["master"].Rows.Count > 0)
                {
                    if (_ds.Tables["master"].Rows[0][columnName] != null)
                    {
                        value = _ds.Tables["master"].Rows[0][columnName].ToString();
                    }
                }
            }
            return value;
        }
        #endregion

        #region ResetRequestData
        public static void ResetRequestData()
        {
            _ds.Clear();
            _ds = new DataSet("DS");
        }
        #endregion

        #region SetDataSetValue
        public static DataSet SetDataSetValue(DataSet ds, string tableName, string columnName, string value)
        {
            if (ds != null && ds.Tables.Count > 0)
            {
                if (ds.Tables[tableName] != null && ds.Tables[tableName].Rows.Count > 0)
                {
                    foreach(DataRow dr in ds.Tables[tableName].Rows)
                    {
                        if (dr[columnName] != null)
                        {
                            dr[columnName] = value;
                        }
                    }
                    
                }
            }
            return ds;
        }
        #endregion

        #region ConvertDictionaryToDataSet
        public static DataSet ConvertDictionaryToDataSet(Dictionary<string, object> model)
        {
            DataSet ds = new DataSet("DS");
            DataTable dt = new DataTable();
            //dt.Locale = new CultureInfo("th-TH");
            Type type;
            DataRow dr;
            foreach (var item in model)
            {
                type = (item.Value != null) ? item.Value.GetType() : null;
                if (type == typeof(Dictionary<string, object>))
                {
                    dt = new DataTable(item.Key.ToUpper());
                    //dt.Locale = new CultureInfo("th-TH");
                    foreach (var item2 in (Dictionary<string, object>)item.Value)
                    {
                        type = (item2.Value != null) ? item2.Value.GetType() : null;
                        if (type != typeof(Dictionary<string, object>))
                        {
                            dt.Columns.Add(item2.Key);
                        }
                    }
                    dr = dt.NewRow();
                    foreach (var item2 in (Dictionary<string, object>)item.Value)
                    {
                        type = (item2.Value != null) ? item2.Value.GetType() : null;
                        if (type != typeof(Dictionary<string, object>))
                        {
                            dr[item2.Key] = item2.Value;
                        }
                    }
                    dt.Rows.Add(dr);
                    ds.Tables.Add(dt);
                }
                else if (type == typeof(ArrayList))
                {
                    dt = new DataTable(item.Key.ToUpper());
                    //dt.Locale = new CultureInfo("th-TH");
                    foreach (var item2 in (ArrayList)item.Value)
                    {
                        foreach (var item3 in (Dictionary<string, object>)item2)
                        {
                            if (!dt.Columns.Contains(item3.Key))
                            {
                                dt.Columns.Add(item3.Key);
                            }
                        }
                        dr = dt.NewRow();
                        foreach (var item3 in (Dictionary<string, object>)item2)
                        {
                            type = (item3.Value != null) ? item3.Value.GetType() : null;
                            if (type != typeof(Dictionary<string, object>))
                            {
                                dr[item3.Key] = item3.Value;
                            }
                        }
                        dt.Rows.Add(dr);
                    }
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }
        #endregion

        #region GenerateSuccessData
        public static string GenerateSuccessData(object dataResult)
        {
            try
            {
                Dictionary<string, object> dicSuccess = new Dictionary<string, object>();
                object status = 0;
                if (dataResult.GetType() == typeof(Dictionary<string, object>))
                {
                    Dictionary<string, object> dicResult = (Dictionary<string, object>)dataResult;
                    status = dicResult["status"];
                }
                dicSuccess.Add("status", status);
                dicSuccess.Add("result", dataResult);

                var sr = new StringWriter();
                var jsonWriter = new JsonTextWriter(sr)
                {
                    StringEscapeHandling = StringEscapeHandling.EscapeHtml
                };
                new JsonSerializer().Serialize(jsonWriter, dicSuccess);
                return sr.ToString();

                //return JsonConvert.SerializeObject(dicSuccess);
            }
            catch (Exception ex)
            {
                return GenerateErrorData(ex);
            }
        }

        public static string GenerateSuccessData_PDF(object dataResult)
        {
            try
            {
                Dictionary<string, object> dicSuccess = new Dictionary<string, object>();
                object status = 0;
                object status2 = 0;
                if (dataResult.GetType() == typeof(Dictionary<string, object>))
                {
                    Dictionary<string, object> dicResult = (Dictionary<string, object>)dataResult;
                    status = dicResult["status"];
                    status2 = dicResult["status2"];
                }
                dicSuccess.Add("status", status);
                dicSuccess.Add("status2", status2);
                dicSuccess.Add("result", dataResult);

                var sr = new StringWriter();
                var jsonWriter = new JsonTextWriter(sr)
                {
                    StringEscapeHandling = StringEscapeHandling.EscapeHtml
                };
                new JsonSerializer().Serialize(jsonWriter, dicSuccess);
                return sr.ToString();

                //return JsonConvert.SerializeObject(dicSuccess);
            }
            catch (Exception ex)
            {
                return GenerateErrorData(ex);
            }
        }

        public static string GenerateSuccessData1(object dataResult, object dataResult2)
        {
            try
            {
                Dictionary<string, object> dicSuccess = new Dictionary<string, object>();
                object status = 0;
                if (dataResult.GetType() == typeof(Dictionary<string, object>))
                {
                    Dictionary<string, object> dicResult = (Dictionary<string, object>)dataResult;
                    status = dicResult["status"];
                }
                dicSuccess.Add("status", status);
                dicSuccess.Add("result", dataResult);

                object status2 = 0;
                if (dataResult2.GetType() == typeof(Dictionary<string, object>))
                {
                    Dictionary<string, object> dicResult = (Dictionary<string, object>)dataResult2;
                    status2 = dicResult["status"];
                }
                dicSuccess.Add("status2", status2);
                dicSuccess.Add("result2", dataResult2);

                var sr = new StringWriter();
                var jsonWriter = new JsonTextWriter(sr)
                {
                    StringEscapeHandling = StringEscapeHandling.EscapeHtml
                };
                new JsonSerializer().Serialize(jsonWriter, dicSuccess);
                return sr.ToString();

                //return JsonConvert.SerializeObject(dicSuccess);
            }
            catch (Exception ex)
            {
                return GenerateErrorData(ex);
            }
        }

        public static string GenerateSuccessData2(object dataResult, object dataResult2,object dataRequest_date)
        {
            try
            {
                Dictionary<string, object> dicSuccess = new Dictionary<string, object>();
                object status = 0;
                if (dataResult.GetType() == typeof(Dictionary<string, object>))
                {
                    Dictionary<string, object> dicResult = (Dictionary<string, object>)dataResult;
                    status = dicResult["status"];
                }
                dicSuccess.Add("status", status);
                dicSuccess.Add("result", dataResult);

                object status2 = 0;
                if (dataResult2.GetType() == typeof(Dictionary<string, object>))
                {
                    Dictionary<string, object> dicResult = (Dictionary<string, object>)dataResult2;
                    status2 = dicResult["status"];
                }
                dicSuccess.Add("status2", status2);
                dicSuccess.Add("result2", dataResult2);

                object status3 = 0;
                if (dataRequest_date.GetType() == typeof(Dictionary<string, object>))
                {
                    Dictionary<string, object> dicResult = (Dictionary<string, object>)dataRequest_date;
                    status3 = dicResult["status"];
                }
                dicSuccess.Add("status3", status3);
                dicSuccess.Add("result3", dataRequest_date);

                var sr = new StringWriter();
                var jsonWriter = new JsonTextWriter(sr)
                {
                    StringEscapeHandling = StringEscapeHandling.EscapeHtml
                };
                new JsonSerializer().Serialize(jsonWriter, dicSuccess);
                return sr.ToString();

                //return JsonConvert.SerializeObject(dicSuccess);
            }
            catch (Exception ex)
            {
                return GenerateErrorData(ex);
            }
        }
        #endregion

        #region GenerateErrorData
        public static string GenerateErrorData(Exception ex)
        {
            Dictionary<string, object> dicError = new Dictionary<string, object>();
            dicError.Add("status", -1);
            dicError.Add("message", ex.Message);
            dicError.Add("message_stack_trace", ex.StackTrace);

            var sr = new StringWriter();
            var jsonWriter = new JsonTextWriter(sr)
            {
                StringEscapeHandling = StringEscapeHandling.EscapeHtml
            };
            new JsonSerializer().Serialize(jsonWriter, dicError);
            return sr.ToString();

            //return JsonConvert.SerializeObject(dicError); ;
        }
        #endregion

        #region GenerateReturnData
        public static string GenerateResponseData(HttpResponseBase res)
        {
            Dictionary<string, object> dicResponse = new Dictionary<string, object>();
            dicResponse.Add("status", -1);
            dicResponse.Add("status_code", res.StatusCode);
            dicResponse.Add("message", res.StatusDescription);

            var sr = new StringWriter();
            var jsonWriter = new JsonTextWriter(sr)
            {
                StringEscapeHandling = StringEscapeHandling.EscapeHtml
            };
            new JsonSerializer().Serialize(jsonWriter, dicResponse);
            return sr.ToString();

            //return JsonConvert.SerializeObject(dicError); ;
        }
        #endregion

        #region GenerateGUID
        public static DataSet GenerateNewGUID(DataSet ds, string tableName, string columnName)
        {
            if (ds != null && ds.Tables.Count > 0)
            {
                if (ds.Tables[tableName] != null && ds.Tables[tableName].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[tableName].Rows)
                    {
                        if (dr[columnName] == null)
                        {
                            dr[columnName] = Guid.NewGuid();
                        }
                    }
                }
            }
            return ds;
        }
        public static DataSet GenerateGUID(DataSet ds, string tableName, string columnName, string checkColumnName)
        {
            return GenerateGUID(ds, tableName, columnName, checkColumnName, "New");
        }
        public static DataSet GenerateGUID(DataSet ds, string tableName, string columnName, string checkColumnName, string checkValueWith)
        {
            if (ds != null && ds.Tables.Count > 0)
            {
                if (ds.Tables[tableName] != null && ds.Tables[tableName].Rows.Count > 0)
                {
                    foreach (DataRow dr in ds.Tables[tableName].Rows)
                    {
                        if (dr[columnName] != null && dr[checkColumnName] != null)
                        {
                            if((dr[checkColumnName] +"").ToLower() == checkValueWith.ToLower())
                            {
                                dr[columnName] = Guid.NewGuid();
                            }
                        }
                    }
                }
            }
            return ds;
        }
        #endregion
    }
}