  using IFCC.WEB.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using IFCC.DAL;
using GemBox.Spreadsheet;
using System.Web.Security;
using System.Text;
using System.IO;

namespace IFCC_Report.Controllers
{
    public class StudentController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }


        #region GetByDate
        [HttpPost]
        public object GetByDate()
        {
           
            if (Response.StatusCode == 200)
            {
                try
                {
                   
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];
                    string project = null;
                    if (string.IsNullOrEmpty(dr["projectName"] + string.Empty))
                    {
                         project = "%";
                    }
                    else
                    {
                        project = dr["projectName"] + string.Empty;
                    }
                    DataTable dt = CallbackDAL.Instance.GetCallBack(project, dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
                    return DataHelper.GenerateSuccessData(dt);
                }
                catch (Exception ex)
                { }
                return null;
            }
            else
            {
                return DataHelper.GenerateResponseData(Response);
            }
        }
        #endregion

        #region GetUpdateStatus
        [HttpPost]
        public object GetUpdateStatus()
        {

            if (Response.StatusCode == 200)
            {
                try
                {

                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[1].Rows[0];
                    string id = (dr["ID"] + string.Empty).Trim();
                    

                    DataTable result = CallbackDAL.Instance.GetUpdateStatus(id, "1");
                    return DataHelper.GenerateSuccessData(result);
                    
                }
                catch (Exception ex)
                { }
                return null;
            }
            else
            {
                return DataHelper.GenerateResponseData(Response);
            }
        }
        #endregion

        #region Export
        [HttpPost]
        public object Export() 
        {
        if(Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/CallBack_Template.xlsx"));
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp = workbook.Worksheets[0];
                    GemBox.Spreadsheet.ExcelRow er = null;


                    int i = 0;
                    string project = null;
                    if (string.IsNullOrEmpty(dr["projectName"] + string.Empty))
                    {
                        project = "%";
                    }
                    else
                    {
                        project = dr["projectName"] + string.Empty;
                    }


                    DataTable dt = new DataTable();
                    dt = CallbackDAL.Instance.GetCallBack(project, dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);

                    dt.Columns.Add("Call_Text",typeof(string));

                    for (i = 0;i < dt.Rows.Count; i++){
                        if (dt.Rows[i]["Export"].ToString() == "False" || dt.Rows[i]["Export"].ToString() == "false")
                        {
                            dt.Rows[i]["Call_Text"] = "ยังไม่ได้โทร";
                        }
                        else if (dt.Rows[i]["Export"].ToString() == "True" || dt.Rows[i]["Export"].ToString() == "true")
                        {
                            dt.Rows[i]["Call_Text"] = "โทรแล้ว";
                        }
                        
                    
                    }
                    wsTemp.Cells["E2"].Value  = "CallBack Report ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp.Cells["E" + (i + 5)].Value = dt.Rows[i]["Callback_Id"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 5)].Value = dt.Rows[i]["Project_Name"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 5)].Value = dt.Rows[i]["CDN"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["H" + (i + 5)].Value = dt.Rows[i]["Telephone_No"].ToString().Trim();
                            wsTemp.Cells["H" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["I" + (i + 5)].Value = dt.Rows[i]["Callback_Date"].ToString().Trim();
                            wsTemp.Cells["I" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["J" + (i + 5)].Value = dt.Rows[i]["Callback_Time"].ToString().Trim();
                            wsTemp.Cells["J" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["K" + (i + 5)].Value = dt.Rows[i]["Callback_Type"].ToString().Trim();
                            wsTemp.Cells["K" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["L" + (i + 5)].Value = dt.Rows[i]["Export"].ToString().Trim();
                            wsTemp.Cells["L" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["M" + (i + 5)].Value = dt.Rows[i]["Call_Text"].ToString().Trim();
                            wsTemp.Cells["M" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }
                        catch (Exception e)
                        {

                        }

                    }



                    string Path = this.CheckPath();
                    string fileName = "CallBack Report ประจำช่วงวันที่ " + dr["startDate"] + string.Empty +  " - " + dr["endDate"] + string.Empty + ".xlsx";
                    workbook.Save(Path + "\\" + fileName);
                    byte[] text = Encoding.UTF8.GetBytes(fileName);

                    //fileName = MachineKey.Encode(text, MachineKeyProtection.All);
                    Console.WriteLine(fileName);
                    return DataHelper.GenerateSuccessData(fileName);
                }
                catch(Exception ex) 
                {
                
                }
            }
            return DataHelper.GenerateSuccessData(null);
        }
        #endregion


        #region CheckPath
        private string CheckPath()
        {
            try
            {
                string PathExport = Server.MapPath("~/Export");
                if (!Directory.Exists(PathExport))
                    Directory.CreateDirectory(PathExport);

                return PathExport;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        #endregion

        #region DownloadFile

        public ActionResult DownloadFile(string FileName)
        {
            //byte[] plaintextBytes = MachineKey.Decode(FileName, MachineKeyProtection.All);
            //FileName = Encoding.UTF8.GetString(plaintextBytes);
            string fullPath = Path.Combine(Server.MapPath("~/Export"), FileName);
            return File(fullPath, "application/vnd.ms-excel", FileName);
        }
        #endregion

    }
}
