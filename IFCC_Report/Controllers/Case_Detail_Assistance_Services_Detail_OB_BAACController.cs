using GemBox.Spreadsheet;
using IFCC.DAL;
using IFCC.WEB.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;

namespace IFCC_Report.Controllers
{
    public class Case_Detail_Assistance_Services_Detail_OB_BAACController : Controller
    {
        // GET: Case_Detail_Assistance_Services_Detail_OB_BAAC
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Out()
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
                    string a = Url.Action().ToString().ToLower();
                    if(a  == "/case_detail_assistance_services_detail_ob_baac/getbydate/b1")
                    {
                        DataSet ds = DataHelper.GetRequestData(HttpContext);
                        DataRow dr = ds.Tables[0].Rows[0];
                        DataTable dt = CCAReportDAL.Instance.GetCaseDetail(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
                        dt.Columns.Add("Case_ID");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            dt.Rows[i]["Case_ID"] = i + 1;
                        }
                        return DataHelper.GenerateSuccessData(dt);
                    }
                    if(a == "/case_detail_assistance_services_detail_ob_baac/getbydate/b2")
                    {
                        DataSet ds = DataHelper.GetRequestData(HttpContext);
                        DataRow dr = ds.Tables[0].Rows[0];
                        DataTable dt = CCAReportDAL.Instance.GetOutboundDetail(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
                        dt.Columns.Add("Case_ID");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            dt.Rows[i]["Case_ID"] = i + 1;
                        }
                        return DataHelper.GenerateSuccessData(dt);
                    }
                   
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
            if (Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/Book1.xlsx"));
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp1 = workbook.Worksheets[0];
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp2 = workbook.Worksheets[1];
                    GemBox.Spreadsheet.ExcelRow er = null;

                    int i = 0;
                    string sdate = dr["startDate"] + string.Empty.Trim();
                    string edate = dr["endDate"] + string.Empty.Trim();

                    DateTime sd = DateTime.ParseExact(sdate, "yyyyMMdd",
                                  CultureInfo.InvariantCulture);

                    DateTime ed = DateTime.ParseExact(edate, "yyyyMMdd",
                         CultureInfo.InvariantCulture);
                    DataTable dt = new DataTable();
                    dt = CCAReportDAL.Instance.GetCaseDetail(sdate, edate);

                    //Sheet Summary
                    //wsTemp1.Cells["B2"].Value = "Perfomance Summary CCA ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt.Rows.Count ; i++)
                    {
                        try
                        {
                            wsTemp1.Cells["A" + (i + 2)].Value = dt.Rows[i]["CaseID"].ToString().Trim();
                            wsTemp1.Cells["A" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["B" + (i + 2)].Value = dt.Rows[i]["SaveDataTime"].ToString().Trim();
                            wsTemp1.Cells["B" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["C" + (i + 2)].Value = dt.Rows[i]["PIN"].ToString().Trim();
                            wsTemp1.Cells["C" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["D" + (i + 2)].Value = dt.Rows[i]["Memberclass"].ToString().Trim();
                            wsTemp1.Cells["D" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["E" + (i + 2)].Value = dt.Rows[i]["ownername"].ToString().Trim();
                            wsTemp1.Cells["E" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["F" + (i + 2)].Value = dt.Rows[i]["CallName"].ToString().Trim();
                            wsTemp1.Cells["F" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["G" + (i + 2)].Value = dt.Rows[i]["ServiceGroup_Name"].ToString().Trim();
                            wsTemp1.Cells["G" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["H" + (i + 2)].Value = dt.Rows[i]["ServiceType_Name"].ToString().Trim();
                            wsTemp1.Cells["H" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["I" + (i + 2)].Value = dt.Rows[i]["ServiceRequest"].ToString().Trim();
                            wsTemp1.Cells["I" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["J" + (i + 2)].Value = dt.Rows[i]["Solution"].ToString().Trim();
                            wsTemp1.Cells["J" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["K" + (i + 2)].Value = dt.Rows[i]["Agent_ID"].ToString().Trim();
                            wsTemp1.Cells["K" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                        }
                        catch (Exception e)
                        {

                        }
                    }

                    ////Sheet Detail

                    DataTable dt2 = new DataTable();
                    dt2 = CCAReportDAL.Instance.GetOutboundDetail(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);

                    //wsTemp2.Cells["B2"].Value = "Perfomance Detail CCA ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt2.Rows.Count  ; i++)
                    {
                        try
                        {
                            wsTemp2.Cells["A" + (i + 2)].Value = dt2.Rows[i]["CaseID"].ToString().Trim();
                            wsTemp2.Cells["A" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["B" + (i + 2)].Value = dt2.Rows[i]["Datetimes"].ToString().Trim();
                            wsTemp2.Cells["B" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["C" + (i + 2)].Value = dt2.Rows[i]["PIN"].ToString().Trim();
                            wsTemp2.Cells["C" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["D" + (i + 2)].Value = dt2.Rows[i]["Memberclass"].ToString().Trim();
                            wsTemp2.Cells["D" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["E" + (i + 2)].Value = dt2.Rows[i]["ownername"].ToString().Trim();
                            wsTemp2.Cells["E" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["F" + (i + 2)].Value = dt2.Rows[i]["CallName"].ToString().Trim();
                            wsTemp2.Cells["F" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["G" + (i + 2)].Value = dt2.Rows[i]["ServiceGroup_Name"].ToString().Trim();
                            wsTemp2.Cells["G" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["H" + (i + 2)].Value = dt2.Rows[i]["ServiceType_Name"].ToString().Trim();
                            wsTemp2.Cells["H" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["I" + (i + 2)].Value = dt2.Rows[i]["CallContact"].ToString().Trim();
                            wsTemp2.Cells["I" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["J" + (i + 2)].Value = dt2.Rows[i]["Details"].ToString().Trim();
                            wsTemp2.Cells["J" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["K" + (i + 2)].Value = dt2.Rows[i]["Agent_ID"].ToString().Trim();
                            wsTemp2.Cells["K" + (i + 2)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                        

                        }
                        catch (Exception e)
                        {
                            e.Message.ToString();
                        }
                    }

                    string Path = this.CheckPath();
                    string fileName = "Perfomance CCA Report ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty + ".xlsx";
                    workbook.Save(Path + "\\" + fileName);
                    byte[] text = Encoding.UTF8.GetBytes(fileName);

                    //fileName = MachineKey.Encode(text, MachineKeyProtection.All);
                    Console.WriteLine(fileName);
                    return DataHelper.GenerateSuccessData(fileName);
                }
                catch (Exception ex)
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
