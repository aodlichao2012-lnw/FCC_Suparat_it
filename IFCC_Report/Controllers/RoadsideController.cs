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
using System.Web.UI;

namespace IFCC_Report.Controllers
{
    public class RoadsideController : Controller
    {
   
        public ActionResult BAAC()
        {
            return View();
        }

        public ActionResult Tbank()
        {
            return View();
        }

        public ActionResult MTL()
        {
            return View();
        }

        public ActionResult KMT()
        {
            return View();
        }
        public ActionResult Cigna()
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
                    if (a == "/roadside/getbydate/baac")
                    {
                        DataSet ds = DataHelper.GetRequestData(HttpContext);
                        DataRow dr = ds.Tables[0].Rows[0];
                        DataTable dt = BAACReportDAL.Instance.GetRoadsideReport(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
                        dt.Columns.Add("Case_ID");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            dt.Rows[i]["Case_ID"] = i + 1;
                        }
                        return DataHelper.GenerateSuccessData(dt);
                    }
                    else if (a == "/roadside/getbydate/tbank")
                    {
                        DataSet ds = DataHelper.GetRequestData(HttpContext);
                        DataRow dr = ds.Tables[0].Rows[0];
                        DataTable dt = BAACReportDAL.Instance.GetTbankReportt(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
                        dt.Columns.Add("Case_ID");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            dt.Rows[i]["Case_ID"] = i + 1;
                        }
                        return DataHelper.GenerateSuccessData(dt);
                    }   
                    else if (a == "/roadside/getbydate/mtl")
                    {
                        DataSet ds = DataHelper.GetRequestData(HttpContext);
                        DataRow dr = ds.Tables[0].Rows[0];
                        DataTable dt = BAACReportDAL.Instance.GetMTLReport(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
                        dt.Columns.Add("Case_ID");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            dt.Rows[i]["Case_ID"] = i + 1;
                        }
                        return DataHelper.GenerateSuccessData(dt);
                    }   
                    else if (a == "/roadside/getbydate/kmt")
                    {
                        DataSet ds = DataHelper.GetRequestData(HttpContext);
                        DataRow dr = ds.Tables[0].Rows[0];
                        DataTable dt = BAACReportDAL.Instance.GetKMTReport(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
                        dt.Columns.Add("Case_ID");
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {

                            dt.Rows[i]["Case_ID"] = i + 1;
                        }
                        return DataHelper.GenerateSuccessData(dt);
                    }   
                    else if (a == "/roadside/getbydate/cigna")
                    {
                        DataSet ds = DataHelper.GetRequestData(HttpContext);
                        DataRow dr = ds.Tables[0].Rows[0];
                        DataTable dt = BAACReportDAL.Instance.GetCignaReport(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
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
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/Rosider.xlsx"));
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp1 = workbook.Worksheets[0];
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp2 = workbook.Worksheets[1];
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp3 = workbook.Worksheets[2];
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp4 = workbook.Worksheets[3];
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp5 = workbook.Worksheets[4];
                    //GemBox.Spreadsheet.ExcelWorksheet wsTemp6 = workbook.Worksheets[5];
                    GemBox.Spreadsheet.ExcelRow er = null;

                    int i = 0;
                    string sdate = dr["startDate"] + string.Empty.Trim();
                    string edate = dr["endDate"] + string.Empty.Trim();

                    DateTime sd = DateTime.ParseExact(sdate, "yyyyMMdd",
                                  CultureInfo.InvariantCulture);

                    DateTime ed = DateTime.ParseExact(edate, "yyyyMMdd",
                         CultureInfo.InvariantCulture);
                    DataTable dt = new DataTable();
                    dt = BAACReportDAL.Instance.GetRoadsideReport(sdate, edate);

                    //Sheet Summary
                    //wsTemp1.Cells["B2"].Value = "Perfomance Summary CCA ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp1.Cells["B" + (i + 3)].Value = dt.Rows[i]["date"].ToString().Trim();
                            wsTemp1.Cells["B" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["C" + (i + 3)].Value = dt.Rows[i]["ชื่อสกุล"].ToString().Trim();
                            wsTemp1.Cells["C" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["D" + (i + 3)].Value = dt.Rows[i]["Pin"].ToString().Trim();
                            wsTemp1.Cells["D" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["E" + (i + 3)].Value = dt.Rows[i]["MemberClass"].ToString().Trim();
                            wsTemp1.Cells["E" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["F" + (i + 3)].Value = dt.Rows[i]["ทะเบียนรถ"].ToString().Trim();
                            wsTemp1.Cells["F" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["G" + (i + 3)].Value = dt.Rows[i]["ยี่ห้อรถ"].ToString().Trim();
                            wsTemp1.Cells["G" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["H" + (i + 3)].Value = dt.Rows[i]["อายุรถ"].ToString().Trim();
                            wsTemp1.Cells["H" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["I" + (i + 3)].Value = dt.Rows[i]["case_Id"].ToString().Trim();
                            wsTemp1.Cells["I" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["J" + (i + 3)].Value = dt.Rows[i]["ServiceGroup_Name"].ToString().Trim();
                            wsTemp1.Cells["J" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["K" + (i + 3)].Value = dt.Rows[i]["ServiceType_Name"].ToString().Trim();
                            wsTemp1.Cells["K" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["L" + (i + 3)].Value = dt.Rows[i]["ผู้ให้บริการ"].ToString().Trim();
                            wsTemp1.Cells["L" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["M" + (i + 3)].Value = dt.Rows[i]["actual_service_amount"].ToString().Trim();
                            wsTemp1.Cells["M" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["N" + (i + 3)].Value = dt.Rows[i]["Create_by"].ToString().Trim();
                            wsTemp1.Cells["N" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["O" + (i + 3)].Value = dt.Rows[i]["incident_remark"].ToString().Trim();
                            wsTemp1.Cells["O" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }
                        catch (Exception e)
                        {

                        }
                    }

                    ////Sheet Detail

                    DataTable dt2 = new DataTable();
                    dt2 = BAACReportDAL.Instance.GetTbankReportt(sdate , edate);

                    //wsTemp2.Cells["B2"].Value = "Perfomance Detail CCA ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt2.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp2.Cells["B" + (i + 3)].Value = dt2.Rows[i]["date"].ToString().Trim();
                            wsTemp2.Cells["B" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["C" + (i + 3)].Value = dt2.Rows[i]["C_TBANK_FULL_NAME"].ToString().Trim();
                            wsTemp2.Cells["C" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["D" + (i + 3)].Value = dt2.Rows[i]["CUST_ID"].ToString().Trim();
                            wsTemp2.Cells["D" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["E" + (i + 3)].Value = dt2.Rows[i]["C_TBANK_LEVEL_CARD"].ToString().Trim();
                            wsTemp2.Cells["E" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["F" + (i + 3)].Value = dt2.Rows[i]["ทะเบียนรถ"].ToString().Trim();
                            wsTemp2.Cells["F" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["G" + (i + 3)].Value = dt2.Rows[i]["ยี่ห้อรถ"].ToString().Trim();
                            wsTemp2.Cells["G" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["H" + (i + 3)].Value = dt2.Rows[i]["อายุรถ"].ToString().Trim();
                            wsTemp2.Cells["H" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["I" + (i + 3)].Value = dt2.Rows[i]["case_Id"].ToString().Trim();
                            wsTemp2.Cells["I" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["J" + (i + 3)].Value = dt2.Rows[i]["Scope_Name"].ToString().Trim();
                            wsTemp2.Cells["J" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["K" + (i + 3)].Value = dt2.Rows[i]["Service_Name"].ToString().Trim();
                            wsTemp2.Cells["K" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["L" + (i + 3)].Value = dt2.Rows[i]["ผู้ให้บริการ"].ToString().Trim();
                            wsTemp2.Cells["L" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["M" + (i + 3)].Value = dt2.Rows[i]["ค่าบริการ"].ToString().Trim();
                            wsTemp2.Cells["M" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["N" + (i + 3)].Value = dt2.Rows[i]["agent_name"].ToString().Trim();
                            wsTemp2.Cells["N" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["O" + (i + 3)].Value = dt2.Rows[i]["etc"].ToString().Trim();
                            wsTemp2.Cells["O" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);



                        }
                        catch (Exception e)
                        {
                            e.Message.ToString();
                        }
                    }

                //
                    DataTable dt3 = new DataTable();
                    dt3 = BAACReportDAL.Instance.GetMTLReport(sdate , edate);

                    //wsTemp2.Cells["B2"].Value = "Perfomance Detail CCA ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt3.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp3.Cells["B" + (i + 3)].Value = dt3.Rows[i]["date"].ToString().Trim();
                            wsTemp3.Cells["B" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["C" + (i + 3)].Value = dt3.Rows[i]["name"].ToString().Trim();
                            wsTemp3.Cells["C" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["D" + (i + 3)].Value = dt3.Rows[i]["Client_ID"].ToString().Trim();
                            wsTemp3.Cells["D" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["E" + (i + 3)].Value = dt3.Rows[i]["Cust_Class"].ToString().Trim();
                            wsTemp3.Cells["E" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["F" + (i + 3)].Value = dt3.Rows[i]["ทะเบียนรถ"].ToString().Trim();
                            wsTemp3.Cells["F" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["G" + (i + 3)].Value = dt3.Rows[i]["ยี่ห้อรถ"].ToString().Trim();
                            wsTemp3.Cells["G" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["H" + (i + 3)].Value = dt3.Rows[i]["อายุรถ"].ToString().Trim();
                            wsTemp3.Cells["H" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["I" + (i + 3)].Value = dt3.Rows[i]["Case_No"].ToString().Trim();
                            wsTemp3.Cells["I" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["J" + (i + 3)].Value = dt3.Rows[i]["Scope_Service"].ToString().Trim();
                            wsTemp3.Cells["J" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["K" + (i + 3)].Value = dt3.Rows[i]["service_type"].ToString().Trim();
                            wsTemp3.Cells["K" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["L" + (i + 3)].Value = dt3.Rows[i]["ผู้ให้บริการ"].ToString().Trim();
                            wsTemp3.Cells["L" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["M" + (i + 3)].Value = dt3.Rows[i]["ค่าบริการ"].ToString().Trim();
                            wsTemp3.Cells["M" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["N" + (i + 3)].Value = dt3.Rows[i]["Agent_Name"].ToString().Trim();
                            wsTemp3.Cells["N" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp3.Cells["O" + (i + 3)].Value = dt3.Rows[i]["Mark"].ToString().Trim();
                            wsTemp3.Cells["O" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);



                        }
                        catch (Exception e)
                        {
                            e.Message.ToString();
                        }
                    }
                    DataTable dt4 = new DataTable();
                    dt4 = BAACReportDAL.Instance.GetKMTReport(sdate , edate);

                    //wsTemp2.Cells["B2"].Value = "Perfomance Detail CCA ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt4.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp4.Cells["B" + (i + 3)].Value = dt4.Rows[i]["date"].ToString().Trim();
                            wsTemp4.Cells["B" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["C" + (i + 3)].Value = dt4.Rows[i]["name"].ToString().Trim();
                            wsTemp4.Cells["C" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["D" + (i + 3)].Value = dt4.Rows[i]["Cust_Id"].ToString().Trim();
                            wsTemp4.Cells["D" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["E" + (i + 3)].Value = dt4.Rows[i]["Cust_Group"].ToString().Trim();
                            wsTemp4.Cells["E" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["F" + (i + 3)].Value = dt4.Rows[i]["ทะเบียนรถ"].ToString().Trim();
                            wsTemp4.Cells["F" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["G" + (i + 3)].Value = dt4.Rows[i]["Make"].ToString().Trim();
                            wsTemp4.Cells["G" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["H" + (i + 3)].Value = dt4.Rows[i]["อายุรถ"].ToString().Trim();
                            wsTemp4.Cells["H" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["I" + (i + 3)].Value = dt4.Rows[i]["case_Id"].ToString().Trim();
                            wsTemp4.Cells["I" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["J" + (i + 3)].Value = dt4.Rows[i]["Service_Group"].ToString().Trim();
                            wsTemp4.Cells["J" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["K" + (i + 3)].Value = dt4.Rows[i]["service_type"].ToString().Trim();
                            wsTemp4.Cells["K" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["L" + (i + 3)].Value = dt4.Rows[i]["ผู้ให้บริการ"].ToString().Trim();
                            wsTemp4.Cells["L" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["M" + (i + 3)].Value = dt4.Rows[i]["Mechanical_Cost"].ToString().Trim();
                            wsTemp4.Cells["M" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["N" + (i + 3)].Value = dt4.Rows[i]["Agent_Name"].ToString().Trim();
                            wsTemp4.Cells["N" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp4.Cells["O" + (i + 3)].Value = dt4.Rows[i]["Mark"].ToString().Trim();
                            wsTemp4.Cells["O" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                        }
                        catch (Exception e)
                        {
                            e.Message.ToString();
                        }
                    }
                    DataTable dt5 = new DataTable();
                    dt5 = BAACReportDAL.Instance.GetCignaReport(sdate , edate);

                    //wsTemp2.Cells["B2"].Value = "Perfomance Detail CCA ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt5.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp5.Cells["B" + (i + 3)].Value = dt5.Rows[i]["date"].ToString().Trim();
                            wsTemp5.Cells["B" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["C" + (i + 3)].Value = dt5.Rows[i]["name"].ToString().Trim();
                            wsTemp5.Cells["C" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["D" + (i + 3)].Value = dt5.Rows[i]["Update_Group"].ToString().Trim();
                            wsTemp5.Cells["D" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            //wsTemp5.Cells["E" + (i + 3)].Value = dt5.Rows[i]["MemberClass"].ToString().Trim();
                            //wsTemp5.Cells["E" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["E" + (i + 3)].Value = dt5.Rows[i]["ทะเบียนรถ"].ToString().Trim();
                            wsTemp5.Cells["E" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["F" + (i + 3)].Value = dt5.Rows[i]["ยี่ห้อรถ"].ToString().Trim();
                            wsTemp5.Cells["F" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["G" + (i + 3)].Value = dt5.Rows[i]["อายุรถ"].ToString().Trim();
                            wsTemp5.Cells["G" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["H" + (i + 3)].Value = dt5.Rows[i]["case_Id"].ToString().Trim();
                            wsTemp5.Cells["H" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["I" + (i + 3)].Value = dt5.Rows[i]["Scope_Service"].ToString().Trim();
                            wsTemp5.Cells["I" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["J" + (i + 3)].Value = dt5.Rows[i]["Service_Type"].ToString().Trim();
                            wsTemp5.Cells["J" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["K" + (i + 3)].Value = dt5.Rows[i]["Provd_Name"].ToString().Trim();
                            wsTemp5.Cells["K" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["L" + (i + 3)].Value = dt5.Rows[i]["Service_Cost"].ToString().Trim();
                            wsTemp5.Cells["L" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["M" + (i + 3)].Value = dt5.Rows[i]["Agent_Name"].ToString().Trim();
                            wsTemp5.Cells["M" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp5.Cells["N" + (i + 3)].Value = dt5.Rows[i]["Mark"].ToString().Trim();
                            wsTemp5.Cells["N" + (i + 3)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

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
