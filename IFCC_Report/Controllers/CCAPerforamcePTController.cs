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
    public class CCAPerforamcePTController : Controller
    {
        // GET: CCAPerforamcePT
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
                    DataTable dt = CCAReportDAL.Instance.GetDetailPerformanceReport( dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
                    dt.Columns.Add("Case_Id");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        dt.Rows[i]["Case_Id"] = i + 1;
                    }
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
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/Performance_CCA_Template.xlsx"));
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
                    dt = CCAReportDAL.Instance.GetSummaryPerformanceReport(sdate, edate);

                    //Sheet Summary
                    wsTemp1.Cells["B2"].Value = "Perfomance Summary CCA ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp1.Cells["B" + (i + 5)].Value = dt.Rows[i]["SaveDateTime"].ToString().Trim();
                            wsTemp1.Cells["B" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["C" + (i + 5)].Value = dt.Rows[i]["Agent"].ToString().Trim();
                            wsTemp1.Cells["C" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["D" + (i + 5)].Value = dt.Rows[i]["Inbound_Petrol"].ToString().Trim();
                            wsTemp1.Cells["D" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["E" + (i + 5)].Value = dt.Rows[i]["Inbound_Roadside"].ToString().Trim();
                            wsTemp1.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["F" + (i + 5)].Value = dt.Rows[i]["OB Co-ordinate With PT Station"].ToString().Trim();
                            wsTemp1.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["G" + (i + 5)].Value = dt.Rows[i]["OB Co-ordinate With Provider"].ToString().Trim();
                            wsTemp1.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["H" + (i + 5)].Value = dt.Rows[i]["Out_ประสานงานลูกค้า"].ToString().Trim();
                            wsTemp1.Cells["H" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["I" + (i + 5)].Value = dt.Rows[i]["Information & Other"].ToString().Trim();
                            wsTemp1.Cells["I" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["J" + (i + 5)].Value = dt.Rows[i]["Out_ประสานงานGSC"].ToString().Trim();
                            wsTemp1.Cells["J" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp1.Cells["K" + (i + 5)].Value = dt.Rows[i]["Total"].ToString().Trim();
                            wsTemp1.Cells["K" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }
                        catch (Exception e)
                        {

                        }
                    }

                    ////Sheet Detail

                    DataTable dt2 = new DataTable();
                    dt2 = CCAReportDAL.Instance.GetDetailPerformanceReport(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);

                    wsTemp2.Cells["B2"].Value = "Perfomance Detail CCA ประจำช่วงวันที่ " + dr["startDate"] + string.Empty + " - " + dr["endDate"] + string.Empty;

                    for (i = 0; i < dt2.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp2.Cells["B" + (i + 6)].Value = dt2.Rows[i]["SaveDateTime"].ToString().Trim();
                            wsTemp2.Cells["B" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["C" + (i + 6)].Value = dt2.Rows[i]["DetailID"].ToString().Trim();
                            wsTemp2.Cells["C" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["D" + (i + 6)].Value = dt2.Rows[i]["Inbound_Petrol"].ToString().Trim();
                            wsTemp2.Cells["D" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["E" + (i + 6)].Value = dt2.Rows[i]["Inbound_Roadside"].ToString().Trim();
                            wsTemp2.Cells["E" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["F" + (i + 6)].Value = dt2.Rows[i]["Inbound_Infor"].ToString().Trim();
                            wsTemp2.Cells["F" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["G" + (i + 6)].Value = dt2.Rows[i]["Inbound_Other"].ToString().Trim();
                            wsTemp2.Cells["G" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["H" + (i + 6)].Value = dt2.Rows[i]["ประสานงานปั๊มน้ำมัน"].ToString().Trim();
                            wsTemp2.Cells["H" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["I" + (i + 6)].Value = dt2.Rows[i]["Out_ประสานงานPT"].ToString().Trim();
                            wsTemp2.Cells["I" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["J" + (i + 6)].Value = dt2.Rows[i]["Out_Provider"].ToString().Trim();
                            wsTemp2.Cells["J" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["K" + (i + 6)].Value = dt2.Rows[i]["Out_ประสานงานGSC"].ToString().Trim();
                            wsTemp2.Cells["K" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["L" + (i + 6)].Value = dt2.Rows[i]["ประสานงานการทางพิเศษ"].ToString().Trim();
                            wsTemp2.Cells["L" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp2.Cells["M" + (i + 6)].Value = dt2.Rows[i]["Out_ประสานงานลูกค้า"].ToString().Trim();
                            wsTemp2.Cells["M" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            //if (dt2.Rows[i]["Agent"].ToString().Trim() == null)
                            //{
                            //    dt2.Rows[i]["Agent"] = "ไม่สามารถระบบ Agent ได้";
                            //}
                            wsTemp2.Cells["N" + (i + 6)].Value = dt2.Rows[i]["Agent"].ToString().Trim();
                            wsTemp2.Cells["N" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }
                        catch (Exception e)
                        {
                            e.Message.ToString();
                        }
                    }

                    string Path = this.CheckPath();
                    string fileName = "Perfomance CCA Report ประจำช่วงวันที่ " + dr["startDate"] + string.Empty +  " - " + dr["endDate"] + string.Empty + ".xlsx";
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