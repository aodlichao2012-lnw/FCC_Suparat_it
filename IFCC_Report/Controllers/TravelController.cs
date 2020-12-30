using GemBox.Spreadsheet;
using IFCC.DAL;
using IFCC.WEB.Services;
using NPOI.SS.Util;
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
    public class TravelController : Controller
    {
        // GET: Hotel
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


                    DataTable dt = BAACReportDAL.Instance.GetTravalReport(dr["startDate"] + string.Empty, dr["endDate"] + string.Empty);
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
            if (Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/TravelAndFilghtReport-Template.xlsx"));
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp = workbook.Worksheets[0];
                    GemBox.Spreadsheet.ExcelRow er = null;


                    int i = 0;
                    string sdate = dr["startDate"] + string.Empty.Trim();
                    string edate = dr["endDate"] + string.Empty.Trim();

                    DateTime sd = DateTime.ParseExact(sdate, "yyyyMMdd",
                                  CultureInfo.InvariantCulture);

                    DateTime ed = DateTime.ParseExact(edate, "yyyyMMdd",
                         CultureInfo.InvariantCulture);

                    DataTable dt = new DataTable();
                    dt = BAACReportDAL.Instance.GetTravalReport(sdate, edate);

                    wsTemp.Cells["B2"].Value = "ตารางการแลกสิทธิ์หมวดการเดินทางและการขนส่ง ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH"));

                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp.Cells["B" + (i + 6)].Value = dt.Rows[i]["Case_Id"].ToString().Trim();
                            wsTemp.Cells["B" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["C" + (i + 6)].Value = dt.Rows[i]["ServiceGroup_Name"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 6)].Value = dt.Rows[i]["ServiceType_Name"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["E" + (i + 6)].Value = dt.Rows[i]["Create_Date"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 6)].Value = dt.Rows[i]["Owner_Name"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 6)].Value = dt.Rows[i]["Owner_Phone"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["H" + (i + 6)].Value = dt.Rows[i]["Book_Name"].ToString().Trim();
                            wsTemp.Cells["H" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["I" + (i + 6)].Value = dt.Rows[i]["Book_Phone"].ToString().Trim();
                            wsTemp.Cells["I" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["J" + (i + 6)].Value = dt.Rows[i]["Book_Date"].ToString().Trim();
                            wsTemp.Cells["J" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["K" + (i + 6)].Value = dt.Rows[i]["Book_Time"].ToString().Trim();
                            wsTemp.Cells["K" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["L" + (i + 6)].Value = dt.Rows[i]["Flight"].ToString().Trim();
                            wsTemp.Cells["L" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["M" + (i + 6)].Value = dt.Rows[i]["Flight_Date"].ToString().Trim();
                            wsTemp.Cells["M" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["N" + (i + 6)].Value = dt.Rows[i]["Flight_Time"].ToString().Trim();
                            wsTemp.Cells["N" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["O" + (i + 6)].Value = dt.Rows[i]["สถานที่ไปรับ"].ToString().Trim();
                            wsTemp.Cells["O" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["P" + (i + 6)].Value = dt.Rows[i]["สถานที่ไปส่ง"].ToString().Trim();
                            wsTemp.Cells["P" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["Q" + (i + 6)].Value = dt.Rows[i]["Address"].ToString().Trim();
                            wsTemp.Cells["Q" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["R" + (i + 6)].Value = dt.Rows[i]["Book_Status"].ToString().Trim();
                            wsTemp.Cells["R" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["S" + (i + 6)].Value = dt.Rows[i]["DeliveryType"].ToString().Trim();
                            wsTemp.Cells["S" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["T" + (i + 6)].Value = dt.Rows[i]["Car_PlateNo"].ToString().Trim();
                            wsTemp.Cells["T" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["U" + (i + 6)].Value = dt.Rows[i]["Driver"].ToString().Trim();
                            wsTemp.Cells["U" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["V" + (i + 6)].Value = dt.Rows[i]["Create_By"].ToString().Trim();
                            wsTemp.Cells["V" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);


                            wsTemp.Cells["W" + (i + 6)].Value = dt.Rows[i]["น้องหอมจัง"].ToString().Trim();
                            wsTemp.Cells["W" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }
                        catch (Exception e)
                        {

                        }

                    }



                    string Path = this.CheckPath();
                    string fileName = "ตารางการแลกสิทธิ์หมวดการเดินทางและการขนส่ง ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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