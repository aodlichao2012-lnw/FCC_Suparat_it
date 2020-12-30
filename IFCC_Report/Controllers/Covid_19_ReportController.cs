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
    public class Covid_19_ReportController : Controller
    {
        // GET: Covid_19_Report
        public ActionResult Index()
        {
            return View();
        }

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
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/TMB_COVIT-19_Template.xlsx"));
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
                    dt = BlacktieDAL.Instance.GetCovid_19(sdate, edate);

                    
                    wsTemp.Cells["B2"].Value = "Data From " + sd.ToString("dd MMMMM yyyy", new CultureInfo("en-US")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("en-US"));
                    wsTemp.Cells["B2"].Style.Font.Size = 600;
                    wsTemp.Cells["B2"].Style.Font.Color = SpreadsheetColor.FromName(ColorName.Purple);
                    wsTemp.Cells["B2"].Style.Font.Weight= ExcelFont.BoldWeight;

                    int count = dt.Rows.Count + 1;
                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp.Cells["B" + (i + 5)].Value = dt.Rows[i]["No"].ToString().Trim();
                            wsTemp.Cells["B" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["C" + (i + 5)].Value = dt.Rows[i]["Reference_ID"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 5)].Value = dt.Rows[i]["Redempton_Date"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["E" + (i + 5)].Value = dt.Rows[i]["Account_Owner_Name"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 5)].Value = dt.Rows[i]["Privilege_Voucher_Number"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 5)].Value = dt.Rows[i]["Policy_number"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["H" + (i + 5)].Value = dt.Rows[i]["Policy_Effective_Date"].ToString().Trim();
                            wsTemp.Cells["H" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["I" + (i + 5)].Value = dt.Rows[i]["Policy_Owner_Name"].ToString().Trim();
                            wsTemp.Cells["I" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["J" + (i + 5)].Value = dt.Rows[i]["Donate_Amount"].ToString().Trim();
                            wsTemp.Cells["J" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);


                        }
                        catch (Exception e)
                        {

                        }

                    }
                   
                    string Path = this.CheckPath();
                    string fileName = "Covid - 19  Report  ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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

        #region ExportGrab_Food
        [HttpPost]
        public object ExportGrab_Food()
        {
            if (Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/GrabFood_Report_Template.xlsx"));
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
                    dt = BlacktieDAL.Instance.GetGrabFood(sdate, edate);


                    wsTemp.Cells["B2"].Value = "Data From " + sd.ToString("dd MMMMM yyyy", new CultureInfo("en-US")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("en-US"));
                    wsTemp.Cells["B2"].Style.Font.Size = 600;
                    wsTemp.Cells["B2"].Style.Font.Color = SpreadsheetColor.FromName(ColorName.Green);
                    wsTemp.Cells["B2"].Style.Font.Weight = ExcelFont.BoldWeight;

                    int count = dt.Rows.Count + 1;
                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp.Cells["B" + (i + 5)].Value = dt.Rows[i]["NO"].ToString().Trim();
                            wsTemp.Cells["B" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["C" + (i + 5)].Value = dt.Rows[i]["FL03"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 5)].Value = dt.Rows[i]["create_date"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["E" + (i + 5)].Value = dt.Rows[i]["Guest_Name"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 5)].Value = dt.Rows[i]["Guest_mobile"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 5)].Value = dt.Rows[i]["redeem_code"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["H" + (i + 5)].Value = dt.Rows[i]["privilege_code"].ToString().Trim();
                            wsTemp.Cells["H" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }
                        catch (Exception e)
                        {

                        }

                    }

                    string Path = this.CheckPath();
                    string fileName = "Grab Food   Report  ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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

        #region ExportHealth_Care_Plus
        [HttpPost]
        public object ExportHealth_Care_Plus()
        {
            if (Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/HealthCarePlus_Report_Template.xlsx"));
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
                    dt = BlacktieDAL.Instance.GetHealthCarePlus(sdate, edate);

                    wsTemp.Cells["B2"].Value = "Data From " + sd.ToString("dd MMMMM yyyy", new CultureInfo("en-US")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("en-US"));
                    wsTemp.Cells["B2"].Style.Font.Size = 600;
                    wsTemp.Cells["B2"].Style.Font.Color = SpreadsheetColor.FromName(ColorName.Orange);
                    wsTemp.Cells["B2"].Style.Font.Weight = ExcelFont.BoldWeight;


                    int count = dt.Rows.Count + 1;
                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp.Cells["B" + (i + 5)].Value = dt.Rows[i]["NO"].ToString().Trim();
                            wsTemp.Cells["B" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["C" + (i + 5)].Value = dt.Rows[i]["Reference_ID"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 5)].Value = dt.Rows[i]["Redemption_Date"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["E" + (i + 5)].Value = dt.Rows[i]["Account_Owner_Name"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 5)].Value = dt.Rows[i]["Privilege_Voucher_Number"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 5)].Value = dt.Rows[i]["Policy_Number"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["H" + (i + 5)].Value = dt.Rows[i]["Policy_Effective_Date"].ToString().Trim();
                            wsTemp.Cells["H" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["I" + (i + 5)].Value = dt.Rows[i]["Policy_Owner_Name"].ToString().Trim();
                            wsTemp.Cells["I" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["J" + (i + 5)].Value = dt.Rows[i]["Donate_Amount"].ToString().Trim();
                            wsTemp.Cells["J" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }
                        catch (Exception e)
                        {

                        }

                    }

                    string Path = this.CheckPath();
                    string fileName = "Health Care Plus Report  ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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

        #region ExportHotel
        [HttpPost]
        public object ExportHotel()
        {
            if (Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/HotelBooking_Report_Templete.xlsx"));
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
                    dt = BlacktieDAL.Instance.GetHotel(sdate, edate);


                    wsTemp.Cells["B2"].Value = "Data From " + sd.ToString("dd MMMMM yyyy", new CultureInfo("en-US")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("en-US"));
                    wsTemp.Cells["B2"].Style.Font.Size = 600;
                    wsTemp.Cells["B2"].Style.Font.Color = SpreadsheetColor.FromName(ColorName.Blue);
                    wsTemp.Cells["B2"].Style.Font.Weight = ExcelFont.BoldWeight;

                    int count = dt.Rows.Count + 1;
                    int re = 0;
                    int re2 = 0;
                    int re3 = 0;
                    int re4 = 0;
                    int total = 0;
                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp.Cells["B" + (i + 5)].Value = dt.Rows[i]["NO"].ToString().Trim();
                            wsTemp.Cells["B" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["C" + (i + 5)].Value = dt.Rows[i]["privilege_code"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 5)].Value = dt.Rows[i]["create_date"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["E" + (i + 5)].Value = dt.Rows[i]["create_time"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 5)].Value = dt.Rows[i]["Account_name"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 5)].Value = dt.Rows[i]["Mobile"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["H" + (i + 5)].Value = dt.Rows[i]["Guest_Name"].ToString().Trim();
                            wsTemp.Cells["H" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["I" + (i + 5)].Value = dt.Rows[i]["Guest_Mobile"].ToString().Trim();
                            wsTemp.Cells["I" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["J" + (i + 5)].Value = dt.Rows[i]["Email"].ToString().Trim();
                            wsTemp.Cells["J" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["K" + (i + 5)].Value = dt.Rows[i]["Remineder_Date"].ToString().Trim();
                            wsTemp.Cells["K" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["L" + (i + 5)].Value = dt.Rows[i]["Check_Out_Date"].ToString().Trim();
                            wsTemp.Cells["L" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["M" + (i + 5)].Value = dt.Rows[i]["Location"].ToString().Trim();
                            wsTemp.Cells["M" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["N" + (i + 5)].Value = dt.Rows[i]["Location_Detail"].ToString().Trim();
                            wsTemp.Cells["N" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                            
                            wsTemp.Cells["O" + (i + 5)].Value = dt.Rows[i]["Passenger_Adult_Count"].ToString().Trim();
                            wsTemp.Cells["O" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["P" + (i + 5)].Value = dt.Rows[i]["Passenger_Child_Count"].ToString().Trim();
                            wsTemp.Cells["P" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["Q" + (i + 5)].Value = dt.Rows[i]["Remark"].ToString().Trim();
                            wsTemp.Cells["Q" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["R" + (i + 5)].Value = dt.Rows[i]["Redeem_Code"].ToString().Trim();
                            wsTemp.Cells["R" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["S" + (i + 5)].Value = dt.Rows[i]["Redeem_Code_2"].ToString().Trim();
                            wsTemp.Cells["S" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["T" + (i + 5)].Value = dt.Rows[i]["Redeem_Code_3"].ToString().Trim();
                            wsTemp.Cells["T" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["U" + (i + 5)].Value = dt.Rows[i]["Redeem_Code_4"].ToString().Trim();
                            wsTemp.Cells["U" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            if (!string.IsNullOrEmpty(dt.Rows[i]["Redeem_Code"].ToString()))
                            {
                                re = 1;
                            }
                            else { re = 0; }

                            if (!string.IsNullOrEmpty(dt.Rows[i]["Redeem_Code_2"].ToString()))
                            {
                                re2 = 1;
                            }
                            else { re2 = 0; }

                            if (!string.IsNullOrEmpty(dt.Rows[i]["Redeem_Code_3"].ToString()))
                            {
                                re3 = 1;
                            }
                            else { re3 = 0; }
                            if (!string.IsNullOrEmpty(dt.Rows[i]["Redeem_Code_3"].ToString()))
                            {
                                re4 = 1;
                            }
                            else { re4 = 0; }

                            total = re + re2 + re3 + re4;

                            wsTemp.Cells["V" + (i + 5)].Value = total;
                            wsTemp.Cells["V" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }
                        catch (Exception e)
                        {

                        }

                    }

                    string Path = this.CheckPath();
                    string fileName = "Hotel Booking Report ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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

        #region ExportRed_Cross
        [HttpPost]
        public object ExportRed_Cross()
        {
            if (Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/RedCross_Report_Templete.xlsx"));
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
                    dt = BlacktieDAL.Instance.GetRedCross(sdate, edate);

                    wsTemp.Cells["B2"].Value = "Data From " + sd.ToString("dd MMMMM yyyy", new CultureInfo("en-US")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("en-US"));
                    wsTemp.Cells["B2"].Style.Font.Size = 600;
                    wsTemp.Cells["B2"].Style.Font.Color = SpreadsheetColor.FromName(ColorName.DarkRed);
                    wsTemp.Cells["B2"].Style.Font.Weight = ExcelFont.BoldWeight;

                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp.Cells["B" + (i + 5)].Value = dt.Rows[i]["No"].ToString().Trim();
                            wsTemp.Cells["B" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["C" + (i + 5)].Value = dt.Rows[i]["create_date"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 5)].Value = dt.Rows[i]["Guest_Name"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["E" + (i + 5)].Value = dt.Rows[i]["Donator_address"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 5)].Value = dt.Rows[i]["Donate_Amount"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 5)].Value = dt.Rows[i]["สิทธิ์ลดหย่อน"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }
                        catch (Exception e)
                        {

                        }

                    }

                    string Path = this.CheckPath();
                    string fileName = "Red Cross Report ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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

        #region Export_S_and_P
        [HttpPost]
        public object Export_S_and_P()
        {
            if (Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/S_and_P_Report_Template.xlsx"));
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
                    dt = BlacktieDAL.Instance.GetS_and_P(sdate, edate);


                    wsTemp.Cells["B2"].Value = "Data From " + sd.ToString("dd MMMMM yyyy", new CultureInfo("en-US")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("en-US"));
                    wsTemp.Cells["B2"].Style.Font.Size = 600;
                    wsTemp.Cells["B2"].Style.Font.Color = SpreadsheetColor.FromName(ColorName.Accent2Darker25Pct);
                    wsTemp.Cells["B2"].Style.Font.Weight = ExcelFont.BoldWeight;

                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp.Cells["B" + (i + 5)].Value = dt.Rows[i]["NO"].ToString().Trim();
                            wsTemp.Cells["B" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["C" + (i + 5)].Value = dt.Rows[i]["FL03"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 5)].Value = dt.Rows[i]["create_date"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["E" + (i + 5)].Value = dt.Rows[i]["Guest_Name"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 5)].Value = dt.Rows[i]["Guest_Mobile"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 5)].Value = dt.Rows[i]["redeem_code"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["H" + (i + 5)].Value = dt.Rows[i]["privilege_code"].ToString().Trim();
                            wsTemp.Cells["H" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                        } 

                        catch (Exception e)
                        {

                        }

                    }

                    string Path = this.CheckPath();
                    string fileName = "S and P Report ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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

        #region ExportVaccine
        [HttpPost]
        public object ExportVaccine()
        {
            if (Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/Vaccine_Report_Templete.xlsx"));
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
                    
                    dt = BlacktieDAL.Instance.GetVaccine(sdate, edate);
                    dt.Columns.Add("NO");

                    wsTemp.Cells["B2"].Value = "Data From " + sd.ToString("dd MMMMM yyyy", new CultureInfo("en-US")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("en-US"));
                    wsTemp.Cells["B2"].Style.Font.Size = 600;
                    wsTemp.Cells["B2"].Style.Font.Color = SpreadsheetColor.FromName(ColorName.Orange);
                    wsTemp.Cells["B2"].Style.Font.Weight = ExcelFont.BoldWeight;

                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            dt.Rows[i]["NO"] = i + 1;

                            wsTemp.Cells["B" + (i + 5)].Value = dt.Rows[i]["NO"].ToString().Trim();
                            wsTemp.Cells["B" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["C" + (i + 5)].Value = dt.Rows[i]["create_date"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 5)].Value = dt.Rows[i]["create_time"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["E" + (i + 5)].Value = dt.Rows[i]["Account_Name"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 5)].Value = dt.Rows[i]["เบอร์เจ้าของสิทธิ์"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 5)].Value = dt.Rows[i]["Guest_Name"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["H" + (i + 5)].Value = dt.Rows[i]["Guest_Birthday"].ToString().Trim();
                            wsTemp.Cells["H" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["I" + (i + 5)].Value = dt.Rows[i]["เบอร์ผู้รับVaccine"].ToString().Trim();
                            wsTemp.Cells["I" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["J" + (i + 5)].Value = dt.Rows[i]["Remineder_Date"].ToString().Trim();
                            wsTemp.Cells["J" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["K" + (i + 5)].Value = dt.Rows[i]["Remineder_Time"].ToString().Trim();
                            wsTemp.Cells["K" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["L" + (i + 5)].Value = dt.Rows[i]["Location"].ToString().Trim();
                            wsTemp.Cells["L" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["M" + (i + 5)].Value = dt.Rows[i]["Location_Detail"].ToString().Trim();
                            wsTemp.Cells["M" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["N" + (i + 5)].Value = dt.Rows[i]["Redeem_Code"].ToString().Trim();
                            wsTemp.Cells["N" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["O" + (i + 5)].Value = dt.Rows[i]["Number"].ToString().Trim();
                            wsTemp.Cells["O" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                        }

                        catch (Exception e)
                        {

                        }

                    }

                    string Path = this.CheckPath();
                    string fileName = "Vaccine Report ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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

        #region ExportWine
        [HttpPost]
        public object ExportWine()
        {
            if (Response.StatusCode == 200)
            {
                try
                {
                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];

                    SpreadsheetInfo.SetLicense("ETZX-IT28-33Q6-1HA2");
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/WineReport_Template.xlsx"));
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

                    dt = BlacktieDAL.Instance.GetWine(sdate, edate);

                    wsTemp.Cells["B2"].Value = "Data From " + sd.ToString("dd MMMMM yyyy", new CultureInfo("en-US")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("en-US"));
                    wsTemp.Cells["B2"].Style.Font.Size = 600;
                    wsTemp.Cells["B2"].Style.Font.Color = SpreadsheetColor.FromName(ColorName.Background2Darker50Pct);
                    wsTemp.Cells["B2"].Style.Font.Weight = ExcelFont.BoldWeight;

                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                    

                            wsTemp.Cells["B" + (i + 5)].Value = dt.Rows[i]["NO"].ToString().Trim();
                            wsTemp.Cells["B" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["C" + (i + 5)].Value = dt.Rows[i]["FL03"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 5)].Value = dt.Rows[i]["create_date"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["E" + (i + 5)].Value = dt.Rows[i]["Guest_Name"].ToString().Trim();
                            wsTemp.Cells["E" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["F" + (i + 5)].Value = dt.Rows[i]["Guest_mobile"].ToString().Trim();
                            wsTemp.Cells["F" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["G" + (i + 5)].Value = dt.Rows[i]["redeem_code"].ToString().Trim();
                            wsTemp.Cells["G" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["H" + (i + 5)].Value = dt.Rows[i]["privilege_code"].ToString().Trim();
                            wsTemp.Cells["H" + (i + 5)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                        }

                        catch (Exception e)
                        {

                        }

                    }

                    string Path = this.CheckPath();
                    string fileName = "Wine Report ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + " - " + ed.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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