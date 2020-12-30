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
    public class KmotorReportController : Controller
    {
        // GET: KmotorReport
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


                    DataTable dt = KmotorDAL.Instance.GetKmotorReport(dr["startDate"] + string.Empty);
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
                    ExcelFile workbook = ExcelFile.Load(Server.MapPath("~/Template/KmotorReport_Template.xlsx"));
                    GemBox.Spreadsheet.ExcelWorksheet wsTemp = workbook.Worksheets[0];
                    GemBox.Spreadsheet.ExcelRow er = null;


                    int i = 0;
                    string sdate = dr["startDate"] + string.Empty.Trim();
                   

                    DateTime sd = DateTime.ParseExact(sdate, "yyyyMMdd",
                                  CultureInfo.InvariantCulture);

                    DataTable dt = new DataTable();
                    dt = KmotorDAL.Instance.GetKmotorReport(sdate);

                    int motor = 0;
                    wsTemp.Cells["B3"].Value = sd.ToString("dd MMMMM yyyy", new CultureInfo("en-US"));

                    int count = dt.Rows.Count + 1;
                    for (i = 0; i < dt.Rows.Count; i++)
                    {
                        try
                        {
                            wsTemp.Cells["C" + (i + 6)].Value = dt.Rows[i]["Agent_Name"].ToString().Trim();
                            wsTemp.Cells["C" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                            wsTemp.Cells["D" + (i + 6)].Value = dt.Rows[i]["Kmotor"].ToString().Trim();
                            wsTemp.Cells["D" + (i + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                             motor += Convert.ToInt32(dt.Rows[i]["Kmotor"]);

                        }
                        catch (Exception e)
                        {

                        }

                    }
                    int row = dt.Rows.Count;
                    wsTemp.Cells["C" + (row +6)].Value = "Total";
                    wsTemp.Cells["C" + (row + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);


                    wsTemp.Cells["D" + (row + 6)].Value = motor;
                    wsTemp.Cells["D" + (row + 6)].Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);

                    string Path = this.CheckPath();
                    string fileName =  "Kmotors  Summary  Report (APP) ประจำช่วงวันที่ " + sd.ToString("dd MMMMM yyyy", new CultureInfo("th-TH")) + ".xlsx";
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