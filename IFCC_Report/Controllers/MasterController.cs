using IFCC.WEB.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using IFCC.DAL;


namespace IFCC_Report.Controllers
{
    public class MasterController : Controller
    {

        public ActionResult Index()
        {
            return View();
        }

        #region MasterProjectName
        [HttpPost]
        public object MasterProjectName()
        {

            if (Response.StatusCode == 200)
            {
                try
                {

                    DataSet ds = DataHelper.GetRequestData(HttpContext);
                    DataRow dr = ds.Tables[0].Rows[0];
                  
                    DataTable dt = MasterDAL.Instance.GetProjectName();

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

    }
}
