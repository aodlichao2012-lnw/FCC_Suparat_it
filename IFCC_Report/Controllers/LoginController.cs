using Bunker.Web.Services;
using GSM.WEB.Services;
using IFCC.DAL.IFCC3;
using IFCC.WEB.Services;
using IFCC_Report.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Services.Client;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using System.Web.UI;


namespace IFCC_Report.Controllers
{
    public class LoginController : Controller
    {


        // GET: Login
        public ActionResult Index()
        {
            if(Session != null)
            {
                FormsAuthentication.SignOut();
                Session.Clear();
            }
            else if(Session == null)
            {
                return RedirectToAction("Sigin", "Login");
            }
            return View();
        }


        #region Signin
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Signin()
        {
            try
            {
              
                //string domain = HttpContext.Request.Form["domain"];
                string username = HttpContext.Request.Form["username"];
                string password = HttpContext.Request.Form["password"];
                DataTable dtUser = new DataTable();
                dtUser.Columns.Add("username");
                DataRow drUser = null;
                //DataSet ds = DataHelper.GetRequestData(HttpContext);

                if (Check(username , password))
                {
                    drUser = dtUser.NewRow();
                    //return RedirectToAction("Index", "Student");
                    drUser["username"] = username;
                    dtUser.Rows.Add(drUser);
                }
                else
                {
                    TempData["IsAuth"] = false;
                    TempData["Message"] = "Wrong Password";
                    return RedirectToAction("Index", "login");
                }
                Session["user"] = dtUser;
                string returnUrl = HttpContext.Request.Form["returnUrl"];
                if (!string.IsNullOrEmpty(returnUrl))
                {
                    return Redirect(returnUrl);
                }
                return RedirectToAction("Index", "Home");

            }
            catch (Exception ex)
            {
                ex.Message.ToString();
                TempData["IsAuth"] = false;
                ExceptionHelper.AddException(ex);
                return RedirectToAction("Index", "Error");
            }
        }
        #endregion

        #region Logout

        [HttpPost]
        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            Session.Clear();
            return RedirectToAction("Index", "Login");
        }

        #endregion

        #region CheckUser_Password

        public bool Check(string user, string pass)
        {
            iFCC_WS _db = new iFCC_WS();
            
           if(user == "" || pass == "")
            {
                return false;
            }
            else if (_db.authenApplication(projectID: "1042", user, pass))
            {
                return true;
            }
            else
            {
                TempData["IsAuth"] = false;
                TempData["Message"] = "Wrong Password";
                return false;
            }
        }

        #endregion

    }
}