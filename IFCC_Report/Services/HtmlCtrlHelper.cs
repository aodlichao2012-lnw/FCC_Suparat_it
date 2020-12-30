using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Bunker.Web.Services
{
    public static class HtmlCtrlHelper
    {
        public static MvcHtmlString GeneratePermissionMenu(object disableMenu,string iconClass = "", string linkTo = "", string caption="",string customAttributes = "")
        {
            if(!(bool)disableMenu)
            {
                #region Generate
                //<li class="">
                //< a href = "@Url.Action("Index", "ArrivalNotice")" >
                //< i class="fa fa-cube"></i> แจ้งตู้เข้า
                //</a>
                //</li>

                string liHyperLinkMenuFormat = String.Format("<li class=''><a href='{0}' {3} ><i class='{1}'></i> {2} </a></li>", linkTo, iconClass, caption, customAttributes);
                return MvcHtmlString.Create(liHyperLinkMenuFormat);
                #endregion
            }
            else
            {
                return MvcHtmlString.Empty;
            }
        }
    }
}