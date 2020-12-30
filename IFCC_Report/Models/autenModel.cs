using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Xml.Serialization;

namespace IFCC_Report.Models
{
    [XmlRoot("authenApplication", Namespace = "")]
    public class autenModel
    {
        [XmlElement("projectID")]
        public string ProjectID { get; set; }
        [XmlElement("username")]
        public string username { get; set; }
        [XmlElement("password")]
        public string password { get; set; }
    }
}