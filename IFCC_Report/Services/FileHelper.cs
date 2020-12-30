using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;

namespace GSM.WEB.Services
{
    public static class FileHelper
    {
        #region getImageUrl
        public static string getImageUrl(string dataUri, string imageName, string timeStamp)
        {
            string imageUrl = string.Empty;
            string imageFullName = getTimeStamp(timeStamp) + "-" + Path.GetFileName(imageName);
            HttpRequest request = HttpContext.Current.Request;
            string uploadImagePath = request.MapPath(request.ApplicationPath + @"/upload/images/");
            if (!File.Exists(uploadImagePath + imageFullName))
            {
                var base64Data = Regex.Match(dataUri, @"data:image/(?<type>.+?),(?<data>.+)").Groups["data"].Value;
                var binData = Convert.FromBase64String(base64Data);
                Image image;
                using (MemoryStream ms = new MemoryStream(binData))
                {
                    //string name = imageName
                    image = Image.FromStream(ms);
                    if (!Directory.Exists(uploadImagePath))
                    {
                        Directory.CreateDirectory(uploadImagePath);
                    }
                    image.Save(uploadImagePath + imageFullName);
                }
            }
            imageUrl = getAppPath() + @"/upload/images/" + imageFullName;
            return imageUrl;
        }
        #endregion

        #region getFileUrl
        public static string getFileUrl(string dataUri, string FileName, string timeStamp)
        {
            string fileUrl = string.Empty;
            string fileFullName = getTimeStamp(timeStamp) + "-" + Path.GetFileName(FileName);
            HttpRequest request = HttpContext.Current.Request;
            string uploadFilePath = request.MapPath(request.ApplicationPath + @"/upload/files/");
            if (!File.Exists(uploadFilePath + fileFullName))
            {
                var base64Data = Regex.Match(dataUri, @"data:application/(?<type>.+?),(?<data>.+)").Groups["data"].Value;
                var binData = Convert.FromBase64String(base64Data);
                if (!Directory.Exists(uploadFilePath))
                {
                    Directory.CreateDirectory(uploadFilePath);
                }
                File.WriteAllBytes(uploadFilePath + fileFullName, binData);
            }
            fileUrl = getAppPath() + @"/upload/files/" + fileFullName;
            return fileUrl;
        }
        #endregion

        #region getTimeStamp
        public static string getTimeStamp(string timeStamp)
        {
            return timeStamp
                  .Replace("-", string.Empty)
                  .Replace(":", string.Empty)
                  .Replace(" ", string.Empty);
        }
        #endregion

        #region getAppPath
        public static string getAppPath()
        {
            HttpRequest request = HttpContext.Current.Request;
            string appPath = request.ApplicationPath;
            if (appPath.Substring(appPath.Length - 1, 1) == "/")
            {
                appPath = appPath.Remove(appPath.Length - 1, 1);
            }
            return appPath;
        }
        #endregion
    }
}