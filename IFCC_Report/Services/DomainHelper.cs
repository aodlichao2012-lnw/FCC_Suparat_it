using System;
using System.Configuration;
using System.Globalization;
using System.Reflection.PortableExecutable;
using System.Threading;

namespace GSM.WEB.Services
{
    public enum SystemLanguage { English, Thai };

    public class DateTimeHelper
    {
        public static void SetDefaultCulture(SystemLanguage sysLanguage)
        {
            CultureInfo culture = CultureInfo.CreateSpecificCulture("en-US");
            // Change current culture
            if (sysLanguage == SystemLanguage.Thai)
            {
                culture = CultureInfo.CreateSpecificCulture("th-TH");
            }
            else
            {
                culture = CultureInfo.CreateSpecificCulture("en-US");
            }

            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
        }

        public static CultureInfo GetCultureInfo(SystemLanguage sysLanguage)
        {
            if (sysLanguage == SystemLanguage.Thai)
            {
                return CultureInfo.CreateSpecificCulture("th-TH");
            }
            else
            {
                return CultureInfo.CreateSpecificCulture("en-US");
            }
        }
    }
}