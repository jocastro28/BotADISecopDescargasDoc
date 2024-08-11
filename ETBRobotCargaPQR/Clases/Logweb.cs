using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ETBRobotAsignarCasosPQR.Clases
{
    internal class LogWeb
    {
        public LogWeb()
        {

        }

        public static void Send(string op, string appid, string opc, DateTime hi)
        {
            try
            {
                // get application name
                string appname = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;

                // get winuser
                string winuser = Environment.UserName;

                // Log URL
                string uriString = "http://aiblogaplicaciones.aib.com.co";

                // Create a new WebClient instance.
                WebClient wcLog = new WebClient();

                // Create a new NameValueCollection instance to hold the QueryString parameters and values.
                NameValueCollection myQueryStringCollection = new NameValueCollection
                {
                    // Assign the operation.
                    { "op", op },
                    { "wu", winuser },
                    { "ip", Environment.MachineName.ToLower() },
                    { "lla", appid },
                    { "opc", opc },
                    { "hi", hi.ToString("yyyy-MM-dd HH:mm:ss.fff") },
                    { "hf", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") }
                };

                // Attach QueryString to the WebClient.
                wcLog.QueryString = myQueryStringCollection;

                // Return the log result Web page
                // wcLog.DownloadFile(uriString, "AIB_LOG_" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".html");
                wcLog.DownloadString(uriString);
            }
            catch
            {
                throw;
            }
        }
    }
}
