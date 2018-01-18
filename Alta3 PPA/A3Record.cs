using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Alta3_PPA
{
    class A3Record
    {
        public static void PostIt(Uri uri, string json)
        {
            if (uri.LocalPath == "invalid")
            {
                System.Windows.Forms.MessageBox.Show("ERROR: Please rewrite");

            }
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";
            
            using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                streamWriter.Write(json);
                streamWriter.Flush();
                streamWriter.Close();
            }

            HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
            using (StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
            {
                string result = streamReader.ReadToEnd();
                System.Windows.Forms.MessageBox.Show(result);
            }
        }
       
        public static Uri ConvertToUri(string uri)
        {
            try
            {
                return new Uri(uri);
            }
            catch
            {
                return new Uri(@"http://127.0.0.1:8000/invalid");
            }
        }
    }
}
