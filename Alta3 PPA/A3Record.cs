using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Alta3_PPA
{
    class A3Record
    {
        public static void PostIt(Uri uri, string json)
        {
            if (uri.LocalPath == "invalid")
            {
                System.Windows.Forms.MessageBox.Show("ERROR: URI Is Incorrect - Please rewrite");
                return;
            }
            HttpWebRequest httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";

            try
            {
                using (StreamWriter streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                {
                    streamWriter.Write(json);
                    streamWriter.Flush();
                    streamWriter.Close();
                }
            }
            catch
            {
                MessageBox.Show("FAILED TO SEND! IS THE SERVER UP?", "ERROR!", MessageBoxButtons.OK);
                return;
            }

            try
            {
                HttpWebResponse httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (StreamReader streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
                {
                    string result = streamReader.ReadToEnd();
                    System.Windows.Forms.MessageBox.Show(result);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("NO RESPONSE RECIEVED", "ERROR!", MessageBoxButtons.OK);
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
