using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace WMSFFService
{
    public class MyCustomFunctions
    {
        private string logFile = AppDomain.CurrentDomain.BaseDirectory +
"MyApplicationLog.txt";
        public void WriteToLogFile(string strMessage)
        {
            try
            {
                
                string line = DateTime.Now.ToString() + " | ";
                line += strMessage;
                FileStream fs = new FileStream(logFile, FileMode.Append, FileAccess.Write, FileShare.None);
                StreamWriter swFromFileStream = new StreamWriter(fs);
                swFromFileStream.WriteLine(line);
                swFromFileStream.Flush();
                swFromFileStream.Close();
            }
            catch(Exception ex)
            {

            }
        }
    }
}
