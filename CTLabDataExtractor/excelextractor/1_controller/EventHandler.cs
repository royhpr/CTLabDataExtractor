using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

/*
 * Summary - trace system process status
 * 
 * 
 * 
 * 
 * 
 * */

namespace ExcelExtractor._1_controller
{
    class EventHandler
    {
        public static void Log(string msg)
        {
            StreamWriter logger;
            Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + Constants.SYS_REQUIRED_DIRECTORY);
            string fileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + Constants.SYS_REQUIRED_DIRECTORY + Constants.SYS_EVENT_FILE;
            if (!File.Exists(fileDirectory))
            {
                logger = new StreamWriter(fileDirectory);
            }
            else
            {
                logger = File.AppendText(fileDirectory);
            }

            logger.WriteLine(DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToLongTimeString() + ": " + msg);
            logger.Close();
        }
    }
}
