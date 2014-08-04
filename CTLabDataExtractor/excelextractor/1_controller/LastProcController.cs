using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelExtractor._1_controller
{
    class LastProcController
    {
        private static LastProcController _instance;

        private LastProcController()
        {
            Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + Constants.SYS_REQUIRED_DIRECTORY);
            if (!File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + Constants.SYS_REQUIRED_DIRECTORY + Constants.SYS_LAST_PROC_FILE))
            {
                File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + Constants.SYS_REQUIRED_DIRECTORY + Constants.SYS_LAST_PROC_FILE, Constants.SYS_LAST_PROC_INI_VALUE);
            }
        }

        public static LastProcController GetInstance()
        {
            if (_instance == null)
                _instance = new LastProcController();

            return _instance;
        }

        public string ReadLastProc()
        {
            try
            {
                return File.ReadAllText(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + Constants.SYS_REQUIRED_DIRECTORY + Constants.SYS_LAST_PROC_FILE);
            }
            catch(Exception ex)
            {
                throw new Exception(Constants.PROC_ERROR_READ_LAST_PROC + Constants.PROC_ACTUAL_ERROR + ex.Message);
            }
        }

        public void WriteLastProc(string lastproc)
        {
            try
            {
                string folderName = lastproc.Substring(lastproc.LastIndexOf('\\') + 1, lastproc.Length - lastproc.LastIndexOf('\\') - 1);
                File.WriteAllText(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + Constants.SYS_REQUIRED_DIRECTORY + Constants.SYS_LAST_PROC_FILE, folderName);
            }
            catch(Exception ex)
            {
                throw new Exception(Constants.PROC_ERROR_WRITE_LAST_PROC + Constants.PROC_ACTUAL_ERROR + ex.Message);
            }
        }
    }
}
