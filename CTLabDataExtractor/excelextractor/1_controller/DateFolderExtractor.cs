using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ExcelExtractor._1_controller
{
    class DateFolderExtractor
    {
        private static DateFolderExtractor _instance;

        private DateFolderExtractor()
        {
        }

        public static DateFolderExtractor GetInstance()
        {
            if (_instance == null)
                _instance = new DateFolderExtractor();

            return _instance;
        }

        public List<string> GetListofDateFolders(string path)
        {
            List<string> result = new List<string>();
            try
            {
                foreach (string s in Directory.GetDirectories(path))
                {
                    result.Add(s);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(Constants.PROC_ERROR_EXTRACT_DATE_FOLDER + Constants.PROC_ACTUAL_ERROR + ex.Message);
            }

            return result;
        }
    }
}
