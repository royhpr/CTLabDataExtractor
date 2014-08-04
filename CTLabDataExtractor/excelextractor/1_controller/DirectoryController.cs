using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

/*
 * 
 * 
 * 
 * 
 * 
 * */

namespace ExcelExtractor._1_controller
{
    class DirectoryController
    {
        private static DirectoryController _instance;

        private DirectoryController()
        {
        }

        public static DirectoryController GetInstance()
        {
            if (_instance == null)
                _instance = new DirectoryController();

            return _instance;
        }

        private string ConstructTargetDirectory(string fileDirectory, string source, string dest)
        {
            try
            {
                string result = fileDirectory.Replace(source, dest);
                return Path.GetDirectoryName(result);
            }
            catch(Exception ex)
            {
                throw new Exception(Constants.PROC_ERROR_CONSTRUCT_DIRECTORY +  Constants.PROC_ACTUAL_ERROR + ex.Message);
            }
        }

        public bool IsTheDirectoryExist(string fileDirectory, string source, string dest)
        {
            try
            {
                string destDirectory = ConstructTargetDirectory(fileDirectory, source, dest);

                return Directory.Exists(destDirectory);
            }
            catch
            {
                throw;
            }
        }

        public void CreateDirectory(string fileDirectory, string source, string dest)
        {
            try
            {
                string destDirectory = ConstructTargetDirectory(fileDirectory, source, dest);

                Directory.CreateDirectory(destDirectory);
            }
            catch
            {
                throw;
            }
        }

        public string ExtractFolderName(string directoryPath)
        {
            try
            {
                return directoryPath.Substring(directoryPath.LastIndexOf('\\') + 1, directoryPath.Length - directoryPath.LastIndexOf('\\') - 1);
            }
            catch(Exception ex)
            {
                throw new Exception(Constants.PROC_ERROR_EXTRACT_FOLDER_NAME + Constants.PROC_ACTUAL_ERROR + ex.Message);
            }
        }
    }
}
