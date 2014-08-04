using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtractor._2_model
{
    class DirectoryModel
    {
        private string _source;
        private string _destination;
        private static DirectoryModel _instance;

        private DirectoryModel()
        {
        }

        public static DirectoryModel GetInstance()
        {
            if (_instance == null)
                _instance = new DirectoryModel();

            return _instance;
        }

        public string Source
        {
            get { return _source; }
            set { _source = value; }
        }

        public string Destination
        {
            get { return _destination; }
            set { _destination = value; }
        }
    }
}
