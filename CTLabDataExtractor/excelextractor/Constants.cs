using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExtractor
{
    public static class Constants
    {
        public static string BTN_START { get { return "START"; } }
        public static string BTN_STOP { get { return "STOP"; } }

        public static string SYS_REQUIRED_DIRECTORY { get { return "\\EDMI\\CTLabDataExtractor"; } }
        public static string SYS_LAST_PROC_FILE { get { return "\\LastProc.txt"; } }
        public static string SYS_EVENT_FILE { get { return "\\Event.txt"; } }
        public static string SYS_LAST_PROC_INI_VALUE { get { return "21-11-1988"; } }
        public static string SYS_XLS_FORMAT { get { return "*.xls"; } }
        public static string SYS_DATE_FORMATE { get { return "d-M-yyyy"; } }
        public static string SYS_LAST_PROC_PREFIX { get { return "Last Proc: "; } }
        public static string SYS_MANDOTARY_FIELD_NOT_FILLED { get { return "Both Source and Destination field can not be empty"; } }
        public static string SYS_MANDOTARY_INFO { get { return "Please select appropriate location!"; } }

        public static int XLS_SHEET_INDEX { get { return 4; } }
        public static string XLS_SOURCE_SUMMARY_RANGE { get { return "A1:AJ37"; } }
        public static string XLS_SOURCE_CURRENT_RATIO_ERROR_RANGE { get { return "A77:AJ85"; } }
        public static string XLS_SOURCE_PHASE_DISP_RANGE { get { return "A88:AJ96"; } }
        public static string XLS_DEST_SUMMARY_RANGE { get { return "A1:AJ37"; } }
        public static string XLS_DEST_CURRENT_RATION_ERROR_RANGE { get { return "A45:AJ53"; } }
        public static string XLS_DEST_PHASE_DISP_RANGE { get { return "A61:AJ69"; } }
        public static string XLS_DEST_SHEET_NAME { get { return "Sheet1"; } }

        public static string EXCEP_MSG_SEPARATOR { get { return ";"; } }

        public static string PROC_START { get { return "start processing"; } }
        public static string PROC_CHECK_FOLDER { get { return "checking folder: "; } }
        public static string PROC_PROCESSING_FOLDER { get { return "processing folder: "; } }
        public static string PROC_CHECK_FILE { get { return "checking file: "; } }
        public static string PROC_PROCESSING_FILE { get { return "processing file: "; } }
        public static string PROC_SKIP_FILE { get { return "skipped file: "; } }
        public static string PROC_SKIP_FOLDER { get { return "skipped folder: "; } }
        public static string PROC_CANCEL { get { return "Process cancelled by user!"; } }
        public static string PROC_ERROR_OCCURRED { get { return "Exception error occurrs at background..."; } }
        public static string PROC_COMPLETED { get { return "File Extraction Completed!"; } }
        public static string PROC_FILE_ERROR { get { return "Error proc file->"; } }
        public static string PROC_FOLDER_ERROR { get { return "Error proc folder->"; } }
        public static string PROC_ACTUAL_ERROR { get { return ";Actual Error: "; } }
        public static string PROC_OPEN_APP { get { return "Error open Excel Applciation!"; } }
        public static string PROC_ERROR_EXTRACT_DATE_FOLDER { get { return "Error extracting date folders!"; } }
        public static string PROC_ERROR_CONSTRUCT_DIRECTORY { get { return "Error construction dest directory!"; } }
        public static string PROC_ERROR_EXTRACT_FOLDER_NAME { get { return "Error extracting folder name from dir!"; } }
        public static string PROC_ERROR_COPY_PASTE { get { return "Error copying and pasting content!"; } }

        public static string PROC_ERROR_READ_LAST_PROC { get { return "Error reading last proc!"; } }
        public static string PROC_ERROR_WRITE_LAST_PROC { get { return "Error writing last proc!"; } }
        public static string PROC_ERROR_ORDER_DATE_STRING { get { return "Error ordering folder list!"; } }
        public static string PROC_ERROR_VERIFY_DATE_STRING { get { return "Error verifying folder name!"; } }

        public static string SYS_SET_DEFAULT_LAST_PROC { get { return "Last process date has been reset to default"; } }
        public static string SYS_SET_DEFAULT_LAST_PROC_HEADER { get { return "System will process every single file in next round"; } }
        public static string SYS_CLEAR_CLIPBOARD { get { return "CLEARCLIPBOARD"; } }
    }
}
