using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace ExcelExtractor._1_controller
{
    class XLSFileController
    {
        private static XLSFileController _instance;

        private XLSFileController()
        {
        }

        public static XLSFileController GetInstance()
        {
            if (_instance == null)
                _instance = new XLSFileController();

            return _instance;
        }

        private string ConstructDestFilePath(string fileName, string source, string dest)
        {
            return fileName.Replace(source, dest);
        }

        public bool IsFileExist(string fileName, string source, string dest)
        {
            string destPath = ConstructDestFilePath(fileName, source, dest);

            return File.Exists(destPath);
        }

        private void CopyPathContent(ref  Range src, ref Range dest, ref BackgroundWorker wk)
        {
            try
            {
                src.Copy(Type.Missing);
                dest.PasteSpecial(XlPasteType.xlPasteFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                wk.ReportProgress(0, Constants.SYS_CLEAR_CLIPBOARD);

                src.Copy(Type.Missing);
                dest.PasteSpecial(XlPasteType.xlPasteValuesAndNumberFormats, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                wk.ReportProgress(0, Constants.SYS_CLEAR_CLIPBOARD);

                src.Copy(Type.Missing);
                dest.PasteSpecial(XlPasteType.xlPasteColumnWidths, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                wk.ReportProgress(0, Constants.SYS_CLEAR_CLIPBOARD);
            }
            catch (Exception ex)
            {
                throw new Exception(Constants.PROC_ERROR_COPY_PASTE + Constants.PROC_ACTUAL_ERROR + ex.Message);
            }
        }

        public void CreateXLSFile(ref BackgroundWorker wk, ref Application excel, string fileName, string source, string dest)
        {
            var workbooks = excel.Workbooks;
            Workbook wbSrc = null;
            Workbook wbDest = null;
            Worksheet wbSrcSheet = null;
            Worksheet wbDestSheet = null;
            Range srcSummaryRange = null;
            Range destSummaryRange = null;
            Range srcCurrentRatioErrRange = null;
            Range destCurrentRatioErrRange = null;
            Range srcPhaseDispRange = null;
            Range destPhaseDispRange = null;
            try
            {
                wbSrc = workbooks.Open(fileName);
                wbDest = workbooks.Add();
                wbSrcSheet = wbSrc.Sheets[Constants.XLS_SHEET_INDEX];
                wbDestSheet = wbDest.Sheets[Constants.XLS_DEST_SHEET_NAME];

                srcSummaryRange = wbSrcSheet.get_Range(Constants.XLS_SOURCE_SUMMARY_RANGE);
                destSummaryRange = wbDestSheet.get_Range(Constants.XLS_DEST_SUMMARY_RANGE);
                srcCurrentRatioErrRange = wbSrcSheet.get_Range(Constants.XLS_SOURCE_CURRENT_RATIO_ERROR_RANGE);
                destCurrentRatioErrRange = wbDestSheet.get_Range(Constants.XLS_DEST_CURRENT_RATION_ERROR_RANGE);
                srcPhaseDispRange = wbSrcSheet.get_Range(Constants.XLS_SOURCE_PHASE_DISP_RANGE);
                destPhaseDispRange = wbDestSheet.get_Range(Constants.XLS_DEST_PHASE_DISP_RANGE);

                excel.DisplayAlerts = false;
                CopyPathContent(ref srcSummaryRange, ref destSummaryRange, ref wk);
                CopyPathContent(ref srcCurrentRatioErrRange, ref destCurrentRatioErrRange, ref wk);
                CopyPathContent(ref srcPhaseDispRange, ref destPhaseDispRange, ref wk);

                wbSrc.Close(false, Missing.Value, Missing.Value);
                wbDest.SaveAs(ConstructDestFilePath(fileName, source, dest), Missing.Value,
                            Missing.Value, Missing.Value, Missing.Value, Missing.Value, XlSaveAsAccessMode.xlNoChange,
                            Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                wbDest.Close(false, Missing.Value, Missing.Value);
                excel.DisplayAlerts = true;
            }
            catch (Exception ex)
            {
                wk.ReportProgress(0, Constants.SYS_CLEAR_CLIPBOARD);
                throw new Exception(Constants.PROC_FILE_ERROR + fileName + Constants.PROC_ACTUAL_ERROR + ex.Message);
            }
            finally
            {
                if (srcSummaryRange != null) Marshal.ReleaseComObject(srcSummaryRange);
                if (destSummaryRange != null) Marshal.ReleaseComObject(destSummaryRange);
                if (srcCurrentRatioErrRange != null) Marshal.ReleaseComObject(srcCurrentRatioErrRange);
                if (destCurrentRatioErrRange != null) Marshal.ReleaseComObject(destCurrentRatioErrRange);
                if (srcPhaseDispRange != null) Marshal.ReleaseComObject(srcPhaseDispRange);
                if (destPhaseDispRange != null) Marshal.ReleaseComObject(destPhaseDispRange);

                if (wbSrcSheet != null) Marshal.ReleaseComObject(wbSrcSheet);
                if (wbDestSheet != null) Marshal.ReleaseComObject(wbDestSheet);

                if (wbSrc != null) Marshal.ReleaseComObject(wbSrc);
                if (wbDest != null) Marshal.ReleaseComObject(wbDest);

                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
            }
        }
    }
}
