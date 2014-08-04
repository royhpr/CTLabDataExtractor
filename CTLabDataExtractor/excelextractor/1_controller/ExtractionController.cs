using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.ComponentModel;
using System.IO;
using System.Globalization;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Reflection;

namespace ExcelExtractor._1_controller
{
    class ExtractionController
    {
        private static ExtractionController _instance;
        private LastProcController lastProcCtrl = LastProcController.GetInstance();
        private DateFolderExtractor dateFolderCtrl = DateFolderExtractor.GetInstance();
        private DirectoryController directoryCtrl = DirectoryController.GetInstance();
        private XLSFileController xlsFileCtrl = XLSFileController.GetInstance();

        private ExtractionController()
        {
        }

        public static ExtractionController GetInstance()
        {
            if (_instance == null)
                _instance = new ExtractionController();

            return _instance;
        }

        #region Private Methods

            private string ExtractDateFolderOrFilePath(string dateFolder, ref ExcelExtractor._2_model.DirectoryModel dir)
            {
                return dateFolder.Replace(dir.Source, "");
            }

            private List<string> VerifyFolderNameList(List<string> list)
            {
                List<string> result = new List<string>();

                try
                {
                    foreach (string item in list)
                    {
                        string folderName = directoryCtrl.ExtractFolderName(item);
                        DateTime dt;
                        if (DateTime.TryParseExact(folderName, Constants.SYS_DATE_FORMATE, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                        {
                            result.Add(item);
                        }
                    }

                    return result;
                }
                catch(Exception ex)
                {
                    throw new Exception(Constants.PROC_ERROR_VERIFY_DATE_STRING + Constants.PROC_ACTUAL_ERROR + ex.Message);
                }
            }

            private List<string> SortDirectoryListByDate(List<string> list)
            {
                try
                {
                    List<string> verifiedList = VerifyFolderNameList(list);
                    var orderedList = verifiedList.OrderBy(x => DateTime.ParseExact(directoryCtrl.ExtractFolderName(x), Constants.SYS_DATE_FORMATE, CultureInfo.InvariantCulture)).ToList();
                    return orderedList;
                }
                catch (Exception ex)
                {
                    throw new Exception(Constants.PROC_ERROR_ORDER_DATE_STRING + Constants.PROC_ACTUAL_ERROR + ex.Message);
                }
            }

            private void ProcessFile(ref BackgroundWorker wk, ref ExcelExtractor._2_model.DirectoryModel dir, ref Application excel, string xlsxFile)
            {
                try
                {
                    if (!directoryCtrl.IsTheDirectoryExist(xlsxFile, dir.Source, dir.Destination))
                        directoryCtrl.CreateDirectory(xlsxFile, dir.Source, dir.Destination);

                    if (!xlsFileCtrl.IsFileExist(xlsxFile, dir.Source, dir.Destination))
                    {
                        wk.ReportProgress(0, Constants.PROC_PROCESSING_FILE + ExtractDateFolderOrFilePath(xlsxFile, ref dir));
                        xlsFileCtrl.CreateXLSFile(ref wk, ref excel, xlsxFile, dir.Source, dir.Destination);
                    }
                    else
                        wk.ReportProgress(0, Constants.PROC_SKIP_FILE + ExtractDateFolderOrFilePath(xlsxFile, ref dir));
                }
                catch
                {
                    throw;
                }
            }

        #endregion

        public void StartExtraction(ref BackgroundWorker wk, ref DoWorkEventArgs e, ref ExcelExtractor._2_model.DirectoryModel dir)
        {
            wk.ReportProgress(0, Constants.PROC_START);
            Application excel = null;
            List<string> sortedDirectoryList = null;
            try
            {
                excel = new Application();
                excel.FileValidation = MsoFileValidationMode.msoFileValidationSkip;
                excel.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable;
                excel.Visible = false;
                excel.ScreenUpdating = false;
            }
            catch(Exception ex)
            {
                throw new Exception(Constants.PROC_OPEN_APP + Constants.PROC_ACTUAL_ERROR + ex.Message);
            }

            try
            {
                bool shouldWriteToLastProc = true;
                sortedDirectoryList = SortDirectoryListByDate(Directory.GetDirectories(dir.Source).ToList());
                foreach (string dateFolder in sortedDirectoryList)
                {
                    wk.ReportProgress(0, Constants.PROC_CHECK_FOLDER + ExtractDateFolderOrFilePath(dateFolder, ref dir));
                    if (DateTime.Compare(DateTime.ParseExact(directoryCtrl.ExtractFolderName(dateFolder), Constants.SYS_DATE_FORMATE, CultureInfo.InvariantCulture),
                        DateTime.ParseExact(lastProcCtrl.ReadLastProc(), Constants.SYS_DATE_FORMATE, CultureInfo.InvariantCulture)) >= 0)
                    {
                        wk.ReportProgress(0, Constants.PROC_PROCESSING_FOLDER + ExtractDateFolderOrFilePath(dateFolder, ref dir));
                        foreach (string xlsxFile in Directory.EnumerateFiles(dateFolder, Constants.SYS_XLS_FORMAT, SearchOption.AllDirectories).ToList())
                        {
                            if (wk.CancellationPending)
                            {
                                wk.ReportProgress(0, Constants.SYS_CLEAR_CLIPBOARD);
                                excel.Quit();
                                e.Cancel = true;
                                return;
                            }
                            try
                            {
                                wk.ReportProgress(0, Constants.PROC_CHECK_FILE + ExtractDateFolderOrFilePath(xlsxFile, ref dir));
                                ProcessFile(ref wk, ref dir, ref excel, xlsxFile);
                            }
                            catch (Exception ex)
                            {
                                shouldWriteToLastProc = false;
                                wk.ReportProgress(0, ex.Message);
                            }
                        }
                        if(shouldWriteToLastProc) //Any file inside any folder fail to process, should be process again
                            lastProcCtrl.WriteLastProc(dateFolder);
                    }
                    else
                        wk.ReportProgress(0, Constants.PROC_SKIP_FOLDER + ExtractDateFolderOrFilePath(dateFolder, ref dir));
                }
                excel.Quit();
            }
            catch (Exception ex)
            {
                excel.Quit();
                wk.ReportProgress(0, ex.Message);
            }
            finally
            {
                if(excel != null) Marshal.FinalReleaseComObject(excel);
            }
        }
    }
}
