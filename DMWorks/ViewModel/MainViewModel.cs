using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Threading;
using System.Windows;

namespace DMWorks
{
 
    public class MainViewModel : ViewModelBase{

        private string selectedFileType;
        private string searchDirPath;
        private SldWorksFileAgent sldWorksAgent;
        private string processUpdate;
        private string selectedUnitOption;
        private bool checkDates;
        private DateTime validStartDate;
        private DateTime validEndDate;
        private ObservableCollection<SldWorksPartReport> partReports;
        private SldWorksPartReport selectedPartReport;
        private SldWorksConfigurationDetail selectedConfigDetail;

        public MainViewModel(){
            sldWorksAgent = new SldWorksFileAgent();
            partReports = new ObservableCollection<SldWorksPartReport>();
            selectedPartReport = new SldWorksPartReport();
            selectedConfigDetail = new SldWorksConfigurationDetail();

            sldWorksAgent.TaskChanged += SldWorksAgent_TaskChanged;
            DefineSearchDirectoryCommand = new RelayCommand(DefineSearchDirectoryCommandExecute);
            RunSearchCommand = new RelayCommand(RunSearchCommandExecute);
            ExportExcelCommand = new RelayCommand(ExportExcelCommandExecute);
            SelectedUnitOption = UnitOptions[0];
            validEndDate = DateTime.Now;
            validStartDate = validEndDate.AddDays(-7);
            CheckDates = true;
            selectedFileType = FileTypes[0];
        }

        public List<string> FileTypes => new List<string> { "SLDPRT" };
        public string SelectedFileType { get { return selectedFileType; } set { selectedFileType = value; RaisePropertyChanged("SelectedFileType"); } }
        public string SearchDirPath { get { return searchDirPath; } set { searchDirPath = value; RaisePropertyChanged("SearchDirPath"); } }
        public List<string> UnitOptions => new List<string> {"IPS", "MMGS", "CGS", "MKS" };
        public string SelectedUnitOption { get { return selectedUnitOption; } set { selectedUnitOption = value; RaisePropertyChanged("SelectedUnitOption"); } }
        public bool CheckDates { get { return checkDates; } set { checkDates = value; RaisePropertyChanged("CheckDates"); } }
        public DateTime ValidStartDate { get { return validStartDate; } set { validStartDate = value; RaisePropertyChanged("ValidStartDate"); } }
        public DateTime ValidEndDate { get { return validEndDate; } set { validEndDate = value; RaisePropertyChanged("ValidEndDate"); } }

        public string ProcessUpdate { 
            get { return processUpdate; } 
            set { processUpdate = value; 
                RaisePropertyChanged("ProcessUpdate"); 
                if(processUpdate.EndsWith("File Scraping Complete") && CheckDates){
                    var tmpList = new List<SldWorksPartReport>();
                    foreach (var report in PartReports){
                        tmpList.Add(report);
                    }
                    sldWorksAgent.FindMatchingCreationDates(tmpList);
                    sldWorksAgent.FindMatchingSaveDates(tmpList);
                    sldWorksAgent.FindMatchingCreationAndSaveDates(tmpList);
                }
            } 
        }

        public ObservableCollection<SldWorksPartReport> PartReports {
            get { return partReports; } 
            set { partReports = value; 
                RaisePropertyChanged("PartReports");
                if (PartReports.Count > 0) SelectedPartReport = PartReports[0];
            } 
        }
        public SldWorksPartReport SelectedPartReport { get { return selectedPartReport; } 
            set { selectedPartReport = value; 
                RaisePropertyChanged("SelectedPartReport");
                if (SelectedPartReport.ConfigDetails.Count > 0) SelectedConfigDetail = SelectedPartReport.ConfigDetails[0];
            } 
        }
        public SldWorksConfigurationDetail SelectedConfigDetail { get { return selectedConfigDetail; } set { selectedConfigDetail = value; RaisePropertyChanged("SelectedConfigDetail"); } }


        private void SldWorksAgent_TaskChanged(object sender, System.EventArgs e){
            var taskEvent = (TaskChangedEventArgs)e;
            ProcessUpdate = $"{taskEvent.EventTime} - {taskEvent.TaskDesc}";
        }

        public RelayCommand DefineSearchDirectoryCommand { get; set; }
        private void DefineSearchDirectoryCommandExecute(){
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();
            if (result == CommonFileDialogResult.Ok) SearchDirPath = dialog.FileName;
        }

        public RelayCommand RunSearchCommand { get; set; }
        private void RunSearchCommandExecute(){
            PartReports.Clear();
            if (String.IsNullOrEmpty(SearchDirPath)){
                MessageBox.Show("No search directory has been defined.");
                return;
            }

            new Thread(() =>{
                var reviewFiles = sldWorksAgent.SearchDirectory(SearchDirPath);
                sldWorksAgent.StartSldWorks();
                DateTime? startDate = ValidStartDate;
                DateTime? endDate = ValidEndDate;
                if (!checkDates){
                    startDate = null;
                    endDate = null;
                }

                foreach (var reviewFile in reviewFiles){
                    var report = sldWorksAgent.ScrapePartFile(reviewFile, SelectedUnitOption, startDate, endDate);
                    App.Current.Dispatcher.Invoke((Action)delegate{
                        PartReports.Add(report);
                    });
                }

                App.Current.Dispatcher.Invoke((Action)delegate{
                    ProcessUpdate = $"{DateTime.Now} - File Scraping Complete";
                });
            }).Start();
            if (String.IsNullOrEmpty(SelectedPartReport.FileName) && PartReports.Count > 0) SelectedPartReport = PartReports[0];
            if (String.IsNullOrEmpty(SelectedConfigDetail.Name) && SelectedPartReport.ConfigDetails.Count > 0) SelectedConfigDetail = SelectedPartReport.ConfigDetails[0];
        }

        public RelayCommand ExportExcelCommand { get; set; }
        private void ExportExcelCommandExecute(){
            if (PartReports.Count  == 0){
                MessageBox.Show("No files have been reviewed.");
                return;
            }

            var reports = new List<SldWorksPartReport>();
            foreach (var rep in PartReports){
                reports.Add(rep);
            }

            ExcelExport.ExportPartReports(@"C:\Users\brock\Desktop\TEMP\export.xlsx", reports);
        }

    }
}