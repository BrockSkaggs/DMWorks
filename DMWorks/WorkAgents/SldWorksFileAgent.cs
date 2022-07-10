using SldWorks;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DMWorks
{
    public class SldWorksFileAgent : TaskChangedBase {

        //SolidWorks 2021 is version number 29.
        private const string sldWorksId = "SldWorks.Application.29";
        private ISldWorks sldWorksApp;

        public SldWorksFileAgent(){

        }

        public List<FileInfo> SearchDirectory(string dirPath){
            var reviewFiles = new List<FileInfo>();
            var di = new DirectoryInfo(dirPath);
            foreach (var file in di.GetFiles()){
                //TODO: Add logic to filter files.
                reviewFiles.Add(file);
            }
            return reviewFiles;
        }

        public void StartSldWorks(){
            OnTaskChanged("Accessing SOLIDWORKS...", DateTime.Now);
            Type swType = Type.GetTypeFromProgID(sldWorksId);
            ISldWorks swApp = (ISldWorks)Activator.CreateInstance(swType);
            swApp.Visible = true;
            sldWorksApp = swApp;
        }

        private int DetermineSwDocType(string fileExtension){
            //Ref: http://help.solidworks.com/2020/english/api/swconst/SOLIDWORKS.Interop.swconst~SOLIDWORKS.Interop.swconst.swDocumentTypes_e.html
            if (fileExtension.ToLower().EndsWith("sldprt")) return 1;
            if (fileExtension.ToLower().EndsWith("sldasm")) return 2;
            if (fileExtension.ToLower().EndsWith("slddrw")) return 3;
            return -1;
        }

        private SldWorksFeatureDetail ScrapeFeature(Feature feature){
            var featDet = new SldWorksFeatureDetail();
            featDet.Name = feature.Name;
            featDet.CreationDate = Convert.ToDateTime(feature.DateCreated);
            featDet.ModifiedDate = Convert.ToDateTime(feature.DateModified);
            featDet.CreatedBy = feature.CreatedBy;
            featDet.Description = feature.Description;

            if(new List<string> { "3DProfileFeature", "ProfileFeature" }.Contains(feature.GetTypeName2())){
                featDet.IsSketch = true;
                Sketch sketch = feature.GetSpecificFeature2();
                var sketchStatus = sketch.GetConstrainedStatus();
                if (sketchStatus == (int)SwConst.swConstrainedStatus_e.swUnknownConstraint) featDet.SketchStatus = "swUnknownConstraint";
                else if (sketchStatus == (int)SwConst.swConstrainedStatus_e.swUnderConstrained) featDet.SketchStatus = "swUnderConstrained";
                else if (sketchStatus == (int)SwConst.swConstrainedStatus_e.swFullyConstrained) featDet.SketchStatus = "swFullyConstrained";
                else if (sketchStatus == (int)SwConst.swConstrainedStatus_e.swOverConstrained) featDet.SketchStatus = "swOverConstrained";
                else if (sketchStatus == (int)SwConst.swConstrainedStatus_e.swNoSolution) featDet.SketchStatus = "swNoSolution";
                else if (sketchStatus == (int)SwConst.swConstrainedStatus_e.swInvalidSolution) featDet.SketchStatus = "swInvalidSolution";
                else if (sketchStatus == (int)SwConst.swConstrainedStatus_e.swAutosolveOff) featDet.SketchStatus = "swAutoSolveOff";
            }
            
            return featDet;
        }

        private void ConvertMassPropsToUserUnits(SldWorksConfigurationDetail cfgDet, string userUnits){
            Double kilogramToPoundFactor = 2.204622;
            if(userUnits == "IPS"){
                cfgDet.Mass_user = cfgDet.Mass_kg * kilogramToPoundFactor;
                cfgDet.Volume_user = cfgDet.Volume_m3 * Math.Pow(100, 3) * (1 / Math.Pow(2.54, 3));
                for (int i = 0; i < 3; i++){
                    cfgDet.CenterOfMass_user[i] = cfgDet.CenterOfMass_m[i] * (100 / 2.54);
                }
            }
            else if (userUnits == "MKS"){
                cfgDet.Mass_user = cfgDet.Mass_kg;
                cfgDet.Volume_user = cfgDet.Volume_m3;
                cfgDet.CenterOfMass_user = cfgDet.CenterOfMass_m;
            }
            else if (userUnits == "MMGS"){
                cfgDet.Mass_user = cfgDet.Mass_kg / 1000;
                cfgDet.Volume_user = cfgDet.Volume_m3 * Math.Pow(1000, 3);
                for (int i = 0; i < 3; i++){
                    cfgDet.CenterOfMass_user[i] = cfgDet.CenterOfMass_m[i] * 1000;
                }
            }
            else if (userUnits == "CGS"){
                cfgDet.Mass_user = cfgDet.Mass_kg / 1000;
                cfgDet.Volume_user = cfgDet.Volume_m3 * Math.Pow(100, 3);
                for (int i = 0; i < 3; i++){
                    cfgDet.CenterOfMass_user[i] = cfgDet.CenterOfMass_m[i] * 100;
                }
            }
        }

        private void CheckForDateErrors(SldWorksPartReport report, DateTime validStart, DateTime validEnd){
            if(report.CreationDate < validStart || report.CreationDate > validEnd){
                report.DateViolation = true;
                report.WarningNote += "Invalid creation date.\n";
            }
            if (report.LastSaveDate < validStart || report.LastSaveDate > validEnd){
                report.DateViolation = true;
                report.WarningNote += "Invalid lave save date.\n";
            }
        }

        public SldWorksPartReport ScrapePartFile(FileInfo file, string userUnits, DateTime? validStartDate = null, DateTime? validEndDate = null){
            
            var docType = DetermineSwDocType(file.Extension);
            var report = new SldWorksPartReport{ FileName = file.Name};

            var errors = 0;
            var warnings = 0;
            var mDoc = sldWorksApp.OpenDoc6(file.FullName, docType, 1, "", errors, warnings);
            var massProps = mDoc.Extension.CreateMassProperty();

            report.SavedBy = mDoc.SummaryInfo[(int)SwConst.swSummInfoField_e.swSumInfoSavedBy];
            //Ref: http://help.solidworks.com/2020/English/api/swconst/SolidWorks.Interop.swconst~SolidWorks.Interop.swconst.swSummInfoField_e.html?verRedirect=1
            report.CreationDate = Convert.ToDateTime(mDoc.SummaryInfo[(int)SwConst.swSummInfoField_e.swSumInfoCreateDate2]); 
            report.LastSaveDate = Convert.ToDateTime(mDoc.SummaryInfo[(int)SwConst.swSummInfoField_e.swSumInfoSaveDate2]);
            if(validStartDate != null) CheckForDateErrors(report, validStartDate.GetValueOrDefault(), validEndDate.GetValueOrDefault());

            var cfgNames = mDoc.GetConfigurationNames();
            foreach (var cfgName in cfgNames){
                if(sldWorksApp.GetActiveConfigurationName(file.FullName) != cfgName) mDoc.ShowConfiguration2(cfgName);
                var cfgDet = new SldWorksConfigurationDetail();
                cfgDet.Name = cfgName;
                cfgDet.Material = mDoc.MaterialIdName;
                cfgDet.Mass_kg = massProps.Mass;
                cfgDet.Volume_m3 = massProps.Volume;
                cfgDet.CenterOfMass_m = (double[])massProps.CenterOfMass;
                for (int i = 0; i < mDoc.GetFeatureCount(); i++){
                    Feature swFeature = mDoc.FeatureByPositionReverse(i);
                    if (swFeature.Name.Contains("Origin")) break; //Stop looking at features above the Material node.
                    cfgDet.FeatureDetails.Add(ScrapeFeature(swFeature));
                }
                ConvertMassPropsToUserUnits(cfgDet, userUnits);
                report.ConfigDetails.Add(cfgDet);
            }

            sldWorksApp.CloseDoc("");

            OnTaskChanged($"Reviewed File: {file.Name}", DateTime.Now);
            return report;
        }

        public void FindMatchingCreationDates(List<SldWorksPartReport> reports){
            int matchId = 0;
            foreach (var report in reports){
                if (!String.IsNullOrEmpty(report.CreationMatchId)) continue;
                var matchingCreationDates = reports.Where(x => x.CreationDate == report.CreationDate && x.FileName != report.FileName).ToList();
                if (matchingCreationDates.Count == 0) continue;

                foreach (var match in matchingCreationDates){
                    match.CreationMatchId = Convert.ToString(matchId);
                }
                report.CreationMatchId = Convert.ToString(matchId);
                matchId++;
            }
        }

        public void FindMatchingSaveDates(List<SldWorksPartReport> reports){
            int matchId = 0;
            foreach (var report in reports){
                if (!String.IsNullOrEmpty(report.SaveMatchId)) continue;
                var matchingSaveDates = reports.Where(x => x.LastSaveDate == report.LastSaveDate && x.FileName != report.FileName).ToList();
                if (matchingSaveDates.Count == 0) continue;

                foreach (var match in matchingSaveDates){
                    match.SaveMatchId = Convert.ToString(matchId);
                }
                report.SaveMatchId = Convert.ToString(matchId);
                matchId++;
            }
        }

        public void FindMatchingCreationAndSaveDates(List<SldWorksPartReport> reports){
            int matchId = 0;
            foreach (var report in reports){
                if (String.IsNullOrEmpty(report.CreationMatchId) || String.IsNullOrEmpty(report.SaveMatchId)) continue;

                var matchList = reports.Where(x => x.CreationMatchId == report.CreationMatchId && x.SaveMatchId == report.SaveMatchId && x.FileName != report.FileName).ToList();
                if (matchList.Count == 0) continue;
                foreach (var match in matchList){
                    match.CreationSaveMatchId = Convert.ToString(matchId);
                }
                report.CreationSaveMatchId = Convert.ToString(matchId);
                matchId++;
            }
        }


    }
}
