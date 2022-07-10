using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace DMWorks
{
    public static class ExcelExport
    {

        public static void ExportPartReports(string filePath, List<SldWorksPartReport> reports){
            var excel = new Application();
            excel.Visible = true;

            var workBook = excel.Workbooks.Add();
            var activeSheet = (Worksheet)workBook.ActiveSheet;

            PopulatePartFeaturesSheet(activeSheet, reports);
            var cfgSheet = (Worksheet)workBook.Sheets.Add();
            PopulatePartConfigsSheet(cfgSheet, reports);
            var fileSheet = (Worksheet)workBook.Sheets.Add();
            PopulatePartFilesSheet(fileSheet, reports);
        }

        private static void PopulatePartFilesSheet(Worksheet sheet, List<SldWorksPartReport> reports){
            sheet.Name = "Files";
            sheet.Cells[1, 1].Value = "File";
            sheet.Cells[1, 2].Value = "Creation Date";
            sheet.Cells[1, 3].Value = "Last Save Date";
            sheet.Cells[1, 4].Value = "Saved By";
            sheet.Cells[1, 5].Value = "Date Violation";
            sheet.Cells[1, 6].Value = "Creation Match ID";
            sheet.Cells[1, 7].Value = "Save Match ID";
            sheet.Cells[1, 8].Value = "Creation & Save Match ID";

            var rowIdx = 2;
            foreach (var report in reports){
                sheet.Cells[rowIdx, 1].Value = report.FileName;
                sheet.Cells[rowIdx, 2].Value = report.CreationDate;
                sheet.Cells[rowIdx, 3].Value = report.LastSaveDate;
                sheet.Cells[rowIdx, 4].Value = report.SavedBy;
                sheet.Cells[rowIdx, 5].Value = report.DateViolation;
                sheet.Cells[rowIdx, 6].Value = report.CreationMatchId;
                sheet.Cells[rowIdx, 7].Value = report.SaveMatchId;
                sheet.Cells[rowIdx, 8].Value = report.CreationSaveMatchId;
                rowIdx++;
            }

            sheet.Range["A1", "H1"].Columns.AutoFit();
            sheet.Range["A1", "H1"].Font.Bold = true;
            sheet.Range["A2", "D2"].Columns.AutoFit();
        }

        private static void PopulatePartConfigsSheet(Worksheet sheet, List<SldWorksPartReport> reports){
            sheet.Name = "Configs";
            sheet.Cells[1, 1].Value = "File";
            sheet.Cells[1, 2].Value = "Configuration";
            sheet.Cells[1, 3].Value = "Material";
            sheet.Cells[1, 4].Value = "Mass";
            sheet.Cells[1, 5].Value = "Volume";
            sheet.Cells[1, 6].Value = "Center of Mass";

            var rowIdx = 2;
            foreach (var report in reports){
                foreach (var config in report.ConfigDetails){
                    sheet.Cells[rowIdx, 1].Value = report.FileName;
                    sheet.Cells[rowIdx, 2].Value = config.Name;
                    sheet.Cells[rowIdx, 3].Value = config.Material;
                    sheet.Cells[rowIdx, 4].Value = config.Mass_user;
                    sheet.Cells[rowIdx, 5].Value = config.Volume_user;
                    sheet.Cells[rowIdx, 6].Value = config.CenterOfMass_user;
                    rowIdx++;
                }
            }

            sheet.Range["A1", "F1"].Columns.AutoFit();
            sheet.Range["A1", "F1"].Font.Bold = true;
            sheet.Range["C2"].Columns.AutoFit();
        }

        private static void PopulatePartFeaturesSheet(Worksheet sheet, List<SldWorksPartReport> reports){
            sheet.Name = "Features";
            sheet.Cells[1, 1].Value = "File";
            sheet.Cells[1, 2].Value = "Configuration";
            sheet.Cells[1, 3].Value = "Feature";
            sheet.Cells[1, 4].Value = "Created By";
            sheet.Cells[1, 5].Value = "Creation Date";
            sheet.Cells[1, 6].Value = "Modified Date";
            sheet.Cells[1, 7].Value = "Description";
            sheet.Cells[1, 8].Value = "Is Sketch?";
            sheet.Cells[1, 9].Value = "Sketch Status";

            var rowIdx = 2;
            foreach (var report in reports){
                foreach (var config in report.ConfigDetails){
                    foreach (var feat in config.FeatureDetails){
                        sheet.Cells[rowIdx, 1].Value = report.FileName;
                        sheet.Cells[rowIdx, 2].Value = config.Name;
                        sheet.Cells[rowIdx, 3].Value = feat.Name;
                        sheet.Cells[rowIdx, 4].Value = feat.CreatedBy;
                        sheet.Cells[rowIdx, 5].Value = feat.CreationDate;
                        sheet.Cells[rowIdx, 6].Value = feat.ModifiedDate;
                        sheet.Cells[rowIdx, 7].Value = feat.Description;
                        sheet.Cells[rowIdx, 8].Value = feat.IsSketch;
                        sheet.Cells[rowIdx, 9].Value = feat.SketchStatus;
                        rowIdx++;
                    }
                }
            }

            sheet.Range["A1", "I1"].Font.Bold = true;
            sheet.Range["A1", "I1"].Columns.AutoFit();
            sheet.Range["A2"].Columns.AutoFit();
        }
    }
}
