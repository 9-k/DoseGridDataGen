using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using ESAPIUtilities;
using System.Data;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.1")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
[assembly: ESAPIScript(IsWriteable = true)]

namespace DoseGridDataGen
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                using (Application app = Application.CreateApplication())
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }
        }
        static void Execute(Application app)
        {
            // It doesn't make sense to also make trackedparams external because this isnt nor should
            // it be user settable. 
            List<string> TRACKEDPARAMS = new List<string>()
            {
                // target stuff
                "Actual Target Name",
                "TG263 Target Name Guess",
                "Target Volume",
                "Target Mean Dose",
                "Target Hot Spot",
                "Target V95",
                // oar stuff
                "Actual OAR Name",
                "TG263 OAR Name Guess",
                "OAR Volume",
                "OAR Mean Dose",
                "OAR D0.03",
                // misc
                "Minimum Separation",
                "Centroid Separation",
                "Grid Size",
                "MRN",
                "Course Name",
                "Plan Name",
                "Plan Dx"
            };

            DataTable result = new DataTable();
            foreach (string columnName in TRACKEDPARAMS) { result.Columns.Add(columnName); }

            string settingsPath = "settings.txt";
            //string assyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            Dictionary <string,string> settingsDict = ESAPIUtility.SettingsDict(settingsPath);

            string mrnPath = Path.GetFullPath(settingsDict["mrnPath"]);
            string TG263Path = Path.GetFullPath(settingsDict["TG263Path"]);
            string ICDCodesPath = Path.GetFullPath(settingsDict["ICDCodesPath"]);
            string doseGridsPath = Path.GetFullPath(settingsDict["doseGridsPath"]);

            string projectCourseID = settingsDict["projectCourseID"];
            int PARSEMATCHCUTOFF = Int32.Parse(settingsDict["parseMatchCutoff"]);
            List<string> DOSEGRIDS = File.ReadAllLines(doseGridsPath).ToList();
            List<string> ICDCODELIST = File.ReadAllLines(ICDCodesPath).ToList();  

            DataTable TG263Table = new DataTable();
            using (var reader = new StreamReader(TG263Path))
            {
                var headers = reader.ReadLine().Split(',');

                // Create columns based on CSV headers
                foreach (var header in headers)
                {
                    TG263Table.Columns.Add(header);
                }

                // Read rows and add them to the DataTable
                while (!reader.EndOfStream)
                {
                    var rowValues = reader.ReadLine().Split(',');
                    TG263Table.Rows.Add(rowValues);
                }
            }

            IEnumerable<string> mrnList = File.ReadLines(mrnPath)
                .Select(line => line.Trim())
                .Where(line => !string.IsNullOrEmpty(line));

            foreach (string mrn in mrnList)
            {
                // copying between patients is not possible.
                // I'm just going to do everything in situ then NOT save any modifications... so it all gets rolled back.
                Patient patient = app.OpenPatientById(mrn);
                patient.BeginModifications();
                List<Course> coursesMatchingDxList =
                    patient.Courses
                    .Where(course => course.Diagnoses
                    .Any(dx => ICDCODELIST.Any(code => dx.Code.Contains(code)))).ToList();

                // now I don't know why this would ping but just in case:
                if (coursesMatchingDxList.Count <= 0) { continue; }

                Course GridProj = patient.AddCourse();
                GridProj.Id = projectCourseID;

                foreach (Course course in coursesMatchingDxList)
                {
                    // if course name contains qa,,,, skip?
                    string CourseDx = course.Diagnoses.Where(dx => ICDCODELIST.Any(code => dx.Code.Contains(code))).First().Code;
                    foreach (ExternalPlanSetup originalPlan in course.ExternalPlanSetups)
                    {
                        if (!originalPlan.IsTreated) { continue; }

                        ExternalPlanSetup plan = GridProj.CopyPlanSetup(originalPlan) as ExternalPlanSetup;
                        string ActualTargetName = plan.TargetVolumeID;
                        Structure targetStructure = plan.StructureSet.Structures.Where(structure => structure.Id == ActualTargetName).FirstOrDefault();

                        string TG263TargetNameGuess = StructureParserMethods.StructureParser(
                            ActualTargetName, 
                            TG263Table, 
                            cutoff: PARSEMATCHCUTOFF)
                            .matches.First().Item1["TG263-Primary Name"] as string;
                        double TargetVolume = targetStructure.Volume;
                        double TargetMeanDose = plan.GetDVHCumulativeData(
                            targetStructure, 
                            DoseValuePresentation.Absolute, 
                            VolumePresentation.Relative, 
                            1).MeanDose.Dose;

                        double TargetHotSpot = plan.GetDVHCumulativeData(
                            targetStructure,
                            DoseValuePresentation.Absolute,
                            VolumePresentation.Relative,
                            1).MaxDose.Dose;

                        double TargetV95 = ESAPIUtility.GetVXX(targetStructure, plan, 95, VolumePresentation.Relative);

                        foreach (string gridSize in DOSEGRIDS)
                        {
                            plan.SetCalculationOption(plan.PhotonCalculationModel, "CalculationGridSizeInCM", gridSize);
                            plan.SetCalculationOption(plan.PhotonCalculationModel, "CalculationGridSizeInCMForSRSAndHyperArc", gridSize);
                            plan.CalculateDose();
                            foreach (Structure oar in plan.StructureSet.Structures)
                            {
                                string ActualOARName = oar.Id;
                                TG263ParseResult TG263OARParse = StructureParserMethods.StructureParser(
                                    ActualOARName,
                                    TG263Table,
                                    cutoff: PARSEMATCHCUTOFF);

                                if (oar.Id == "BODY" ||
                                    oar.Id == ActualTargetName ||
                                    TG263OARParse.IsEval ||
                                    TG263OARParse.IsPRV ||
                                    TG263OARParse.IsOpti ||
                                    TG263OARParse.IsDerived ||
                                    TG263OARParse.IsPlanning ||
                                    TG263OARParse.matches.Count < 1
                                    ) { continue; }
                                string TG263OARNameGuess = TG263OARParse.matches.First().Item1["TG263-Primary Name"] as string;
                                double OARVolume = oar.Volume;
                                double OARMeanDose = plan.GetDVHCumulativeData(
                                    targetStructure,
                                    DoseValuePresentation.Absolute,
                                    VolumePresentation.Relative,
                                    1).MeanDose.Dose;

                                double OARHotSpot = plan.GetDVHCumulativeData(
                                    targetStructure,
                                    DoseValuePresentation.Absolute,
                                    VolumePresentation.Relative,
                                    1).MaxDose.Dose;

                                double OARD003 = plan.GetDoseAtVolume(oar, 0.03, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;

                                double MinimumSeparation = ESAPIUtility.MinimumStructureDistance(targetStructure, oar);
                                double CentroidSeparation = (targetStructure.CenterPoint - oar.CenterPoint).Length;
                                DataRow newRow = result.NewRow();
                                newRow["Actual Target Name"] = ActualTargetName;
                                newRow["TG263 Target Name Guess"] = TG263TargetNameGuess;
                                newRow["Target Volume"] = TargetVolume.ToString();
                                newRow["Target Mean Dose"] = TargetMeanDose.ToString();
                                newRow["Target Hot Spot"] = TargetHotSpot.ToString();
                                newRow["Target V95"] = TargetV95.ToString();
                                // oar stuff
                                newRow["Actual OAR Name"] = ActualOARName;
                                newRow["TG263 OAR Name Guess"] = TG263OARNameGuess;
                                newRow["OAR Volume"] = OARVolume.ToString();
                                newRow["OAR Mean Dose"] = OARMeanDose.ToString();
                                newRow["OAR D0.03"] = OARD003.ToString();
                                // misc
                                newRow["Minimum Separation"] = MinimumSeparation.ToString();
                                newRow["Centroid Separation"] = CentroidSeparation.ToString();
                                newRow["Grid Size"] = gridSize;
                                newRow["MRN"] = mrn.ToString();
                                newRow["Course Name"] = course.Id;
                                newRow["Plan Name"] = originalPlan.Id;
                                newRow["Course Dx"] = CourseDx;
                                result.Rows.Add(newRow);
                            }
                        }
                    }
                }
                app.ClosePatient();
            }

            string EscapeCsvField(string field)
            {
                if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
                {
                    return "\"" + field.Replace("\"", "\"\"") + "\"";
                }
                return field;
            }


            using (StreamWriter writer = new StreamWriter("output.csv", false, Encoding.UTF8))
            {
                // Optionally write the column headers
                var columnNames = result.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                writer.WriteLine(string.Join(",", columnNames));

                // Write each row
                foreach (DataRow row in result.Rows)
                {
                    var fields = row.ItemArray.Select(field => EscapeCsvField(field.ToString()));
                    writer.WriteLine(string.Join(",", fields));
                }
            }
        }

    }
}
