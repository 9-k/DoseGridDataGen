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
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;
using System.Windows;


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
                using (VMS.TPS.Common.Model.API.Application app = VMS.TPS.Common.Model.API.Application.CreateApplication())
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }
        }
        static void Execute(VMS.TPS.Common.Model.API.Application app)
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
                "Target Dose Coverage",
                "Target Sampling Coverage",
                // oar stuff
                "Actual OAR Name",
                "TG263 OAR Name Guess",
                "OAR Volume",
                "OAR Mean Dose",
                "OAR Hot Spot",
                "OAR D0.03",
                "OAR Dose Coverage",
                "OAR Sampling Coverage",
                // misc
                "Minimum Separation",
                "Centroid Separation",
                "Grid Size",
                "MRN",
                "Course Name",
                "Plan Name",
                "Course Dx",
                "Calculation Model",
                "Calculation Time"
            };

            ESAPIUtility.AutoClickOk();
            Console.WriteLine("AutoOK set up.");

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
            string pattern = @"(?<=^|,)(\""(?:[^\""]|\""\"")*\""|[^,]*)";

            using (var reader = new StreamReader(TG263Path))
            {
                bool isHeader = true;
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (line == null) continue;

                    var matches = Regex.Matches(line, pattern);
                    var fields = new string[matches.Count];

                    for (int i = 0; i < matches.Count; i++)
                    {
                        fields[i] = matches[i].Value.Trim('"');  // Remove surrounding quotes if present
                    }

                    if (isHeader)
                    {
                        // Add columns based on the header row
                        foreach (var header in fields)
                        {
                            TG263Table.Columns.Add(header);
                        }
                        isHeader = false;
                    }
                    else
                    {
                        // Add rows for each data line
                        TG263Table.Rows.Add(fields);
                    }
                }
            }
#if DEBUG
            foreach (DataColumn col in TG263Table.Columns)
            {
                Console.WriteLine(col.ColumnName);
            }
#endif

            List<string> mrnList = File.ReadAllLines(mrnPath).ToList();

            foreach (string mrn in mrnList)
            {
                Stopwatch patientWatch = Stopwatch.StartNew();
                // copying between patients is not possible.
                // I'm just going to do everything in situ then NOT save any modifications...
                // so it all gets rolled back.
                // new result table for each patient, so that if it crashes or a patient
                // fails I can just merge the tables later
                // rather than losing all my data
                DataTable result = new DataTable();
                foreach (string columnName in TRACKEDPARAMS) { result.Columns.Add(columnName); }
                Patient patient = null;
                Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                Console.WriteLine("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~");
                try 
                { 
                    patient = app.OpenPatientById(mrn);
                    Console.WriteLine($"Successfully opened patient {mrn}.");
                }
                catch
                {
                    Console.WriteLine($"Failed to open patient {mrn}. Continuing:");
                    continue;
                }
                
                patient.BeginModifications();
                
#if DEBUG       // just do all courses and don't worry about filtering for matching to icd list 
                List<Course> coursesMatchingDxList = patient.Courses.ToList();
#else 
                List<Course> coursesMatchingDxList =
                    patient.Courses
                    .Where(course => course.Diagnoses
                    .Any(dx => ICDCODELIST.Any(code => dx.Code.Contains(code)))).ToList(); 
#endif

                // now I don't know why this would ping but just in case:
                if (coursesMatchingDxList.Count <= 0) 
                {
                    Console.WriteLine("For some reason this patient is in the list but has no courses with matching diagnoses. Skipping:");
                    continue; 
                }

                Course GridProj = patient.AddCourse();
                GridProj.Id = projectCourseID;

                foreach (Course course in coursesMatchingDxList)
                {
                    Console.WriteLine("================================");
                    Console.WriteLine($"Working on course {course.Id}!");
#if DEBUG
                    string CourseDx = "DEBUG!";
#else
                    if (course.Id.ToUpper().Contains("DNU") || course.Id.ToUpper().Contains("QA") || course.Id.ToUpper().Contains("TEST")) { continue; }
                    string CourseDx = course.Diagnoses.Where(dx => ICDCODELIST.Any(code => dx.Code.Contains(code))).First().Code;          
#endif
                    if (course.ExternalPlanSetups.Count() < 1)
                    {
                        Console.WriteLine("Weirdly no plans in the course... skipping.");
                        continue;
                    }

                    foreach (ExternalPlanSetup originalPlan in course.ExternalPlanSetups)
                    {
                        Stopwatch planwatch = Stopwatch.StartNew();
#if !DEBUG
                        //if (!originalPlan.IsTreated) 
                        //{
                        //    Console.WriteLine("Looks like this plan wasn't treated - skipping.");
                        //    continue; 
                        //}
#endif
                        Console.WriteLine("-----------------");
                        Console.WriteLine($"Working on plan {originalPlan.Id}");
                        ExternalPlanSetup plan = GridProj.CopyPlanSetup(originalPlan) as ExternalPlanSetup;
                        string ActualTargetName = plan.TargetVolumeID;
                        Structure targetStructure = plan.StructureSet.Structures.Where(structure => structure.Id == ActualTargetName).FirstOrDefault();

                        string TG263TargetNameGuess = StructureParserMethods.StructureParser(
                            ActualTargetName, 
                            TG263Table, 
                            cutoff: PARSEMATCHCUTOFF)
                            .matches.First().Item1["TG263-Primary Name"] as string;

                        foreach (string gridSize in DOSEGRIDS)
                        {
                            Console.WriteLine("..........");
                            Console.WriteLine($"Working on grid size {gridSize}!");
                            foreach (string calcModel in new List<string>() { "AcurosXB_156MR3", "AAA_15606" })
                            {
                                Console.WriteLine("...");
                                Console.WriteLine($"Setting parameters for calc model {calcModel}!");
                                plan.SetCalculationModel(CalculationType.PhotonVolumeDose, calcModel);

                                bool didSetGrid = plan.SetCalculationOption(plan.PhotonCalculationModel, "CalculationGridSizeInCM", gridSize);
                                bool didSetSRSGrid = plan.SetCalculationOption(plan.PhotonCalculationModel, "CalculationGridSizeInCMForSRSAndHyperArc", gridSize);
                                
                                if (!(didSetGrid && didSetSRSGrid))
                                {
                                    Console.WriteLine("For some reason the script failed to set the gridsize. Aborting:");
                                    throw new Exception();
                                }
                                bool gpuOn = plan.SetCalculationOption(plan.PhotonCalculationModel, "UseGPU", "Yes");
                                // it's fine if the gpu isn't set lol so i'm not checking it
                                double calcTime = 0;
                                try
                                {
                                    Console.WriteLine($"Trying to calc {calcModel} at DG {gridSize}");
                                    Stopwatch stopwatch = Stopwatch.StartNew();
                                    CalculationResult calcResult = plan.CalculateDose();
                                    stopwatch.Stop();
                                    calcTime = stopwatch.Elapsed.TotalSeconds;
                                    Console.WriteLine($"Calculation complete in {calcTime} seconds.");
                                    if (!calcResult.Success)
                                    {
                                        Console.WriteLine("Calculation failed and I'm not digging through the logs to figure out why.");
                                        Console.WriteLine("Retrying with gpu off. If this doesn't work I'm skipping.");
                                        bool gpuOff = plan.SetCalculationOption(plan.PhotonCalculationModel, "UseGPU", "No");
                                        if (!gpuOff)
                                        {
                                            Console.WriteLine("Well, I couldn't figure out how to undo the GPU. Skipping:");
                                            continue;
                                        }
                                        CalculationResult backupCalcResult = plan.CalculateDose();
                                        if (!backupCalcResult.Success)
                                        {
                                            Console.WriteLine("Well, it failed again. Skipping to the next dosegrid:");
                                            continue;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Something went wrong calcing {calcModel} at DG {gridSize}. Skipping:");
                                    Console.WriteLine(ex.ToString());
                                    continue;
                                }
                                

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

                                double TargetDoseCoverage = plan.GetDVHCumulativeData(
                                        targetStructure,
                                        DoseValuePresentation.Absolute,
                                        VolumePresentation.Relative,
                                        1).Coverage;

                                double TargetSamplingCoverage = plan.GetDVHCumulativeData(
                                    targetStructure,
                                    DoseValuePresentation.Absolute,
                                    VolumePresentation.Relative,
                                    1).SamplingCoverage;

                                int structurecounter = 1;
                                foreach (Structure oar in plan.StructureSet.Structures)
                                {
                                    Console.WriteLine($"Working on structure {oar.Id} which is {structurecounter.ToString()}/{plan.StructureSet.Structures.Count().ToString()}");
                                    structurecounter++;
                                    if (oar.Id.ToUpper() == "BODY") { Console.WriteLine("BODY detected; skipping"); continue; }
                                    if (oar.Id == ActualTargetName) { Console.WriteLine("This is the target; skipping."); continue; }
                                    if (oar.Id.ToUpper().Equals("BRAIN")) { Console.WriteLine("This is brain; skipping."); continue; }
                                    if (oar.Id.ToUpper().Contains("RING")) { Console.WriteLine("This is ring; skipping."); continue; }
                                    if (oar.Id.ToUpper().Contains("PREV")) { Console.WriteLine("This is previous; skipping."); continue; }
                                    if (oar.Id.ToUpper().Contains("EXP")) { Console.WriteLine("This is exp; skipping."); continue; }
                                    if (oar.Id.ToUpper().Contains("GTV")) { Console.WriteLine("This is gtv; skipping."); continue; }
                                    if (oar.Id.ToUpper().Contains("CTV")) { Console.WriteLine("This is ctv; skipping."); continue; }
                                    if (oar.Id.ToUpper().Contains("ITV")) { Console.WriteLine("This is itv; skipping."); continue; }

                                    string ActualOARName = oar.Id;
                                    TG263ParseResult TG263OARParse = StructureParserMethods.StructureParser(
                                        ActualOARName,
                                        TG263Table,
                                        cutoff: PARSEMATCHCUTOFF);

                                    bool dontLikeFlag = false;
                                    if (TG263OARParse.IsEval) { Console.WriteLine("IsEval"); dontLikeFlag = true; }
                                    if (TG263OARParse.IsPRV) { Console.WriteLine("IsPRV"); dontLikeFlag = true; }
                                    if (TG263OARParse.IsOpti) { Console.WriteLine("IsOpti"); dontLikeFlag = true; }
                                    if (TG263OARParse.IsDerived) { Console.WriteLine("IsDerived"); dontLikeFlag = true; }
                                    if (TG263OARParse.IsPlanning) { Console.WriteLine("IsPlanning"); dontLikeFlag = true; }
                                    if (TG263OARParse.IsCouch) { Console.WriteLine("IsCouch"); dontLikeFlag= true; }
                                    if (TG263OARParse.matches.Count < 1) { Console.WriteLine("No match?"); dontLikeFlag = true; }
                                    if (dontLikeFlag)
                                    {
                                        Console.WriteLine($"I don't like the structure {oar.Id}, because of one of the above reasons. Skipping it.");
                                        continue; 
                                    }
                                    string TG263OARNameGuess = TG263OARParse.matches.First().Item1["TG263-Primary Name"] as string;
                                    Console.WriteLine($"Matching {oar.Id} to {TG263OARNameGuess} with similarity {TG263OARParse.matches.First().Item2.ToString()}");
                                    double OARVolume = oar.Volume;
                                    Console.WriteLine("Got volume.");

                                    double OARMeanDose = plan.GetDVHCumulativeData(
                                        oar,
                                        DoseValuePresentation.Absolute,
                                        VolumePresentation.Relative,
                                        1).MeanDose.Dose;
                                    Console.WriteLine("got mean dose");


                                    double OARHotSpot = plan.GetDVHCumulativeData(
                                        oar,
                                        DoseValuePresentation.Absolute,
                                        VolumePresentation.Relative,
                                        1).MaxDose.Dose;
                                    Console.WriteLine("got hotspot");


                                    double OARD003 = 0.0;
                                    var OARDVHDataAbsVol = plan.GetDVHCumulativeData(
                                        oar,
                                        DoseValuePresentation.Absolute,
                                        VolumePresentation.AbsoluteCm3,
                                        0.01);
                                    foreach (var v in OARDVHDataAbsVol.CurveData )
                                    {
                                        if (v.Volume <= 0.03)
                                        {
                                            OARD003 = v.DoseValue.Dose;
                                            break;
                                        }
                                    }

                                    //double OARD003 = plan.GetDoseAtVolume(oar, 0.03, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose;
                                    Console.WriteLine("got d003");


                                    double OARDoseCoverage = plan.GetDVHCumulativeData(
                                        oar,
                                        DoseValuePresentation.Absolute,
                                        VolumePresentation.Relative,
                                        1).Coverage;
                                    Console.WriteLine("got dose coverage");

                                    double OARSamplingCoverage = plan.GetDVHCumulativeData(
                                        oar,
                                        DoseValuePresentation.Absolute,
                                        VolumePresentation.Relative,
                                        1).SamplingCoverage;
                                    Console.WriteLine("got sampling coverage");

                                    double MinimumSeparation = ESAPIUtility.MinimumStructureDistance(targetStructure, oar);
                                    Console.WriteLine("got minsep");
                                    double CentroidSeparation = (targetStructure.CenterPoint - oar.CenterPoint).Length;
                                    Console.WriteLine("got centroid sep");
                                    DataRow newRow = result.NewRow();
                                    newRow["Actual Target Name"] = ActualTargetName;
                                    newRow["TG263 Target Name Guess"] = TG263TargetNameGuess;
                                    newRow["Target Volume"] = TargetVolume.ToString();
                                    newRow["Target Mean Dose"] = TargetMeanDose.ToString();
                                    newRow["Target Hot Spot"] = TargetHotSpot.ToString();
                                    newRow["Target V95"] = TargetV95.ToString();
                                    newRow["Target Dose Coverage"] = TargetDoseCoverage.ToString();
                                    newRow["Target Sampling Coverage"] = TargetSamplingCoverage.ToString();
                                    // oar stuff
                                    newRow["Actual OAR Name"] = ActualOARName;
                                    newRow["TG263 OAR Name Guess"] = TG263OARNameGuess;
                                    newRow["OAR Volume"] = OARVolume.ToString();
                                    newRow["OAR Mean Dose"] = OARMeanDose.ToString();
                                    newRow["OAR Hot Spot"] = TargetHotSpot.ToString();
                                    newRow["OAR D0.03"] = OARD003.ToString();
                                    newRow["OAR Dose Coverage"] = OARDoseCoverage.ToString();
                                    newRow["OAR Sampling Coverage"] = OARSamplingCoverage.ToString();
                                    // misc
                                    newRow["Minimum Separation"] = MinimumSeparation.ToString();
                                    newRow["Centroid Separation"] = CentroidSeparation.ToString();
                                    newRow["Grid Size"] = gridSize;
                                    newRow["MRN"] = mrn.ToString();
                                    newRow["Course Name"] = course.Id;
                                    newRow["Plan Name"] = originalPlan.Id; 
                                    newRow["Course Dx"] = CourseDx;
                                    newRow["Calculation Model"] = calcModel;
                                    newRow["Calculation Time"] = calcTime;
                                    try
                                    {
                                        result.Rows.Add(newRow);
                                        Console.WriteLine($"Successfully added row about oar {oar.Id}.");
                                    }
                                    catch
                                    {
                                        Console.WriteLine($"Well, I couldn't write this row for OAR {oar.Id}. Skipping?");
                                    }
                                }
                            }
                        }
                        planwatch.Stop();
                        double planTime = planwatch.Elapsed.TotalSeconds;
                        Console.WriteLine($"Plan analysis complete in {planTime} seconds.");
                    }
                }

                string EscapeCsvField(string field)
                {
                    if (field.Contains(",") || field.Contains("\"") || field.Contains("\n"))
                    {
                        return "\"" + field.Replace("\"", "\"\"") + "\"";
                    }
                    return field;
                }


                using (StreamWriter writer = new StreamWriter($"{mrn}.csv", false, Encoding.UTF8))
                {
                    var columnNames = result.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                    writer.WriteLine(string.Join(",", columnNames));

                    // Write each row
                    foreach (DataRow row in result.Rows)
                    {
                        var fields = row.ItemArray.Select(field => EscapeCsvField(field.ToString()));
                        writer.WriteLine(string.Join(",", fields));
                    }
                }
                Console.WriteLine($"Wrote {mrn} analysis to csv.");
                patientWatch.Stop();
                double patientTime = patientWatch.Elapsed.TotalSeconds;
                Console.WriteLine($"Analysis complete in {patientTime} seconds.");
                app.ClosePatient();
            }
            Console.WriteLine("All Done!");
            Console.ReadLine();
        }
    }
}
