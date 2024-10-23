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
// [assembly: ESAPIScript(IsWriteable = true)]

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
            // you absolutely need to ensure this list matches with the database crawl list
            List<string> ICDCODELIST = new List<string>()
                {
                    "C44",
                    "C69",
                    "C70",
                    "C71",
                    "C72",
                    "C75.1",
                    "C76",
                    "C79.3",
                    "C96",
                };

            List<string> TRACKEDOARS = new List<string>()
            {

            };

            List<string> TRACKEDPARAMS = new List<string>()
            {
                "Target FMAID",
                "OAR FMAID",
                "Target Volume",
                "OAR Volume",
                "Minimum Separation",
                "Centroid Separation",
                "Prescription Dose To Target",
                "True Dose To Target",
                "Grid Size", 
                "MRN",
                "Course Name",
                "Plan Name"
            };

            int CUTOFF = 90;
            string settingsPath = "settings.txt";
            string assyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            Dictionary <string,string> settingsDict = ESAPIUtility.SettingsDict(settingsPath);

            string mrnPath = Path.GetFullPath(settingsDict["mrnPath"]);
            string TG263Path = Path.GetFullPath(settingsDict["TG263Path"]);
            string dummyPtId = settingsDict["dummyPtId"];

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

            int courseCounter = 0;
            int patientCounter = 0;
            foreach (string mrn in mrnList)
            {
                Patient patient = app.OpenPatientById(mrn);
                List<Course> coursesMatchingDxList =
                    patient.Courses
                    .Where(course => course.Diagnoses
                    .Any(dx => ICDCODELIST.Any(code => dx.Code.Contains(code)))).ToList();

                foreach (Course course in coursesMatchingDxList)
                {
                    foreach (ExternalPlanSetup plan in course.ExternalPlanSetups)
                    {
                        if (!plan.IsTreated) { continue; }
                        var target = plan.TargetVolumeID;
                        TG263ParseResult targetParse = StructureParserMethods.StructureParser(target, TG263Table, cutoff: CUTOFF);
                    }
                    courseCounter++;
                }
                app.ClosePatient();
            }
        }

    }
}
