using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

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
            string settings = "settings.txt";
            string assyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string settingsPath = Path.Combine(assyPath, settings);

            Dictionary<string, string> settingsDict = new Dictionary<string, string>();

            foreach (var line in File.ReadLines(settingsPath))
            {
                var parts = line.Split(new[] { ':' }, 2);
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim();
                    string value = parts[1].Trim();
                    settingsDict[key] = value;
                }
                else
                {
                    throw new ApplicationException("Malformed settings.txt file. Ensure each line of of the form *:*?");
                }
            }

            string mrnPath = Path.Combine(assyPath, settingsDict["mrnFile"]);
            string tg263Path = Path.Combine(assyPath, settingsDict["tg263File"]);
            string dummyPtId = settingsDict["dummyPtId"];

            IEnumerable<string> mrnList = File.ReadLines(mrnPath)
                .Select(line => line.Trim())
                .Where(line => !string.IsNullOrEmpty(line));


            int courseCounter = 0;
            int patientCounter = 0;
            foreach (string mrn in mrnList)
            {
                Patient patient = app.OpenPatientById(mrn);
                


                app.ClosePatient();
            }
        }
    }
}
