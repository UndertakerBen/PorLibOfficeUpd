using System;
using System.IO;

namespace Portable_Libre_Office.Files
{
    public partial class Registrymodifications
    {
        public static void CreateRegistrymodification_XCU(string applicationPath)
        {
            if (!Directory.Exists(applicationPath + "\\UserData\\user"))
            {
                Directory.CreateDirectory(applicationPath + "\\UserData\\user");
                if (!File.Exists(applicationPath + "\\UserData\\user\\registrymodifications.xcu"))
                {
                    File.WriteAllText(applicationPath + "\\UserData\\user\\registrymodifications.xcu", "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + Environment.NewLine +
                    "<oor:items xmlns:oor=\"http://openoffice.org/2001/registry\" xmlns:xs=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">" + Environment.NewLine +
                    "<item oor:path=\"/org.openoffice.Office.Common/Misc\"><prop oor:name=\"FirstRun\" oor:op=\"fuse\"><value>false</value></prop></item>" + Environment.NewLine +
                    "<item oor:path=\"/org.openoffice.Office.Common/Misc\"><prop oor:name=\"ShowTipOfTheDay\" oor:op=\"fuse\"><value>false</value></prop></item>" + Environment.NewLine +
                    "<item oor:path=\"/org.openoffice.Setup/Office\"><prop oor:name=\"ooSetupInstCompleted\" oor:op=\"fuse\"><value>true</value></prop></item>" + Environment.NewLine +
                    "<item oor:path=\"/org.openoffice.Setup/Product\"><prop oor:name=\"ooSetupLastVersion\" oor:op=\"fuse\"><value>10.0</value></prop></item>" + Environment.NewLine +
                    "</oor:items>");
                }
            }
        }
    }
}
