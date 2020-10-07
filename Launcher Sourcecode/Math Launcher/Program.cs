using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;

namespace Portable_Libre_Math_Launcher
{
    class Program
    {
        static void Main()
        {
            string applicationPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            var sb = new System.Text.StringBuilder();
            string[] CommandLineArgs = Environment.GetCommandLineArgs();
            for (int i = 1; i < CommandLineArgs.Length; i++)
            {
                sb.Append(" \"" + CommandLineArgs[i] + "\"");
            }
            Process.Start(applicationPath + "\\Libre Office\\program\\smath.exe", sb.ToString());
        }
    }
}
