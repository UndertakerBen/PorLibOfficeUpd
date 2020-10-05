using System;
using System.IO;
using System.Windows.Forms;

namespace Portable_Libre_Office.Files
{
    public partial class Regfile
    {
        private static readonly string[,] impress = new string[10, 3] { { "Cgm", "7", "Computer Graphics Metafile" }, { "Pot", "8", "Microsoft PowerPoint 97–2003 Dokumentvorlage" }, { "Pps", "7", "Microsoft PowerPoint Bildschirmpräsentation" }, { "Ppt", "7", "Microsoft PowerPoint 97–2003 Präsentation" }, { "Potm", "8", "Microsoft PowerPoint Dokumentvorlage" }, { "Potx", "8", "Microsoft PowerPoint Dokumentvorlage" }, { "Ppsx", "7", "Microsoft PowerPoint Bildschirmpräsentation" }, { "Pptm", "7", "Microsoft PowerPoint Präsentation" }, { "Pptx", "7", "Microsoft PowerPoint Präsentation" }, { "Uop", "7", "Uniform Office Format Präsentation" } };
        private static readonly string[,] writer = new string[10, 3] { { "Doc", "1", "Microsoft Word 97–2003 Dokument" }, { "Docx", "1", "Microsoft Word Dokument" }, { "Dot", "2", "Microsoft Word 97–2003 Dokumentvorlage" }, { "Rtf", "1", "Rich Text Dokument" }, { "602", "1", "T602-Textdatei" }, { "Docm", "1", "Microsoft Word Dokument" }, { "Dotm", "1", "Microsoft Word Dokumentvorlage" }, { "Dotx", "2", "Microsoft Word Dokumentvorlage" }, { "Lwp", "1", "Lotus Word Pro Dokument" }, { "Uot", "1", "Uniform Office Format Textdokument" } };
        private static readonly string[,] calc = new string[9, 3] { { "Xls", "3", "Microsoft Excel 97–2003 Arbeitsblatt" }, { "Xlt", "4", "Microsoft Excel 97–2003 Dokumentvorlage" }, { "Iqy", "0", "Microsoft Excel Web Query-Datei" }, { "Wb2", "3", "Lotus Quattro Pro Tabellendokument" }, { "Xlsb", "3", "Microsoft Excel Arbeitsblatt" }, { "Xlsm", "3", "Microsoft Excel Arbeitsblatt" }, { "Xlsx", "3", "Microsoft Excel Arbeitsblatt" }, { "Xltm", "4", "Microsoft Excel Dokumentvorlage" }, { "Xltx", "4", "Microsoft Excel Dokumentvorlage" } };
        private static readonly string[] scalc = new string[12] { "csv", "htm", "html", "xml", "123", "dbf", "dif", "slk", "sxc", "wk1", "wks", "xlw" };
        private static readonly string[] swrite = new string[6] { "htm", "html", "xml", "hwp", "sxw", "wpd" };
        private static readonly string[] progName = new string[7] { "Base", "Calc", "Draw", "Impress", "Math", "Office", "Writer" };
        private static readonly string[,] lwriter = new string[6, 3] { { "Odt", "1", "OpenDocument Text" }, { "Odm", "9", "OpenDocument Globaldokument" }, { "Oth", "10", "HTML-Dokument Dokumentvorlage" }, { "Ott", "2", "OpenDocument Textdokument Dokumentvorlage" }, { "Stw", "2", "OpenOffice.org 1.1 Textdokument Dokumentvorlage" }, { "Sxg", "9", "OpenOffice.org 1.1 Globaldokument" } };
        private static readonly string[,] ldraw = new string[5, 3] { { "Odg", "5", "OpenDocument Zeichnung"}, {"Fodg","5","OpenDocument Zeichnung" }, { "Otg", "6", "OpenDocument Zeichnung Dokumentvorlage" }, { "Std", "6", "OpenOffice.org 1.1 Zeichnung Dokumentvorlage" }, { "Sxd", "5", "OpenOffice.org 1.1 Zeichnung" } };
        private static readonly string[,] limpress = new string[5, 3] { { "Odp", "7", "OpenDocument Zeichnung Dokumentvorlage" }, { "Fodp", "7", "OpenDocument Präsentation" }, { "Otp", "8", "OpenDocument Zeichnung Dokumentvorlage" }, { "Sti", "8", "OpenOffice.org 1.1 Präsentation Dokumentvorlage" }, { "Sxi", "7", "OpenOffice.org 1.1 Präsentation" } };
        private static readonly string[,] lcalc = new string[4, 3] { { "Fods", "3", "OpenDocument Tabellendokument" }, { "Ods", "3", "OpenDocument Tabellendokument" }, { "Ots", "4", "OpenDocument Tabellendokument Dokumentvorlage" }, { "Stc", "4", "OpenOffice.org 1.1 Tabellendokument Dokumentvorlage" } };
        private static readonly string[,] lmath = new string[3, 3] { { "Mml", "12", "OpenOffice.org 1.1 Formel" }, { "Sxm", "12", "OpenOffice.org 1.1 Formel" }, { "Odf", "12", "OpenDocument Formel" } };
        private static Microsoft.Win32.RegistryKey key;
        public static void RegCreate(string applicationPath)
        {
            for (int i = 0; i < impress.GetLength(0); i++)
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + impress[i, 0].ToLower());
                key.SetValue(default, "LibreOffice" + impress[i, 0] + ".PORTABLE");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + impress[i, 0].ToLower() + "\\OpenWithProgids");
                key.SetValue("LibreOffice" + impress[i, 0] + ".PORTABLE", " ");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + impress[i, 0] + ".PORTABLE");
                key.SetValue(default, impress[i, 2]);
                key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Impress");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + impress[i, 0] + ".PORTABLE\\DefaultIcon");
                key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin," + impress[i, 1]);
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + impress[i, 0] + ".PORTABLE\\Shell");
                key.SetValue(default, "open");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + impress[i, 0] + ".PORTABLE\\Shell\\open\\command");
                key.SetValue(default, "\"" + applicationPath + "\\Libre Impress Portable Launcher.exe\" -o \"%1\"");
                key.Close();
            }
            for (int i = 0; i < writer.GetLength(0); i++)
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + writer[i, 0].ToLower());
                key.SetValue(default, "LibreOffice" + writer[i, 0] + ".PORTABLE");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + writer[i, 0].ToLower() + "\\OpenWithProgids");
                key.SetValue("LibreOffice" + writer[i, 0] + ".PORTABLE", " ");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + writer[i, 0] + ".PORTABLE");
                key.SetValue(default, writer[i, 2]);
                key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Writer");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + writer[i, 0] + ".PORTABLE\\DefaultIcon");
                key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin," + writer[i, 1]);
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + writer[i, 0] + ".PORTABLE\\Shell");
                key.SetValue(default, "open");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + writer[i, 0] + ".PORTABLE\\Shell\\open\\command");
                key.SetValue(default, "\"" + applicationPath + "\\Libre Writer Portable Launcher.exe\" -o \"%1\"");
                key.Close();
            }
            for (int i = 0; i < calc.GetLength(0); i++)
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + calc[i, 0].ToLower());
                key.SetValue(default, "LibreOffice" + calc[i, 0] + ".PORTABLE");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + calc[i, 0].ToLower() + "\\OpenWithProgids");
                key.SetValue("LibreOffice" + calc[i, 0] + ".PORTABLE", " ");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + calc[i, 0] + ".PORTABLE");
                key.SetValue(default, calc[i, 2]);
                key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Calc");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + calc[i, 0] + ".PORTABLE\\DefaultIcon");
                key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin," + calc[i, 1]);
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + calc[i, 0] + ".PORTABLE\\Shell");
                key.SetValue(default, "open");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + calc[i, 0] + ".PORTABLE\\Shell\\open\\command");
                key.SetValue(default, "\"" + applicationPath + "\\Libre Calc Portable Launcher.exe\" -o \"%1\"");
                key.Close();
            }
            for (int i = 0; i < ldraw.GetLength(0); i++)
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + ldraw[i, 0].ToLower());
                key.SetValue(default, "LibreOffice" + ldraw[i, 0] + ".PORTABLE");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + ldraw[i, 0].ToLower() + "\\OpenWithProgids");
                key.SetValue("LibreOffice" + ldraw[i, 0] + ".PORTABLE", " ");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + ldraw[i, 0] + ".PORTABLE");
                key.SetValue(default, ldraw[i, 2]);
                key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Draw");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + ldraw[i, 0] + ".PORTABLE\\DefaultIcon");
                key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin," + ldraw[i, 1]);
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + ldraw[i, 0] + ".PORTABLE\\Shell");
                key.SetValue(default, "open");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + ldraw[i, 0] + ".PORTABLE\\Shell\\open\\command");
                key.SetValue(default, "\"" + applicationPath + "\\Libre Draw Portable Launcher.exe\" -o \"%1\"");
                key.Close();
            }
            for (int i = 0; i < limpress.GetLength(0); i++)
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + limpress[i, 0].ToLower());
                key.SetValue(default, "LibreOffice" + limpress[i, 0] + ".PORTABLE");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + limpress[i, 0].ToLower() + "\\OpenWithProgids");
                key.SetValue("LibreOffice" + limpress[i, 0] + ".PORTABLE", " ");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + limpress[i, 0] + ".PORTABLE");
                key.SetValue(default, limpress[i, 2]);
                key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Impress");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + limpress[i, 0] + ".PORTABLE\\DefaultIcon");
                key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin," + limpress[i, 1]);
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + limpress[i, 0] + ".PORTABLE\\Shell");
                key.SetValue(default, "open");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + limpress[i, 0] + ".PORTABLE\\Shell\\open\\command");
                key.SetValue(default, "\"" + applicationPath + "\\Libre Impress Portable Launcher.exe\" -o \"%1\"");
                key.Close();
            }
            for (int i = 0; i < lwriter.GetLength(0); i++)
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + lwriter[i, 0].ToLower());
                key.SetValue(default, "LibreOffice" + lwriter[i, 0] + ".PORTABLE");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + lwriter[i, 0].ToLower() + "\\OpenWithProgids");
                key.SetValue("LibreOffice" + lwriter[i, 0] + ".PORTABLE", " ");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lwriter[i, 0] + ".PORTABLE");
                key.SetValue(default, lwriter[i, 2]);
                key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Writer");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lwriter[i, 0] + ".PORTABLE\\DefaultIcon");
                key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin," + lwriter[i, 1]);
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lwriter[i, 0] + ".PORTABLE\\Shell");
                key.SetValue(default, "open");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lwriter[i, 0] + ".PORTABLE\\Shell\\open\\command");
                key.SetValue(default, "\"" + applicationPath + "\\Libre Writer Portable Launcher.exe\" -o \"%1\"");
                key.Close();
            }
            for (int i = 0; i < lcalc.GetLength(0); i++)
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + lcalc[i, 0].ToLower());
                key.SetValue(default, "LibreOffice" + lcalc[i, 0] + ".PORTABLE");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + lcalc[i, 0].ToLower() + "\\OpenWithProgids");
                key.SetValue("LibreOffice" + lcalc[i, 0] + ".PORTABLE", " ");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lcalc[i, 0] + ".PORTABLE");
                key.SetValue(default, lcalc[i, 2]);
                key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Calc");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lcalc[i, 0] + ".PORTABLE\\DefaultIcon");
                key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin," + lcalc[i, 1]);
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lcalc[i, 0] + ".PORTABLE\\Shell");
                key.SetValue(default, "open");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lcalc[i, 0] + ".PORTABLE\\Shell\\open\\command");
                key.SetValue(default, "\"" + applicationPath + "\\Libre Calc Portable Launcher.exe\" -o \"%1\"");
                key.Close();
            }
            if (File.Exists(applicationPath + "\\Libre Office\\program\\smath.exe"))
            {
                for (int i = 0; i < lmath.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + lmath[i, 0].ToLower());
                    key.SetValue(default, "LibreOffice" + lmath[i, 0] + ".PORTABLE");
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + lmath[i, 0].ToLower() + "\\OpenWithProgids");
                    key.SetValue("LibreOffice" + lmath[i, 0] + ".PORTABLE", " ");
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lmath[i, 0] + ".PORTABLE");
                    key.SetValue(default, lmath[i, 2]);
                    key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Math");
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lmath[i, 0] + ".PORTABLE\\DefaultIcon");
                    key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin," + lmath[i, 1]);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lmath[i, 0] + ".PORTABLE\\Shell");
                    key.SetValue(default, "open");
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOffice" + lmath[i, 0] + ".PORTABLE\\Shell\\open\\command");
                    key.SetValue(default, "\"" + applicationPath + "\\Libre Math Portable Launcher.exe\" -o \"%1\"");
                    key.Close();
                    key.Close();
                }
            }
            for (int i = 0; i < scalc.GetLength(0); i++)
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + scalc[i]);
                key.SetValue(default, "StarCalc.PORTABLE");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + scalc[i] + "\\OpenWithProgids");
                key.SetValue("StarCalc.PORTABLE", " ");
                key.Close();
            }
            if (File.Exists(applicationPath + "\\Libre Office\\program\\sbase.exe"))
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\.odb");
                key.SetValue(default, "LibreOfficeOdb.PORTABLE");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\.odb\\OpenWithProgids");
                key.SetValue("LibreOfficeOdb.PORTABLE", " ");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOfficeOdb.PORTABLE");
                key.SetValue(default, "OpenDocument Datenbank");
                key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Base");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOfficeOdb.PORTABLE\\DefaultIcon");
                key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin,11");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOfficeOdb.PORTABLE\\Shell");
                key.SetValue(default, "open");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\LibreOfficeOdb.PORTABLE\\Shell\\open\\command");
                key.SetValue(default, "\"" + applicationPath + "\\Libre Base Portable Launcher.exe\" -o \"%1\"");
                key.Close();
                key.Close();
            }
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\StarCalc.PORTABLE");
            key.SetValue(default, "OpenOffice.org 1.1 Tabellendokument");
            key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Calc");
            key.Close();
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\StarCalc.PORTABLE\\DefaultIcon");
            key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin,3");
            key.Close();
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\StarCalc.PORTABLE\\Shell");
            key.SetValue(default, "open");
            key.Close();
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\StarCalc.PORTABLE\\Shell\\open\\command");
            key.SetValue(default, "\"" + applicationPath + "\\Libre Calc Portable Launcher.exe\" -o \"%1\"");
            key.Close();
            key.Close();
            for (int i = 0; i < swrite.GetLength(0); i++)
            {
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + swrite[i] + "\\OpenWithProgids");
                key.SetValue("StarWriter.PORTABLE", " ");
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\." + swrite[i]);
                key.SetValue(default, "StarWriter.PORTABLE");
                key.Close();
            }
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\StarWriter.PORTABLE");
            key.SetValue(default, "OpenOffice.org 1.1 Text Dokument");
            key.SetValue("AppUserModelID", "TheDocumentFoundation.LibreOffice.Writer");
            key.Close();
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\StarWriter.PORTABLE\\DefaultIcon");
            key.SetValue(default, applicationPath + "\\Libre Office\\program\\soffice.bin,1");
            key.Close();
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\StarWriter.PORTABLE\\Shell");
            key.SetValue(default, "open");
            key.Close();
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\StarWriter.PORTABLE\\Shell\\open\\command");
            key.SetValue(default, "\"" + applicationPath + "\\Libre Writer Portable Launcher.exe\" -o \"%1\"");
            key.Close();

            for (int i = 0; i < progName.GetLength(0); i++)
            {
                Method1(applicationPath, progName[i]);
            }
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\LibreOffice.PORTABLE\\Capabilities");
            key.SetValue("ApplicationDescription", "LibreOffice.Portable");
            key.SetValue("ApplicationIcon", applicationPath + "\\Libre Office\\program\\soffice.bin,0");
            key.SetValue("ApplicationName", "LibreOffice.Portable");
            key.Close();
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\LibreOffice.PORTABLE\\Capabilities\\FileAssociations");
            key.SetValue(".123", "StarCalc.Portable");
            key.SetValue(".602", "LibreOffice602.Portable");
            key.SetValue(".cgm", "LibreOfficeCgm.Portable");
            key.SetValue(".csv", "LibreOfficeCsv.Portable");
            key.SetValue(".dbf", "LibreOfficeDbf.Portable");
            key.SetValue(".dif", "LibreOfficeDif.Portable");
            key.SetValue(".doc", "LibreOfficeDoc.Portable");
            key.SetValue(".docm", "LibreOfficeDocm.Portable");
            key.SetValue(".docx", "LibreOfficeDocx.Portable");
            key.SetValue(".dot", "LibreOfficeDot.Portable");
            key.SetValue(".dotm", "LibreOfficeDotm.Portable");
            key.SetValue(".dotx", "LibreOfficeDotx.Portable");
            key.SetValue(".htm", "StarWriter.Portable");
            key.SetValue(".html", "StarWriter.Portable");
            key.SetValue(".hwp", "StarWriter.Portable");
            key.SetValue(".iqy", "LibreOfficeIqy.Portable");
            key.SetValue(".lwp", "LibreOfficeLwp.Portable");
            key.SetValue(".odm", "LibreOfficeOdm.Portable");
            key.SetValue(".otg", "LibreOfficeOtg.Portable");
            key.SetValue(".oth", "LibreOfficeOth.Portable");
            key.SetValue(".otp", "LibreOfficeOtp.Portable");
            key.SetValue(".ott", "LibreOfficeOtt.Portable");
            key.SetValue(".pot", "LibreOfficePot.Portable");
            key.SetValue(".potm", "LibreOfficePotm.Portable");
            key.SetValue(".potx", "LibreOfficePotx.Portable");
            key.SetValue(".pps", "LibreOfficePps.Portable");
            key.SetValue(".ppsx", "LibreOfficePpsx.Portable");
            key.SetValue(".ppt", "LibreOfficePpt.Portable");
            key.SetValue(".pptm", "LibreOfficePptm.Portable");
            key.SetValue(".pptx", "LibreOfficePptx.Portable");
            key.SetValue(".rtf", "LibreOfficeRtf.Portable");
            key.SetValue(".slk", "StarCalc.Portable");
            key.SetValue(".stc", "StarCalc.Portable");
            key.SetValue(".std", "StarCalc.Portable");
            key.SetValue(".sti", "StarCalc.Portable");
            key.SetValue(".stw", "StarWriter.Portable");
            key.SetValue(".sxg", "LibreOfficeSxg.Portable");
            key.SetValue(".uop", "LibreOfficeUop.Portable");
            key.SetValue(".uos", "LibreOfficeUos.Portable");
            key.SetValue(".uot", "LibreOfficeUot.Portable");
            key.SetValue(".wb2", "LibreOfficeWb2.Portable");
            key.SetValue(".wk1", "StarCalc.Portable");
            key.SetValue(".wks", "StarCalc.Portable");
            key.SetValue(".wpd", "StarWriter.Portable");
            key.SetValue(".xls", "LibreOfficeXls.Portable");
            key.SetValue(".xlsb", "LibreOfficeXlsb.Portable");
            key.SetValue(".xlsm", "LibreOfficeXlsm.Portable");
            key.SetValue(".xlsx", "LibreOfficeXlsx.Portable");
            key.SetValue(".xlt", "LibreOfficeXlt.Portable");
            key.SetValue(".xltm", "LibreOfficeXltm.Portable");
            key.SetValue(".xltx", "LibreOfficeXltx.Portable");
            key.SetValue(".xlw", "StarCalc.Portable");
            key.SetValue(".xml", "StarWriter.Portable");
            key.SetValue(".fodg", "LibreOfficeFodg.Portable");
            key.SetValue(".fodp", "LibreOfficeFodp.Portable");
            key.SetValue(".fods", "LibreOfficeFods.Portable");
            key.SetValue(".fodt", "LibreOfficeFodt.Portable");
            key.SetValue(".odg", "LibreOfficeOdg.Portable");
            key.SetValue(".odp", "LibreOfficeOdp.Portable");
            key.SetValue(".ods", "LibreOfficeOds.Portable");
            key.SetValue(".odt", "LibreOfficeOdt.Portable");
            key.SetValue(".sxc", "StarCalc.Portable");
            key.SetValue(".sxd", "LibreOfficeSxd.Portable");
            key.SetValue(".sxi", "LibreOfficeSxi.Portable");
            key.SetValue(".sxw", "StarWriter.Portable");
            if (File.Exists(applicationPath + "\\Libre Office\\program\\sbase.exe"))
            {
                key.SetValue(".odb", "LibreOfficeÓdb.Portable");
            }
            if (File.Exists(applicationPath + "\\Libre Office\\program\\smath.exe"))
            {
                key.SetValue(".mml", "LibreOfficeMml.Portable");
                key.SetValue(".odf", "LibreOfficeOdf.Portable");
                key.SetValue(".sxm", "LibreOfficeSxm.Portable");
            }
            key.Close();
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\RegisteredApplications");
            key.SetValue("LibreOffice.PORTABLE", @"SOFTWARE\LibreOffice.PORTABLE\Capabilities");
            key.Close();
            IconReset();
        }
        private static void Method1(string applicationPath, string progName)
        {
            key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\s" + progName.ToLower() + ".exe");
            key.SetValue(default, "\"" + applicationPath + "\"Libre " + progName + @" Portable Launcher.exe""");
            key.SetValue("Path", applicationPath);
            key.Close();
        }
        public static void RegDel(string applicationPath)
        {
            try
            {
                key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE", true);
                key.DeleteSubKeyTree("LibreOffice.PORTABLE", false);
                key.Close();
                key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes", true);
                for (int i = 0; i < impress.GetLength(0); i++)
                {
                    key.DeleteSubKeyTree("LibreOffice" + impress[i, 0] + ".PORTABLE", false);
                }
                for (int i = 0; i < writer.GetLength(0); i++)
                {
                    key.DeleteSubKeyTree("LibreOffice" + writer[i, 0] + ".PORTABLE", false);
                }
                for (int i = 0; i < calc.GetLength(0); i++)
                {
                    key.DeleteSubKeyTree("LibreOffice" + calc[i, 0] + ".PORTABLE", false);
                }
                for (int i = 0; i < ldraw.GetLength(0); i++)
                {
                    key.DeleteSubKeyTree("LibreOffice" + ldraw[i, 0] + ".PORTABLE", false);
                }
                for (int i = 0; i < limpress.GetLength(0); i++)
                {
                    key.DeleteSubKeyTree("LibreOffice" + limpress[i, 0] + ".PORTABLE", false);
                }
                for (int i = 0; i < lwriter.GetLength(0); i++)
                {
                    key.DeleteSubKeyTree("LibreOffice" + lwriter[i, 0] + ".PORTABLE", false);
                }
                for (int i = 0; i < lcalc.GetLength(0); i++)
                {
                    key.DeleteSubKeyTree("LibreOffice" + lcalc[i, 0] + ".PORTABLE", false);
                }
                if (File.Exists(applicationPath + "\\Libre Offcie\\program\\smath.exe"))
                {
                    for (int i = 0; i < lmath.GetLength(0); i++)
                    {
                        key.DeleteSubKeyTree("LibreOffice" + lmath[i, 0] + ".PORTABLE", false);
                    }
                }
                key.DeleteSubKeyTree("StarCalc.PORTABLE", false);
                key.DeleteSubKeyTree("StarWriter.PORTABLE", false);
                key.Close();
                if (File.Exists(applicationPath + "\\Libre Office\\program\\sbase.exe"))
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\.odb", true);
                    key.DeleteValue(default, false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\.odb\\OpenWithProgids", true);
                    key.DeleteValue("LibreOfficeOdb.PORTABLE", false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes", true);
                    key.DeleteSubKeyTree("LibreOfficeOdb.PPORTABLE", false);
                    key.Close();
                }
                for (int i = 0; i < impress.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + impress[i, 0].ToLower(), true);
                    key.DeleteValue(default, false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + impress[i, 0].ToLower() + "\\OpenWithProgids", true);
                    key.DeleteValue("LibreOffice" + impress[i, 0] + ".PORTABLE", false);
                    key.Close();
                }
                for (int i = 0; i < writer.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + writer[i, 0].ToLower(), true);
                    key.DeleteValue(default, false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + writer[i, 0].ToLower() + "\\OpenWithProgids", true);
                    key.DeleteValue("LibreOffice" + writer[i, 0] + ".PORTABLE", false);
                    key.Close();
                }
                for (int i = 0; i < calc.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + calc[i, 0].ToLower(), true);
                    key.DeleteValue(default, false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + calc[i, 0].ToLower() + "\\OpenWithProgids", true);
                    key.DeleteValue("LibreOffice" + calc[i, 0] + ".PORTABLE", false);
                    key.Close();
                }
                for (int i = 0; i < scalc.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + scalc[i] + "\\OpenWithProgids", true);
                    key.DeleteValue("StarCalc.PORTABLE", false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + scalc[i], true);
                    key.DeleteValue(default, false);
                    key.Close();
                }
                for (int i = 0; i < swrite.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + swrite[i] + "\\OpenWithProgids", true);
                    key.DeleteValue("StarWriter.PORTABLE", false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + swrite[i], true);
                    key.DeleteValue(default, false);
                    key.Close();
                }
                for (int i = 0; i < ldraw.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + ldraw[i, 0].ToLower() + "\\OpenWithProgids", true);
                    key.DeleteValue("LibreOffice" + ldraw[i, 0] + ".PORTABLE", false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + ldraw[i, 0].ToLower(), true);
                    key.DeleteValue(default, false);
                    key.Close();

                }
                for (int i = 0; i < limpress.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + limpress[i, 0].ToLower(), true);
                    key.DeleteValue(default, false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + limpress[i, 0].ToLower() + "\\OpenWithProgids", true);
                    key.DeleteValue("LibreOffice" + limpress[i, 0] + ".PORTABLE", false);
                    key.Close();
                }
                for (int i = 0; i < lwriter.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + lwriter[i, 0].ToLower(), true);
                    key.DeleteValue(default, false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + lwriter[i, 0].ToLower() + "\\OpenWithProgids", true);
                    key.DeleteValue("LibreOffice" + lwriter[i, 0] + ".PORTABLE", false);
                    key.Close();
                }
                for (int i = 0; i < lcalc.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + lcalc[i, 0].ToLower(), true);
                    key.DeleteValue(default, false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + lcalc[i, 0].ToLower() + "\\OpenWithProgids", true);
                    key.DeleteValue("LibreOffice" + lcalc[i, 0] + ".PORTABLE", false);
                    key.Close();
                }
                for (int i = 0; i < lmath.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + lmath[i, 0].ToLower(), true);
                    key.DeleteValue(default, false);
                    key.Close();
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Classes\\." + lmath[i, 0].ToLower() + "\\OpenWithProgids", true);
                    key.DeleteValue("LibreOffice" + lmath[i, 0] + ".PORTABLE", false);
                    key.Close();
                }
                for (int i = 0; i < progName.GetLength(0); i++)
                {
                    key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths", true);
                    key.DeleteSubKeyTree("s" + progName[i].ToLower() + ".exe", false);
                    key.Close();
                }
                key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("SOFTWARE\\RegisteredApplications", true);
                key.DeleteValue("LibreOffice.PORTABLE", false);
                key.Close();
                IconReset();
            }
            catch (Exception ex)
            {
                File.AppendAllText(applicationPath + "\\Message.txt", ex.TargetSite + "\n");
                MessageBox.Show(ex.Message);
            }
        }
        private static void IconReset()
        {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            process.StartInfo.FileName = "ie4uinit.exe";
            process.StartInfo.Arguments = " -show";
            process.Start();
            process.WaitForExit();
        }
    }
}
