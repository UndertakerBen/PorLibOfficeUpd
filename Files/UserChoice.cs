using System;
using System.IO;
using System.Globalization;
using System.Windows.Forms;

namespace Portable_Libre_Office
{
    public partial class UserChoice : Form
    {
        private readonly string applicationPath = Application.StartupPath;
        public UserChoice(string filename)
        {
            InitializeComponent();
            //if (!filename.Contains("helppack"))
            //{
                if (File.Exists(applicationPath + "\\Setup.cfg"))
                {
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        foreach (string line in File.ReadLines(@"Setup.cfg"))
                        {
                            if (line.Contains("D-" + checkedListBox1.GetItemText(i) + "-" + Convert.ToString(checkedListBox1.Items[i])))
                            {
                                string[] linesplitt = line.Split(new char[] { '=' }, 2);
                                if (Convert.ToInt32(linesplitt[1]) == 1)
                                {
                                    checkedListBox1.SetItemChecked(i, true);
                                }
                            }
                        }
                    }
                    for (int i = 0; i < checkedListBox4.Items.Count; i++)
                    {
                        foreach (string line in File.ReadLines(@"Setup.cfg"))
                        {
                            if (line.Contains("UI-" + checkedListBox4.GetItemText(i) + "-" + Convert.ToString(checkedListBox4.Items[i])))
                            {
                                string[] linesplitt = line.Split(new char[] { '=' }, 2);
                                if (Convert.ToInt32(linesplitt[1]) == 1)
                                {
                                    checkedListBox4.SetItemChecked(i, true);
                                }
                            }
                        }
                    }
                    for (int i = 0; i < checkedListBox3.Items.Count; i++)
                    {
                        foreach (string line in File.ReadLines(@"Setup.cfg"))
                        {
                            if (line.Contains("E-" + checkedListBox3.GetItemText(i) + "-" + Convert.ToString(checkedListBox3.Items[i])))
                            {
                                string[] linesplitt = line.Split(new char[] { '=' }, 2);
                                if (Convert.ToInt32(linesplitt[1]) == 1)
                                {
                                    checkedListBox3.SetItemChecked(i, true);
                                }
                            }
                        }
                    }
                    for (int i = 0; i < checkedListBox2.Items.Count; i++)
                    {
                        foreach (string line in File.ReadLines(@"Setup.cfg"))
                        {
                            if (line.Contains("Ext-" + checkedListBox2.GetItemText(i) + "-" + Convert.ToString(checkedListBox2.Items[i])))
                            {
                                string[] linesplitt = line.Split(new char[] { '=' }, 2);
                                if (Convert.ToInt32(linesplitt[1]) == 1)
                                {
                                    checkedListBox2.SetItemChecked(i, true);
                                }
                            }
                        }
                    }
                    foreach (string line in File.ReadLines(@"Setup.cfg"))
                    {
                        if (line.Contains("HolgiMode"))
                        {
                            string[] linesplitt = line.Split(new char[] { '=' }, 2);
                            if (Convert.ToInt32(linesplitt[1]) == 1)
                            {
                                checkBox1.Checked = true;
                            }
                        }
                    }
                }
                else if (!File.Exists(applicationPath + "\\Setup.cfg"))
                {
                    CultureInfo culture1 = CultureInfo.CurrentUICulture;
                    switch (culture1.TwoLetterISOLanguageName)
                    {
                        case "de":
                            checkedListBox1.SetItemChecked(19, true);
                            checkedListBox4.SetItemChecked(37, true);
                            break;
                        default:
                            checkedListBox1.SetItemChecked(15, true);
                            checkedListBox4.SetItemChecked(28, true);
                            break;
                    }
                }
                if (filename == null)
                {
                    File.Delete(applicationPath + "\\Setup.cfg");
                }

                button1.Click += new System.EventHandler(Button1_Click);
                void Button1_Click(object sender, EventArgs e)
                {
                    File.AppendAllText(@"Setup.cfg", "[Dictionaries]" + Environment.NewLine);
                    for (int i = 0; i < checkedListBox1.Items.Count; i++)
                    {
                        File.AppendAllText(@"Setup.cfg", "D-" + checkedListBox1.GetItemText(i) + "-" + Convert.ToString(checkedListBox1.Items[i]) + "=" + Convert.ToInt32(checkedListBox1.GetItemCheckState(i)) + Environment.NewLine);
                    }
                    File.AppendAllText(@"Setup.cfg", Environment.NewLine);
                    File.AppendAllText(@"Setup.cfg", "[Extras]" + Environment.NewLine);
                    for (int i = 0; i < checkedListBox3.Items.Count; i++)
                    {
                        File.AppendAllText(@"Setup.cfg", "E-" + checkedListBox3.GetItemText(i) + "-" + Convert.ToString(checkedListBox3.Items[i]) + "=" + Convert.ToInt32(checkedListBox3.GetItemCheckState(i)) + Environment.NewLine);
                    }
                    File.AppendAllText(@"Setup.cfg", Environment.NewLine);
                    File.AppendAllText(@"Setup.cfg", "[Extensions]" + Environment.NewLine);
                    for (int i = 0; i < checkedListBox2.Items.Count; i++)
                    {
                        File.AppendAllText(@"Setup.cfg", "Ext-" + checkedListBox2.GetItemText(i) + "-" + Convert.ToString(checkedListBox2.Items[i]) + "=" + Convert.ToInt32(checkedListBox2.GetItemCheckState(i)) + Environment.NewLine);
                    }
                    File.AppendAllText(@"Setup.cfg", Environment.NewLine);
                    File.AppendAllText(@"Setup.cfg", "[User interface languages]" + Environment.NewLine);
                    for (int i = 0; i < checkedListBox4.Items.Count; i++)
                    {
                        File.AppendAllText(@"Setup.cfg", "UI-" + checkedListBox4.GetItemText(i) + "-" + Convert.ToString(checkedListBox4.Items[i]) + "=" + Convert.ToInt32(checkedListBox4.GetItemCheckState(i)) + Environment.NewLine);
                    }
                    File.AppendAllText(@"Setup.cfg", Environment.NewLine);
                    File.AppendAllText(@"Setup.cfg", "[HolgisInsaneMode]" + Environment.NewLine);
                    File.AppendAllText(@"Setup.cfg", "HolgiMode" + "=" + Convert.ToInt32(checkBox1.Checked) + Environment.NewLine);
                    if (filename == null)
                    {
                        Close();
                    }
                    else
                    {
                        Clean(filename);
                        Close();
                    }
                }
            //}
            
        }
        public static void Clean(string filename)
        {
            string[,] dictLangs = new string[52, 2] { { "Afrikaans", "af" }, { "Albanian", "sq" }, { "Arabic", "ar" }, { "Aragonese", "an" }, { "Belarusian", "be" }, { "Bengali", "bn" }, { "Bosnian", "bs" }, { "Breton", "br" }, { "Bulgarian", "bg" }, { "Catalan", "ca" }, { "Classical Tibetan", "bo" }, { "Croatian", "hr" }, { "Czech", "cs" }, { "Danish", "da" }, { "Dutch", "nl" }, { "English", "en" }, { "Estonian", "et" }, { "French", "fr" }, { "Galician", "gl" }, { "German", "de" }, { "Greek", "el" }, { "Gujarati", "gu" }, { "Hebrew", "he" }, { "Hindi", "hi" }, { "Hungarian", "hu" }, { "Icelandic", "is" }, { "Indonesian", "id" }, { "Italian", "it" }, { "Lao", "lo" }, { "Latvian", "lv" }, { "Lithuanian", "lt" }, { "Nepali", "ne" }, { "Norwegian", "no" }, { "Occitan", "oc" }, { "Polish", "pl" }, { "Portuguese", "pt-PT" }, { "Portuguese (Brazil)", "pt-BR" }, { "Romanian", "ro" }, { "Russian", "ru" }, { "Scottish Gaelic", "gd" }, { "Serbian", "sr" }, { "Sinhala", "si" }, { "Slovak", "sk" }, { "Slovenian", "sl" }, { "Spanish", "es" }, { "Swedish", "sv" }, { "Telugu", "te" }, { "Thai", "th" }, { "Turkish", "tr" }, { "Ukrainian", "uk" }, { "Vietnamese", "vi" }, { "Zulu", "zu" } };
            string[,] uiLangs = new string[119, 2] { { "Afrikaans", "af" }, { "Albanian", "sq" }, { "Amharic", "am" }, { "Arabic", "ar" }, { "Assamese", "as" }, { "Asturian", "ast" }, { "Basque", "eu" }, { "Belarusian", "be" }, { "Bengali (Bangladesh)", "bn" }, { "Bengali (India)", "bn_IN" }, { "Bodo", "brx" }, { "Bosnian", "bs" }, { "Breton", "br" }, { "Bulgarian", "bg" }, { "Burmese", "my" }, { "Catalan", "ca" }, { "Catalan (Valencian)", "ca@valencia" }, { "Central Kurdish", "ckb" }, { "Chinese (simplified)", "zh_CN" }, { "Chinese (traditional)", "zh_TW" }, { "Croatian", "hr" }, { "Czech", "cs" }, { "Danish", "da" }, { "Dogri", "dgo" }, { "Dutch", "nl" }, { "Dzongkha", "dz" }, { "English (South Africa)", "en_ZA" }, { "English (United Kingdom)", "en_GB" }, { "English (United States)", "en_US" }, { "Esperanto", "eo" }, { "Estonian", "et" }, { "Finnish", "fi" }, { "French", "fr" }, { "Frisian", "fy" }, { "Friulian", "fur" }, { "Galician", "gl" }, { "Georgian", "ka" }, { "German", "de" }, { "Greek", "el" }, { "Guarani", "gug" }, { "Gujarati", "gu" }, { "Hebrew", "he" }, { "Hindi", "hi" }, { "Hungarian", "hu" }, { "Icelandic", "is" }, { "Indonesian", "id" }, { "Irish", "ga" }, { "Italian", "it" }, { "Japanese", "ja" }, { "Kabyle", "kab" }, { "Kannada", "kn" }, { "Kashmiri", "ks" }, { "Kazakh", "kk" }, { "Khmer", "km" }, { "Kinyarwanda", "rw" }, { "Konkani", "kok" }, { "Korean", "ko" }, { "Kurdish", "kmr@latin" }, { "Lao", "lo" }, { "Latvian", "lv" }, { "Lithuanian", "lt" }, { "Lower Sorbian", "dsb" }, { "Luxembourgish", "lb" }, { "Macedonian", "mk" }, { "Maithili", "mai" }, { "Malayalam", "ml" }, { "Manipuri", "mni" }, { "Marathi", "mr" }, { "Mongolian", "mn" }, { "Ndebele South", "nr" }, { "Nepali", "ne" }, { "Northern Sotho", "nso" }, { "Norwegian (Bokmål)", "nb" }, { "Norwegian (Nynorsk)", "nn" }, { "Occitan", "oc" }, { "Odia", "or" }, { "Oromo", "om" }, { "Persian", "fa" }, { "Polish", "pl" }, { "Portuguese", "pt" }, { "Portuguese (Brazil)", "pt_BR" }, { "Punjabi", "pa_IN" }, { "Romanian", "ro" }, { "Russian", "ru" }, { "Sanskrit (India)", "sa_IN" }, { "Santili", "sat" }, { "Scottish Gaelic", "gd" }, { "Serbian (Cyrillic)", "sr" }, { "Serbian (Latin)", "sr@latin" }, { "Sidama", "sid" }, { "Silesian", "szl" }, { "Sindhi", "sd" }, { "Sinhala", "si" }, { "Slovak", "sk" }, { "Slovenian", "sl" }, { "Southern Sotho (Sutu)", "st" }, { "Spanish", "es" }, { "Swahili", "sw_TZ" }, { "Swazi", "ss" }, { "Swedish", "sv" }, { "Tajik", "tg" }, { "Tamil", "ta" }, { "Tatar", "tt" }, { "Telugu", "te" }, { "Thai", "th" }, { "Tibetan", "bo" }, { "Tsonga", "ts" }, { "Tswana", "tn" }, { "Turkish", "tr" }, { "Ukrainian", "uk" }, { "Upper Sorbian", "hsb" }, { "Uyghur", "ug" }, { "Uzbek", "uz" }, { "Venda", "ve" }, { "Venetian", "vec" }, { "Vietnamese", "vi" }, { "Welsh", "cy" }, { "Xhosa", "xh" }, { "Zulu", "zu" } };
            string applicationPath = Application.StartupPath;
            //if (!filename.Contains("helppack"))
            //{
            for (int i = 0; i < uiLangs.GetLength(0); i++)
            {
                foreach (string line in File.ReadLines(@"Setup.cfg"))
                {
                    if (line.Contains("UI-" + i + "-" + uiLangs[i, 0]))
                    {
                        string[] linesplitt = line.Split(new char[] { '=' }, 2);
                        if (Convert.ToInt32(linesplitt[1]) == 0)
                        {
                            if (Directory.Exists(applicationPath + "\\Libre Office\\program\\resource\\" + uiLangs[i, 1]))
                            {
                                Directory.Delete(applicationPath + "\\Libre Office\\program\\resource\\" + uiLangs[i, 1], true);
                            }
                            if (Directory.Exists(applicationPath + "\\Libre Office\\share\\autotext\\" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn")))
                            {
                                Directory.Delete(applicationPath + "\\Libre Office\\share\\autotext\\" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn"), true);
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\readmes\\readme_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".txt"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\readmes\\readme_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".txt");
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\Langpack-" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\registry\\Langpack-" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd");
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\ctl_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\registry\\ctl_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd");
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\ctlseqcheck_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\registry\\ctlseqcheck_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd");
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\cjk_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\registry\\cjk_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd");
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\res\\fcfg_langpack_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\registry\\res\\fcfg_langpack_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd");
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\res\\registry_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\registry\\res\\registry_" + uiLangs[i, 1].Replace("_", "-").Replace("@", "-").Replace("latin", "latn") + ".xcd");
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\Scripts\\python\\LibreLogo\\LibreLogo_" + uiLangs[i, 1].Replace("@", "_").Replace("latin", "Latn") + ".properties"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\Scripts\\python\\LibreLogo\\LibreLogo_" + uiLangs[i, 1].Replace("@", "_").Replace("latin", "Latn") + ".properties");
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\wizards\\resources_" + uiLangs[i, 1].Replace("@", "_").Replace("latin", "Latn") + ".properties"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\wizards\\resources_" + uiLangs[i, 1].Replace("@", "_").Replace("latin", "Latn") + ".properties");
                            }
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_" + uiLangs[i, 1].Replace("@", "-").Replace("_", "-").Replace("latin", "Latn") + ".dat"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_" + uiLangs[i, 1].Replace("@", "-").Replace("_", "-").Replace("latin", "Latn") + ".dat");
                            }
                            if (uiLangs[i, 1] == "af")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_af-ZA.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_af-ZA.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "bg")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_bg-BG.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_bg-BG.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "ca")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ca-ES.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ca-ES.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "cs")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_cs-CZ.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_cs-CZ.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "da")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_da-DK.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_da-DK.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "el")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_el-GR.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_el-GR.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "en_GB")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_en-GB.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_en-GB.dat");
                                }
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_en-AU.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_en-AU.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "en_US")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_en-US.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_en-US.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "en_ZA")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_en-ZA.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_en-ZA.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "fa")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_fa-IR.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_fa-IR.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "fi")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_fi-FI.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_fi-FI.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "ga")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ga-IE.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ga-IE.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "hr")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_hr-HR.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_hr-HR.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "hu")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_hu-HU.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_hu-HU.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "is")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_is-IS.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_is-IS.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "ja")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ja-JP.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ja-JP.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "ko")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ko-KR.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ko-KR.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "lb")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_lb-LU.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_lb-LU.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "lt")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_lt-LT.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_lt-LT.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "mn")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_mn-MN.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_mn-MN.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "nl")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_nl-BE.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_nl-BE.dat");
                                }
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_nl-NL.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_nl-NL.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "pl")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_pl-PL.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_pl-PL.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "pt")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_pt-PT.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_pt-PT.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "pt_BR")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_pt-BR.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_pt-BR.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "ro")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ro-RO.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ro-RO.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "ru")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ru-RU.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_ru-RU.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "sk")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sk-SK.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sk-SK.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "sl")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sl-SI.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sl-SI.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "sr")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-CS.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-CS.dat");
                                }
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-ME.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-ME.dat");
                                }
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-RS.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-RS.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "sr@latin")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-Latn-CS.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-Latn-CS.dat");
                                }
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-Latn-ME.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-Latn-ME.dat");
                                }
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-Latn-RS.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sr-Latn-RS.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "sv")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sv-SE.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_sv-SE.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "tr")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_tr-TR.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_tr-TR.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "vi")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_vi-VN.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_vi-VN.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "zh_CN")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_zh-CN.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_zh-CN.dat");
                                }
                            }
                            if (uiLangs[i, 1] == "zh_TW")
                            {
                                if (File.Exists(applicationPath + "\\Libre Office\\share\\autocorr\\acor_zh-TW.dat"))
                                {
                                    File.Delete(applicationPath + "\\Libre Office\\share\\autocorr\\acor_zh-TW.dat");
                                }
                            }
                        }
                    }
                }
            }
            for (int i = 0; i < dictLangs.GetLength(0); i++)
            {
                foreach (string line in File.ReadLines(@"Setup.cfg"))
                {
                    if (line.Contains("D-" + i + "-" + dictLangs[i, 0]))
                    {
                        string[] linesplitt = line.Split(new char[] { '=' }, 2);
                        if (Convert.ToInt32(linesplitt[1]) == 0)
                        {
                            if (Directory.Exists(applicationPath + "\\Libre Office\\share\\extensions\\dict-" + dictLangs[i, 1]))
                            {
                                Directory.Delete(applicationPath + "\\Libre Office\\share\\extensions\\dict-" + dictLangs[i, 1], true);
                            }
                        }
                    }
                }
            }
            foreach (string line in File.ReadLines(@"Setup.cfg"))
            {
                if (line.Contains("Image Filters"))
                {
                    string[] linesplitt = line.Split(new char[] { '=' }, 2);
                    if (Convert.ToInt32(linesplitt[1]) == 0)
                    {
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\graphicfilterlo.dll"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\graphicfilterlo.dll");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\svgfilterlo.dll"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\svgfilterlo.dll");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\wpftdrawlo.dll"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\wpftdrawlo.dll");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\graphicfilter.xcd"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\share\\registry\\graphicfilter.xcd");
                        }
                    }
                }
                else if (line.Contains("XSLT Sample Filters"))
                {
                    string[] linesplitt = line.Split(new char[] { '=' }, 2);
                    if (Convert.ToInt32(linesplitt[1]) == 0)
                    {
                        if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\xsltfilter.xcd"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\share\\registry\\xsltfilter.xcd");
                        }
                        if (Directory.Exists(applicationPath + "\\Libre Office\\share\\xslt"))
                        {
                            Directory.Delete(applicationPath + "\\Libre Office\\share\\xslt", true);
                        }
                    }
                }
                else if (line.Contains("LibreLogo"))
                {
                    string[] linesplitt = line.Split(new char[] { '=' }, 2);
                    if (Convert.ToInt32(linesplitt[1]) == 0)
                    {
                        if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\librelogo.xcd"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\share\\registry\\librelogo.xcd");
                        }
                        if (Directory.Exists(applicationPath + "\\Libre Office\\share\\Scripts\\python\\LibreLogo"))
                        {
                            Directory.Delete(applicationPath + "\\Libre Office\\share\\Scripts\\python\\LibreLogo", true);
                        }
                    }
                }
                else if (line.Contains("BeanShell"))
                {
                    string[] linesplitt = line.Split(new char[] { '=' }, 2);
                    if (Convert.ToInt32(linesplitt[1]) == 0)
                    {
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\bsh.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\bsh.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\ScriptProviderForBeanShell.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\ScriptProviderForBeanShell.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\services\\scriptproviderforbeanshell.rdb"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\services\\scriptproviderforbeanshell.rdb");
                        }
                    }
                }
                else if (line.Contains("JavaScript"))
                {
                    string[] linesplitt = line.Split(new char[] { '=' }, 2);
                    if (Convert.ToInt32(linesplitt[1]) == 0)
                    {
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\js.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\js.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\ScriptProviderForJavaScript.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\ScriptProviderForJavaScript.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\services\\scriptproviderforjavascript.rdb"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\services\\scriptproviderforjavascript.rdb");
                        }
                    }
                }
                else if (line.Contains("MediaWiki Publisher"))
                {
                    string[] linesplitt = line.Split(new char[] { '=' }, 2);
                    if (Convert.ToInt32(linesplitt[1]) == 0)
                    {
                        if (Directory.Exists(applicationPath + "\\Libre Office\\share\\extensions\\wiki-publisher"))
                        {
                            Directory.Delete(applicationPath + "\\Libre Office\\share\\extensions\\wiki-publisher", true);
                        }
                    }
                }
                else if (line.Contains("Programming"))
                {
                    string[] linesplitt = line.Split(new char[] { '=' }, 2);
                    if (Convert.ToInt32(linesplitt[1]) == 0)
                    {
                        if (Directory.Exists(applicationPath + "\\Libre Office\\share\\extensions\\nlpsolver"))
                        {
                            Directory.Delete(applicationPath + "\\Libre Office\\share\\extensions\\nlpsolver", true);
                        }
                    }
                }
                else if (line.Contains("Report Builder"))
                {
                    string[] linesplitt = line.Split(new char[] { '=' }, 2);
                    if (Convert.ToInt32(linesplitt[1]) == 0)
                    {
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\commons-logging-1.2.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\commons-logging-1.2.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\flow-engine.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\flow-engine.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\flute-1.1.6.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\flute-1.1.6.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\libbase-1.1.6.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\libbase-1.1.6.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\libfonts-1.1.6.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\libfonts-1.1.6.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\libformula-1.1.7.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\libformula-1.1.7.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\liblayout.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\liblayout.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\libloader-1.1.6.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\libloader-1.1.6.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\librepository-1.1.6.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\librepository-1.1.6.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\libserializer-1.1.6.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\libserializer-1.1.6.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\libxml-1.1.7.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\libxml-1.1.7.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\reportbuilder.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\reportbuilder.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\reportbuilderwizard.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\reportbuilderwizard.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\sac.jar"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\classes\\sac.jar");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\rptlo.dll"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\rptlo.dll");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\rptuilo.dll"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\rptuilo.dll");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\program\\rptxmllo.dll"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\program\\rptxmllo.dll");
                        }
                        if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\reportbuilder.xcd"))
                        {
                            File.Delete(applicationPath + "\\Libre Office\\share\\registry\\reportbuilder.xcd");
                        }
                    }
                }
            }
            if (File.Exists(applicationPath + "\\Libre Office\\program\\so_activex.dll"))
            {
                File.Delete(applicationPath + "\\Libre Office\\program\\so_activex.dll");
            }
            if (File.Exists(applicationPath + "\\Libre Office\\share\\registry\\onlineupdate.xcd"))
            {
                File.Delete(applicationPath + "\\Libre Office\\share\\registry\\onlineupdate.xcd");
            }
            if (File.Exists(applicationPath + "\\Libre Office\\program\\quickstart.exe"))
            {
                File.Delete(applicationPath + "\\Libre Office\\program\\quickstart.exe");
            }
            if (File.Exists(applicationPath + "\\Libre Office\\share\\template\\common\\wizard\\report\\default.otr"))
            {
                File.Delete(applicationPath + "\\Libre Office\\share\\template\\common\\wizard\\report\\default.otr");
            }
            if (Directory.Exists(applicationPath + "\\Libre Office\\program\\shlxthdl"))
            {
                Directory.Delete(applicationPath + "\\Libre Office\\program\\shlxthdl", true);
            }
            if (Directory.Exists(applicationPath + "\\Libre Office\\Fonts"))
            {
                Directory.Delete(applicationPath + "\\Libre Office\\Fonts", true);
            }
            if (File.Exists(applicationPath + "\\Libre Office\\" + filename))
            {
                File.Delete(applicationPath + "\\Libre Office\\" + filename);
            }
            if (!Directory.Exists(applicationPath + "\\Libre Office\\System64"))
            {
                if (File.Exists(applicationPath + "\\Libre Office\\System\\msvcp140.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System\\msvcp140.dll", applicationPath + "\\Libre Office\\program\\msvcp140.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System\\vccorlib140.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System\\vccorlib140.dll", applicationPath + "\\Libre Office\\program\\vccorlib140.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System\\vcruntime140.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System\\vcruntime140.dll", applicationPath + "\\Libre Office\\program\\vcruntime140.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System\\msvcp140_1.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System\\msvcp140_1.dll", applicationPath + "\\Libre Office\\program\\msvcp140_1.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System\\msvcp140_2.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System\\msvcp140_2.dll", applicationPath + "\\Libre Office\\program\\msvcp140_2.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System\\msvcp140_codecvt_ids.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System\\msvcp140_codecvt_ids.dll", applicationPath + "\\Libre Office\\program\\msvcp140_codecvt_ids.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System\\ucrtbase.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System\\ucrtbase.dll", applicationPath + "\\Libre Office\\program\\ucrtbase.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System\\concrt140.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System\\concrt140.dll", applicationPath + "\\Libre Office\\program\\concrt140.dll", true);
                }
            }
            else if (Directory.Exists(applicationPath + "\\Libre Office\\System64"))
            {
                if (File.Exists(applicationPath + "\\Libre Office\\System64\\msvcp140.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System64\\msvcp140.dll", applicationPath + "\\Libre Office\\program\\msvcp140.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System64\\vccorlib140.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System64\\vccorlib140.dll", applicationPath + "\\Libre Office\\program\\vccorlib140.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System64\\vcruntime140.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System64\\vcruntime140.dll", applicationPath + "\\Libre Office\\program\\vcruntime140.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System64\\msvcp140_1.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System64\\msvcp140_1.dll", applicationPath + "\\Libre Office\\program\\msvcp140_1.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System64\\msvcp140_2.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System64\\msvcp140_2.dll", applicationPath + "\\Libre Office\\program\\msvcp140_2.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System64\\msvcp140_codecvt_ids.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System64\\msvcp140_codecvt_ids.dll", applicationPath + "\\Libre Office\\program\\msvcp140_codecvt_ids.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System64\\ucrtbase.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System64\\ucrtbase.dll", applicationPath + "\\Libre Office\\program\\ucrtbase.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System64\\concrt140.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System64\\concrt140.dll", applicationPath + "\\Libre Office\\program\\concrt140.dll", true);
                }
                if (File.Exists(applicationPath + "\\Libre Office\\System64\\vcruntime140_1.dll"))
                {
                    File.Copy(applicationPath + "\\Libre Office\\System64\\vcruntime140_1.dll", applicationPath + "\\Libre Office\\program\\vcruntime140_1.dll", true);
                }
            }
            foreach (string line in File.ReadLines(@"Setup.cfg"))
            {
                if (line.Contains("HolgiMode"))
                {
                    string[] linesplitt = line.Split(new char[] { '=' }, 2);
                    if (Convert.ToInt32(linesplitt[1]) == 1)
                    {
                        DirectoryInfo diTop = new DirectoryInfo(applicationPath + "\\Libre Office\\share\\config");
                        foreach (var files in diTop.GetFiles("*.zip"))
                        {
                            if (!files.ToString().Contains("colibre.zip"))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\config\\" + files.ToString());
                            }
                        }
                        string[] delJar = new string[20] { "commonwizards.jar", "form.jar", "hsqldb.jar", "java_uno.jar", "juh.jar", "jurt.jar", "libreoffice.jar", "officebean.jar", "query.jar", "report.jar", "ridl.jar", "ScriptFramework.jar", "ScriptProviderForJava.jar", "sdbc_hsqldb.jar", "smoketest.jar", "table.jar", "unoil.jar", "unoloader.jar", "xmerge.jar", "XMergeBridge.jar" };
                        for (int i = 0; i < delJar.Length; i++)
                        {
                            if (File.Exists(applicationPath + "\\Libre Office\\program\\classes\\" + delJar[i]))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\program\\classes\\" + delJar[i]);
                            }
                        }
                        string[] dirDel = new string[5] { "help", "presets", "readmes", "System", "System64" };
                        for (int i = 0; i < dirDel.Length; i++)
                        {
                            if (Directory.Exists(applicationPath + "\\Libre Office\\" + dirDel[i]))
                            {
                                Directory.Delete(applicationPath + "\\Libre Office\\" + dirDel[i], true);
                            }
                        }
                        string[] dirDel2 = new string[5] { "opencl", "opengl", "python-core-3.7.7", "wizards", "resource\\common" };
                        for (int i = 0; i < dirDel2.Length; i++)
                        {
                            if (Directory.Exists(applicationPath + "\\Libre Office\\program\\" + dirDel2[i]))
                            {
                                Directory.Delete(applicationPath + "\\Libre Office\\program\\" + dirDel2[i], true);
                            }
                        }
                        string[] dirDel3 = new string[22] { "autocorr", "basic", "calc", "classification", "dtd", "emojiconfig", "extensions", "fingerprint", "firebird", "gallery", "glade", "labels", "liblangtag", "numbertext", "opengl", "palette", "Scripts", "theme_definitions", "tipoftheday", "wizards", "wordbook", "xpdfimport" };
                        for (int i = 0; i < dirDel3.Length; i++)
                        {
                            if (Directory.Exists(applicationPath + "\\Libre Office\\share\\" + dirDel3[i]))
                            {
                                Directory.Delete(applicationPath + "\\Libre Office\\share\\" + dirDel3[i], true);
                            }
                        }
                        string[] dirDel4 = new string[10] { "webcast", "wizard", "soffice.cfg\\dbaccess", "soffice.cfg\\modules\\dbapp", "soffice.cfg\\modules\\dbbrowser", "soffice.cfg\\modules\\dbquery", "soffice.cfg\\modules\\dbrelation", "soffice.cfg\\modules\\dbreport", "soffice.cfg\\modules\\dbtable", "soffice.cfg\\modules\\dbtdata" };
                        for (int i = 0; i < dirDel4.Length; i++)
                        {
                            if (Directory.Exists(applicationPath + "\\Libre Office\\share\\config\\" + dirDel4[i]))
                            {
                                Directory.Delete(applicationPath + "\\Libre Office\\share\\config\\" + dirDel4[i], true);
                            }
                        }
                        string[] dirDel5 = new string[2] { "wizard", "common\\wizard" };
                        for (int i = 0; i < dirDel5.Length; i++)
                        {
                            if (Directory.Exists(applicationPath + "\\Libre Office\\share\\template\\" + dirDel5[i]))
                            {
                                Directory.Delete(applicationPath + "\\Libre Office\\share\\template\\" + dirDel5[i], true);
                            }
                        }
                        string[] fileDel = new string[82] { "access2base.py", "cli_basetypes.config", "cli_basetypes.dll", "cli_cppuhelper.config", "cli_cppuhelper.dll", "cli_oootypes.config", "cli_oootypes.dll", "cli_ure.config", "cli_ure.dll", "cli_uretypes.config", "cli_uretypes.dll", "CoinMP.dll", "dbahsqllo.dll", "dbalo.dll", "dbaselo.dll", "dbaxmllo.dll", "dbplo.dll", "dbpool2.dll", "dbulo.dll", "desktophelper.txt", "dict_ja.dll", "dict_zh.dll", "Engine12.dll", "gengal.exe", "gpgme-w32spawn.exe", "intro-highres.png", "java_uno.dll", "javaloaderlo.dll", "javavendors.xml", "javavmlo.dll", "JREProperties.class", "jvmfwk3.ini", "localedata_es.dll", "localedata_others.dll", "lpsolve55.dll", "mailmerge.py", "minidump_upload.exe", "msgbox.py", "mwaw.dll", "odbcconfig.exe", "officehelper.py", "opencltest.exe", "policy.1.0.cli_basetypes.dll", "policy.1.0.cli_cppuhelper.dll", "policy.1.0.cli_oootypes.dll", "policy.1.0.cli_ure.dll", "policy.1.0.cli_uretypes.dll", "postgresql-sdbc.ini", "postgresql-sdbc-impllo.dll", "postgresql-sdbclo.dll", "python.exe", "python3.dll", "python37.dll", "pythonloader.py", "pythonloader.uno.ini", "pythonloaderlo.dll", "pythonscript.py", "pyuno.pyd", "regmerge.exe", "regview.exe", "sbase.exe", "senddoc.exe", "setup.ini", "smath.exe", "soffice.com", "soffice_safe.exe", "spsupp_helper.exe", "twain32shim.exe", "ui-previewer.exe", "uno.exe", "uno.py", "unohelper.py", "unoinfo.exe", "unopkg.com", "unopkg.exe", "vccorlib140.dll", "xpdfimport.exe", "services\\postgresql-sdbc.rdb", "services\\pyuno.rdb", "services\\scriptproviderforpython.rdb", "shell\\logo.svg", "shell\\logo_inverted.svg" };
                        for (int i = 0; i < fileDel.Length; i++)
                        {
                            if (File.Exists(applicationPath + "\\Libre Office\\program\\" + fileDel[i]))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\program\\" + fileDel[i]);
                            }
                        }
                        string[] fileDel1 = new string[5] { "dbregisterpage.ui", "hangulhanjaadddialog.ui", "hangulhanjaconversiondialog.ui", "hangulhanjaeditdictdialog.ui", "hangulhanjaoptdialog.ui" };
                        for (int i = 0; i < fileDel1.Length; i++)
                        {
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\config\\soffice.cfg\\cui\\ui\\" + fileDel1[i]))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\config\\soffice.cfg\\cui\\ui\\" + fileDel1[i]);
                            }
                        }
                        string[] fileDel2 = new string[30] { "scalc\\popupmenu\\notebookbar.xml", "scalc\\toolbar\\notebookbarshortcuts.xml", "scalc\\ui\\notebookbar.ui", "scalc\\ui\\notebookbar_compact.ui", "scalc\\ui\\notebookbar_groupedbar_compact.ui", "scalc\\ui\\notebookbar_groupedbar_full.ui", "scalc\\ui\\notebookbar_groups.ui", "sdraw\\popupmenu\\notebookbar.xml", "sdraw\\toolbar\\notebookbarshortcuts.xml", "sdraw\\ui\\notebookbar.ui", "sdraw\\ui\\notebookbar_compact.ui", "sdraw\\ui\\notebookbar_groupedbar_compact.ui", "sdraw\\ui\\notebookbar_groupedbar_full.ui", "sdraw\\ui\\notebookbar_single.ui", "simpress\\popupmenu\\notebookbar.xml", "simpress\\toolbar\\notebookbarshortcuts.xml", "simpress\\ui\\notebookbar.ui", "simpress\\ui\\notebookbar_compact.ui", "simpress\\ui\\notebookbar_groupedbar_compact.ui", "simpress\\ui\\notebookbar_groupedbar_full.ui", "simpress\\ui\\notebookbar_groups.ui", "simpress\\ui\\notebookbar_single.ui", "swriter\\popupmenu\\notebookbar.xml", "swriter\\toolbar\\notebookbarshortcuts.xml", "swriter\\ui\\notebookbar.ui", "swriter\\ui\\notebookbar_compact.ui", "swriter\\ui\\notebookbar_groupedbar_compact.ui", "swriter\\ui\\notebookbar_groupedbar_full.ui", "swriter\\ui\\notebookbar_groups.ui", "swriter\\ui\\notebookbar_single.ui" };
                        for (int i = 0; i < fileDel2.Length; i++)
                        {
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\config\\soffice.cfg\\modules\\" + fileDel2[i]))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\config\\soffice.cfg\\modules\\" + fileDel2[i]);
                            }
                        }
                        string[] fileDel3 = new string[2] { "sfx\\ui\\notebookbar.ui", "sfx\\ui\\notebookbarpopup.ui" };
                        for (int i = 0; i < fileDel3.Length; i++)
                        {
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\config\\soffice.cfg\\" + fileDel3[i]))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\config\\soffice.cfg\\" + fileDel3[i]);
                            }
                        }
                        string[] fileDel4 = new string[4] { "filter\\oox-drawingml-cs-presets", "registry\\base.xcd", "registry\\math.xcd", "template\\common\\presnt\\Vintage.otp" };
                        for (int i = 0; i < fileDel4.Length; i++)
                        {
                            if (File.Exists(applicationPath + "\\Libre Office\\share\\" + fileDel4[i]))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\share\\" + fileDel4[i]);
                            }
                        }
                        string[] fileDel5 = new string[4] { "CREDITS.fodt", "LICENSE.html", "License.txt", "NOTICE" };
                        for (int i = 0; i < fileDel5.Length; i++)
                        {
                            if (File.Exists(applicationPath + "\\Libre Office\\" + fileDel5[i]))
                            {
                                File.Delete(applicationPath + "\\Libre Office\\" + fileDel5[i]);
                            }
                        }
                        if (Directory.GetFiles(applicationPath + "\\Libre Office\\program\\classes").Length == 0)
                        {
                            Directory.Delete(applicationPath + "\\Libre Office\\program\\classes", true);
                        }
                    }
                }
            }
            if (File.Exists(applicationPath + "\\Libre Office\\program\\bootstrap.ini"))
            {
                string[] test = File.ReadAllLines(applicationPath + "\\Libre Office\\program\\bootstrap.ini");
                for (int i = 0; i < test.Length; i++)
                {
                    if (test[i].Contains("UserInstallation="))
                    {
                        test[i] = "UserInstallation=$ORIGIN/../../UserData";
                        File.WriteAllLines(applicationPath + "\\Libre Office\\program\\bootstrap.ini", test);
                    }
                }
            }
            if (!File.Exists(applicationPath + "\\UserData\\user\\registrymodifications.xcu"))
            {
                Files.Registrymodifications.CreateRegistrymodification_XCU(applicationPath);
            }
        }
    }
}
