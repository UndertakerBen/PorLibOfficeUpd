using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Portable_Libre_Office
{
    public partial class Form1 : Form
    {
        private readonly string dummy = null;
        private readonly string applicationPath = Application.StartupPath;
        private readonly string labeltext;
        private readonly CultureInfo culture = CultureInfo.CurrentUICulture;
        private readonly string[,] lang = new string[65, 2] { { "Albanian shqip", "sq" }, { "Amharic አማርኛ", "am" }, { "Arabic العربية", "ar" }, { "Asturian Asturianu", "ast" }, { "Basque euskera", "eu" }, { "Bengali (India) বাংলা (ভারত)", "bn-IN" }, { "Bengali বাংলা", "bn" }, { "Bosnian Bosanski", "bs" }, { "Bulgarian български", "bg" }, { "Catalan (Valencian) català (valencià)", "ca-valencia" }, { "Catalan català", "ca" }, { "Chinese (simplified) 中文 (简体)", "zh-CN" }, { "Chinese (traditional) 中文 (正體)", "zh-TW" }, { "Croatian Hrvatski", "hr" }, { "Czech čeština", "cs" }, { "Danish dansk", "da" }, { "Dutch Nederlands", "nl" }, { "Dzongkha རྫོང་ཁ", "dz" }, { "English (GB) English (GB)", "en-GB" }, { "English (US) English (US)", "en-US" }, { "English (ZA) English (ZA)", "en-ZA" }, { "Esperanto Esperanto", "eo" }, { "Estonian eesti keel", "et" }, { "Finnish suomi", "fi" }, { "French français", "fr" }, { "Galician Galego", "gl" }, { "Georgian ქართული", "ka" }, { "German Deutsch", "de" }, { "Greek ελληνικά", "el" }, { "Gujarati ગુજરાતી", "gu" }, { "Hebrew עברית", "he" }, { "Hindi हिन्दी", "hi" }, { "Hungarian magyar", "hu" }, { "Icelandic Íslenska", "is" }, { "Indonesian Bahasa Indonesia", "id" }, { "Italian italiano", "it" }, { "Japanese 日本語", "ja" }, { "Khmer ខ្មែរ", "km" }, { "Korean 한국어 [韓國語]", "ko" }, { "Lao ພາສາລາວ", "lo" }, { "Latvian latviešu", "lv" }, { "Lithuanian Lietuvių kalba", "lt" }, { "Macedonian македонски", "mk" }, { "Nepali नेपाली", "ne" }, { "Norwegian Bokmål Bokmål", "nb" }, { "Norwegian Nynorsk Nynorsk", "nn" }, { "Oromo Afaan Oromo", "om" }, { "Polish polski", "pl" }, { "Portuguese português", "pt" }, { "Portuguese (Brazil) português (Brasil)", "pt-BR" }, { "Romanian român", "ro" }, { "Russian Русский", "ru" }, { "Sidama", "sid" }, { "Sinhala සිංහල", "si" }, { "Slovak slovenčina", "sk" }, { "Slovenian slovenski", "sl" }, { "Spanish español", "es" }, { "Swedish Svenska", "sv" }, { "Tajik тоҷикӣ", "tg" }, { "Tamil தமிழ்", "ta" }, { "Tibetan བོད་ཡིག", "bo" }, { "Turkish Türkçe", "tr" }, { "Uighur ﺉۇﻲﻏۇﺭچە ", "ug" }, { "Ukrainian Українська", "uk" }, { "Vietnamese tiếng việt", "vi" } };
        private static readonly string[] progName = new string[7] { "Base", "Calc", "Draw", "Impress", "Math", "Office", "Writer" };
        private readonly string deskDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        private readonly string startMenu = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
        string branch = "";
        private readonly int a;
        private readonly int b;
        
        public Form1()
        {
            FileVersionInfo updVersion = FileVersionInfo.GetVersionInfo(applicationPath + "\\Portable Libre Office Updater.exe");
            InitializeComponent();
            Text = Text + " - v" + updVersion.FileVersion;
            checkBox1 = new CheckBox
            {
                AutoSize = true,
                Location = new Point(8, 206),
                Name = "checkBox1",
                Size = new Size(213, 17),
                TabIndex = 8,
                Text = "Delete downloaded files after updating",
                UseVisualStyleBackColor = true
            };
            checkBox1.CheckedChanged += new System.EventHandler(this.CheckBox1_CheckedChanged);
            Controls.Add(checkBox1);
            if (File.Exists(applicationPath + "\\Setup.cfg"))
            {
                foreach (string line in File.ReadLines(@"Setup.cfg"))
                {
                    if (line.Contains("DeleteFiles"))
                    {
                        string[] linesplitt = line.Split(new char[] { '=' }, 2);
                        if (Convert.ToInt32(linesplitt[1]) == 1)
                        {
                            checkBox1.Checked = true;
                        }
                    }
                    else
                    {
                        checkBox1.Checked = false;
                    }
                }
            }
            ComboBox combobox1 = new ComboBox
            {
                Size = new Size(125, 18),
                DropDownWidth = 250
            };
            for (int i = 0; i < lang.GetLength(0); i++)
            {
                combobox1.Items.AddRange(new object[] { lang[i, 0] });
            }
            switch (culture.TwoLetterISOLanguageName)
            {
                case "de":
                    combobox1.SelectedIndex = 27;
                    break;
                default:
                    combobox1.SelectedIndex = 19;
                    break;
            }
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create("https://download.documentfoundation.org/libreoffice/testing/");
                //request.UserAgent = "LibreOffice 6.2.3.2(aecc05fe267cc68dde00352a451aa867b3b546ac; Windows; X86_64; de; )";
                var response = request.GetResponse();
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    var responseFromReader = reader.ReadToEnd();
                    string[] test = responseFromReader.Replace("<tr><td valign=\"top\">&nbsp;</td><td><a href=\"", "|").Replace("/", "").Substring(responseFromReader.IndexOf("valign=", +7)).Split(new char[] { '|' });
                    for (int i = 1; i < test.GetLength(0); i++)
                    {
                        string[] test1 = test[i].Split(new char[] { '"' });
                        ComboBox comboBox = new ComboBox
                        {
                            Size = new Size(300, 20)
                        };
                        groupBox2.Size = new Size(groupBox2.Size.Width, 25 * (test.GetLength(0) - 1));
                        groupBox3.Location = new Point(groupBox2.Location.X, groupBox2.Location.Y + groupBox2.Size.Height + 5);
                        groupBox1.Size = new Size(groupBox2.Size.Width + 18, groupBox3.Location.Y + groupBox3.Size.Height);
                        HttpWebRequest request2 = (HttpWebRequest)WebRequest.Create("https://download.documentfoundation.org/libreoffice/testing/" + test1[0] + "/win/x86_64/");
                        //request2.UserAgent = "LibreOffice 6.2.3.2(aecc05fe267cc68dde00352a451aa867b3b546ac; Windows; X86_64; de; )";
                        try
                        {
                            HttpWebResponse response2 = (HttpWebResponse)request2.GetResponse();
                            if (response2.StatusCode == HttpStatusCode.OK)
                            {
                                using (StreamReader reader2 = new StreamReader(response2.GetResponseStream()))
                                {
                                    string[] responseFromReader2 = reader2.ReadToEnd().Split(new char[] { '\n' });
                                    for (int k = 0; k < responseFromReader2.GetLength(0); k++)
                                    {
                                        if (responseFromReader2[k].Contains(".msi\""))
                                        {
                                            comboBox.Items.AddRange(new object[] { responseFromReader2[k].Substring(responseFromReader2[k].IndexOf("<td><a href=\"", +13)).Split(new char[] { '"' }, 3)[1] });
                                            if (responseFromReader2[k].Contains("x64.msi"))
                                            {
                                                string[] test3 = responseFromReader2[k].Substring(responseFromReader2[k].IndexOf("<td><a href=\"", +13)).Split(new char[] { '"' }, 3)[1].Split(new char[] { '_' }, 3);
                                                labeltext = $"{test3[0]} {test3[1]}";
                                            }
                                        }
                                    }
                                    reader2.Close();
                                }
                            }
                            Label label = new Label
                            {
                                Location = new Point(10, 20 + 25 * (a++)),
                                Size = new Size(260, 23),
                                Text = labeltext,
                                TextAlign = ContentAlignment.BottomLeft,
                                ForeColor = Color.Black,
                                Font = new Font("Microsoft Sans Serif", 9.25F, FontStyle.Bold, GraphicsUnit.Point, 0)
                            };
                            groupBox2.Controls.Add(label);
                            Button button1 = new Button
                            {
                                Text = "x86",
                                Location = new Point(label.Location.X + label.Size.Width + 20, 20 + 25 * (a-1)),
                                ForeColor = Color.Black,
                                Size = new Size(50, 23)
                            };
                            groupBox2.Controls.Add(button1);
                            Button button2 = new Button
                            {
                                Text = "x64",
                                Location = new Point(button1.Location.X + button1.Width + 5, 20 + 25 * (a-1)),
                                ForeColor = Color.Black,
                                Size = new Size(50, 23)
                            };
                            Button button3 = new Button
                            {
                                Text = "x86",
                                Location = new Point(label.Location.X, label.Location.Y),
                                Size = new Size(50, 23)
                            };
                            Button button4 = new Button
                            {
                                Text = "x64",
                                Location = new Point(button3.Location.X + button3.Width + 5, label.Location.Y),
                                Size = new Size(50, 23)
                            };
                            groupBox2.Controls.Add(button2);
                            if (IntPtr.Size != 8)
                            {
                                button2.Enabled = false;
                                button4.Enabled = false;
                            }
                            groupBox2.Padding = new Padding(5);
                            button1.Click += new EventHandler(Button1_Click);
                            button2.Click += new EventHandler(Button2_Click);
                            button3.Click += new EventHandler(Button3_Click);
                            button4.Click += new EventHandler(Button4_Click);
                            groupBox5.Controls.Add(button3);
                            groupBox5.Controls.Add(button4);
                            groupBox5.Size = new Size(groupBox5.Size.Width, 25 * (test.GetLength(0) - 1));
                            groupBox6.Location = new Point(groupBox2.Location.X, groupBox2.Location.Y + groupBox2.Size.Height + 5);
                            groupBox4.Size = new Size(groupBox2.Size.Width + 18, groupBox3.Location.Y + groupBox3.Size.Height);
                            groupBox4.Location = new Point(groupBox1.Location.X + groupBox1.Size.Width + 5, groupBox1.Location.Y);
                            comboBox.SelectedIndex = 0;
                            reader.Close();
                            async void Button1_Click(object sender, EventArgs e)
                            {
                                await DownloadProgAsync("x86", comboBox.Items[0].ToString().Replace("x64", "x86"), test1[0], "testing");
                                File.WriteAllText(applicationPath + "\\UserData\\Branch.log", "Preview");
                                NewMethod();
                            }
                            async void Button2_Click(object sender, EventArgs e)
                            {
                                await DownloadProgAsync("x86_64", comboBox.Items[0].ToString(), test1[0], "testing");
                                File.WriteAllText(applicationPath + "\\UserData\\Branch.log", "Preview");
                                NewMethod();
                            }
                            async void Button3_Click(object sender, EventArgs e)
                            {
                                for (int k = 0; k < comboBox.Items.Count; k++)
                                {
                                    comboBox.SelectedIndex = k;
                                    if (comboBox.SelectedItem.ToString().Contains("_helppack_" + lang[combobox1.SelectedIndex, 1] + ".msi"))
                                    {
                                        comboBox.SelectedIndex = k;
                                        break;
                                    }
                                }
                                File.AppendAllText(@"Helppack.txt", comboBox.SelectedItem.ToString() + "\n");
                                await DownloadProgAsync("x86", comboBox.SelectedItem.ToString().Replace("x64", "x86"), test1[0], "testing");
                            }
                            async void Button4_Click(object sender, EventArgs e)
                            {
                                for (int k = 0; k < comboBox.Items.Count; k++)
                                {
                                    comboBox.SelectedIndex = k;
                                    if (comboBox.SelectedItem.ToString().Contains("_helppack_" + lang[combobox1.SelectedIndex, 1] + ".msi"))
                                    {
                                        comboBox.SelectedIndex = k;
                                        break;
                                    }
                                }
                                File.AppendAllText(@"Helppack.txt", comboBox.SelectedItem.ToString() + "\n");
                                await DownloadProgAsync("x86_64", comboBox.SelectedItem.ToString(), test1[0], "testing");
                            }

                        }
                        catch (WebException)
                        {
                            
                        }
                    }
                }
                HttpWebRequest request3 = (HttpWebRequest)WebRequest.Create("https://download.documentfoundation.org/libreoffice/stable/");
                //request.UserAgent = "LibreOffice 6.2.3.2(aecc05fe267cc68dde00352a451aa867b3b546ac; Windows; X86_64; de; )";
                var response3 = request3.GetResponse();
                using (StreamReader reader = new StreamReader(response3.GetResponseStream()))
                {
                    var responseFromReader = reader.ReadToEnd();
                    string[] test = responseFromReader.Replace("<tr><td valign=\"top\">&nbsp;</td><td><a href=\"", "|").Replace("/", "").Substring(responseFromReader.IndexOf("valign=", +7)).Split(new char[] { '|' });
                    for (int i = 1; i < test.GetLength(0); i++)
                    {
                        string[] test1 = test[i].Split(new char[] { '"' });

                        ComboBox comboBox = new ComboBox
                        {
                            Size = new Size(300, 20)
                        };
                        groupBox3.Size = new Size(groupBox3.Size.Width, 25 * (test.GetLength(0) - 1));
                        groupBox3.Location = new Point(groupBox2.Location.X, groupBox2.Location.Y + groupBox2.Size.Height + 5);
                        groupBox1.Size = new Size(groupBox2.Size.Width + 18, groupBox3.Location.Y + groupBox3.Size.Height);
                        HttpWebRequest request2 = (HttpWebRequest)WebRequest.Create("https://download.documentfoundation.org/libreoffice/stable/" + test1[0] + "/win/x86_64/");
                        //request2.UserAgent = "LibreOffice 6.2.3.2(aecc05fe267cc68dde00352a451aa867b3b546ac; Windows; X86_64; de; )";
                        HttpWebResponse response2 = (HttpWebResponse)request2.GetResponse();
                        if (response2.StatusCode == HttpStatusCode.OK)
                        {
                            using (StreamReader reader2 = new StreamReader(response2.GetResponseStream()))
                            {
                                string[] responseFromReader2 = reader2.ReadToEnd().Split(new char[] { '\n' });
                                for (int k = 0; k < responseFromReader2.GetLength(0); k++)
                                {
                                    if (responseFromReader2[k].Contains(".msi\""))
                                    {
                                        comboBox.Items.AddRange(new object[] { responseFromReader2[k].Substring(responseFromReader2[k].IndexOf("<td><a href=\"", +13)).Split(new char[] { '"' }, 3)[1] });
                                        if (responseFromReader2[k].Contains("x64.msi"))
                                        {
                                            string[] test3 = responseFromReader2[k].Substring(responseFromReader2[k].IndexOf("<td><a href=\"", +13)).Split(new char[] { '"' }, 3)[1].Split(new char[] { '_' }, 3);
                                            labeltext = test3[0] + " " + test3[1];
                                        }
                                    }
                                }
                                reader2.Close();
                            }
                            Label label = new Label
                            {
                                Location = new Point(10, 20 + 25 * (b++)),
                                Size = new Size(260, 23),
                                Text = labeltext,
                                TextAlign = ContentAlignment.BottomLeft,
                                ForeColor = Color.Black,
                                Font = new Font("Microsoft Sans Serif", 9.25F, FontStyle.Bold, GraphicsUnit.Point, 0)
                            };
                            groupBox3.Controls.Add(label);
                            Button button1 = new Button
                            {
                                Text = "x86",
                                Location = new Point(label.Location.X + label.Size.Width + 20, 20 + 25 * (b - 1)),
                                ForeColor = Color.Black,
                                Size = new Size(50, 23)
                            };
                            groupBox3.Controls.Add(button1);
                            Button button2 = new Button
                            {
                                Text = "x64",
                                Location = new Point(button1.Location.X + button1.Width + 5, 20 + 25 * (b - 1)),
                                ForeColor = Color.Black,
                                Size = new Size(50, 23)
                            };
                            Button button3 = new Button
                            {
                                Text = "x86",
                                Location = new Point(label.Location.X, label.Location.Y),
                                Size = new Size(50, 23)
                            };
                            Button button4 = new Button
                            {
                                Text = "x64",
                                Location = new Point(button3.Location.X + button3.Width + 5, label.Location.Y),
                                Size = new Size(50, 23)
                            };
                            groupBox3.Controls.Add(button2);
                            groupBox3.Padding = new Padding(5);
                            button1.Click += new EventHandler(Button1_Click);
                            button2.Click += new EventHandler(Button2_Click);
                            button3.Click += new EventHandler(Button3_Click);
                            button4.Click += new EventHandler(Button4_Click);
                            groupBox6.Controls.Add(button3);
                            groupBox6.Controls.Add(button4);
                            groupBox6.Size = new Size(groupBox6.Size.Width, 25 * (test.GetLength(0) - 1));
                            groupBox6.Location = new Point(groupBox2.Location.X, groupBox2.Location.Y + groupBox2.Size.Height + 5);
                            groupBox4.Size = new Size(groupBox6.Width + 15, groupBox1.Height);
                            groupBox4.Location = new Point(groupBox1.Location.X + groupBox1.Size.Width + 5, groupBox1.Location.Y);
                            comboBox.SelectedIndex = 0;
                            reader.Close();
                            if (IntPtr.Size != 8)
                            {
                                button2.Enabled = false;
                                button4.Enabled = false;
                            }
                            async void Button1_Click(object sender, EventArgs e)
                            {
                                await DownloadProgAsync("x86", comboBox.Items[0].ToString().Replace("x64", "x86"), test1[0], "stable");
                                File.WriteAllText(applicationPath + "\\UserData\\Branch.log", "Stable");
                                NewMethod();
                            }
                            async void Button2_Click(object sender, EventArgs e)
                            {
                                await DownloadProgAsync("x86_64", comboBox.Items[0].ToString(), test1[0], "stable");
                                File.WriteAllText(applicationPath + "\\UserData\\Branch.log", "Stable");
                                NewMethod();
                            }
                            async void Button3_Click(object sender, EventArgs e)
                            {
                                for (int k = 0; k < comboBox.Items.Count; k++)
                                {
                                    comboBox.SelectedIndex = k;
                                    if (comboBox.SelectedItem.ToString().Contains("_helppack_" + lang[combobox1.SelectedIndex, 1] + ".msi"))
                                    {
                                        comboBox.SelectedIndex = k;
                                        break;
                                    }
                                }
                                File.AppendAllText(@"Helppack.txt", comboBox.SelectedItem.ToString() + "\n");
                                await DownloadProgAsync("x86", comboBox.SelectedItem.ToString().Replace("x64", "x86"), test1[0], "stable");
                            }
                            async void Button4_Click(object sender, EventArgs e)
                            {
                                for (int k = 0; k < comboBox.Items.Count; k++)
                                {
                                    comboBox.SelectedIndex = k;
                                    if (comboBox.SelectedItem.ToString().Contains("_helppack_" + lang[combobox1.SelectedIndex, 1] + ".msi"))
                                    {
                                        comboBox.SelectedIndex = k;
                                        break;
                                    }
                                }
                                File.AppendAllText(@"Helppack.txt", comboBox.SelectedItem.ToString() + " - " + test1[0] + "\n");
                                await DownloadProgAsync("x86_64", comboBox.SelectedItem.ToString(), test1[0], "stable");
                            }
                        }
                    }
                    checkBox1.Location = new Point(groupBox1.Location.X + 5, groupBox1.Location.Y + groupBox1.Size.Height + 10);
                    button1.Location = new Point(groupBox4.Location.X + groupBox4.Size.Width - button1.Size.Width, groupBox4.Location.Y + groupBox4.Size.Height + 5);
                    label1.Location = new Point(checkBox1.Location.X, checkBox1.Location.Y + checkBox1.Size.Height + 10);
                    button3.Location = new Point(label1.Location.X + label1.Size.Width + 5, label1.Location.Y - 5);
                    button1.Location = new Point(groupBox4.Location.X + groupBox4.Size.Width - button1.Size.Width, button3.Location.Y);
                    combobox1.Location = new Point(groupBox4.Width - groupBox4.Width + 8, groupBox4.Height - groupBox4.Height + 16);
                    groupBox4.Controls.Add(combobox1);
                    combobox1.BringToFront();
                }
            }
            catch (WebException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            if (File.Exists(applicationPath + "\\Libre Office\\program\\soffice.exe"))
            {
                CheckIsFile64bit(File.OpenRead(applicationPath + "\\Libre Office\\program\\soffice.exe"));
                groupBox1.ForeColor = SystemColors.HotTrack;
                menuStrip1.Enabled = true;
            }
            if (File.Exists(applicationPath + "\\Setup.cfg"))
            {
                button3.Visible = true;
                label1.Visible = true;
            }
            CheckUpdate();
        }
        private async Task DownloadProgAsync(string arch, string filename, string version, string ring)
        {
            GroupBox progressBox = new GroupBox
            {
                Location = new Point(groupBox1.Location.X, button1.Location.Y + button1.Size.Height + 5),
                Size = new Size(groupBox1.Width + 5 + groupBox4.Width, 90),
                BackColor = Color.Lavender,
            };
            Label title = new Label
            {
                AutoSize = false,
                Location = new Point(5, 10),
                Size = new Size(progressBox.Size.Width - 10, 25),
                Text = filename,
                TextAlign = ContentAlignment.BottomCenter
            };
            title.Font = new Font(title.Font.Name, 9.25F, FontStyle.Bold);
            Label downloadLabel = new Label
            {
                AutoSize = false,
                Location = new Point(8, 35),
                Size = new Size(200, 25),
                TextAlign = ContentAlignment.BottomLeft
            };
            Label percLabel = new Label
            {
                AutoSize = false,
                Location = new Point(progressBox.Size.Width - 108, 35),
                Size = new Size(100, 25),
                TextAlign = ContentAlignment.BottomRight
            };
            ProgressBar progressBarneu = new ProgressBar
            {
                Location = new Point(8, 65),
                Size = new Size(progressBox.Size.Width - 18, 7)
            };
            progressBox.Controls.Add(title);
            progressBox.Controls.Add(downloadLabel);
            progressBox.Controls.Add(percLabel);
            progressBox.Controls.Add(progressBarneu);
            Controls.Add(progressBox);
            List<Task> list = new List<Task>();
            if (!File.Exists(applicationPath + "\\" + filename))
            {
                Uri uri = new Uri($"https://download.documentfoundation.org/libreoffice/{ring}/{version}/win/{arch}/{filename}");
                WebClient webClient;
                using (webClient = new WebClient())
                {
                    webClient.DownloadProgressChanged += (o, args) =>
                    {
                        progressBarneu.Value = args.ProgressPercentage;
                        downloadLabel.Text = $"{args.BytesReceived / 1024d / 1024d:0.00} MB's / {args.TotalBytesToReceive / 1024d / 1024d:0.00} MB's";
                        percLabel.Text = $"{args.ProgressPercentage}%";
                    };
                    webClient.DownloadFileCompleted += async (o, args) =>
                    {
                        if (args.Cancelled == true)
                        {
                            MessageBox.Show("Download has been canceled.");
                        }
                        await NewMethod1(filename);
                    };
                    try
                    {
                        var task = webClient.DownloadFileTaskAsync(uri, filename);
                        list.Add(task);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
            else if (File.Exists(applicationPath + "\\" + filename))
            {
                progressBox.Controls.Remove(progressBarneu);
                downloadLabel.AutoSize = true;
                downloadLabel.Font = new Font(downloadLabel.Font.Name, 12F, downloadLabel.Font.Style, downloadLabel.Font.Unit);
                downloadLabel.Location = new Point(8, 50);
                downloadLabel.Text = "Please Wait... it's unpacking.";
                await NewMethod1(filename);
            }
            await Task.WhenAll(list);
            await Task.Delay(2000);
            Controls.Remove(progressBox);
        }
        private async Task NewMethod1(string filename)
        {
            if (!filename.Contains("helppack"))
            {
                try
                {
                    if (Directory.Exists(applicationPath + "\\Libre Office"))
                    {
                        Directory.Delete(applicationPath + "\\Libre Office", true);
                    }
                    Thread.Sleep(1000);
                    if (Directory.Exists(applicationPath + "\\Libre Office"))
                    {
                        Directory.Delete(applicationPath + "\\Libre Office", true);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                string arguments = " /a \"" + filename + "\" /qb Targetdir=\"" + applicationPath + "\\Libre Office\"";
                Process process = new Process();
                process.StartInfo.FileName = "msiexec";
                process.StartInfo.Arguments = arguments;
                process.Start();
                process.WaitForExit();
                if (!File.Exists(applicationPath + "\\Setup.cfg"))
                {
                    UserChoice userChoice = new UserChoice(filename);
                    userChoice.ShowDialog();
                }
                else
                {
                    UserChoice.Clean(filename);
                }
                for (int i = 0; i < progName.GetLength(0); i++)
                {
                    if (File.Exists(applicationPath + "\\Libre Office\\program\\s" + progName[i].ToLower() + ".exe"))
                    {
                        if (!File.Exists(applicationPath + "\\Libre " + progName[i] + " Portable Launcher.exe"))
                        {
                            File.Copy(applicationPath + "\\Bin\\Libre " + progName[i] + " Portable Launcher.exe", applicationPath + "\\Libre " + progName[i] + " Portable Launcher.exe");
                        }
                    }
                    else if (!File.Exists(applicationPath + "\\Libre Office\\program\\s" + progName[i].ToLower() + ".exe"))
                    {
                        if (File.Exists(applicationPath + "\\Libre " + progName[i] + " Portable Launcher.exe"))
                        {
                            File.Delete(applicationPath + " \\Libre " + progName[i] + " Portable Launcher.exe");
                        }
                        if (File.Exists(deskDir + "\\Libre " + progName[i] + " Portable Launcher.lnk"))
                        {
                            File.Delete(deskDir + "\\Libre " + progName[i] + " Portable Launcher.lnk");
                        }
                        if (File.Exists(startMenu + "\\LibreOffice Portable\\Libre " + progName[i] + " Portable Launcher.lnk"))
                        {
                            File.Delete(startMenu + "\\LibreOffice Portable\\Libre " + progName[i] + " Portable Launcher.lnk");
                        }
                        if (File.Exists(deskDir + "\\Libre " + progName[i] + " Portable.lnk"))
                        {
                            File.Delete(deskDir + "\\Libre " + progName[i] + " Portable.lnk");
                        }
                        if (File.Exists(startMenu + "\\LibreOffice Portable\\Libre " + progName[i] + " Portable.lnk"))
                        {
                            File.Delete(startMenu + "\\LibreOffice Portable\\Libre " + progName[i] + " Portable.lnk");
                        }
                    }
                }
                if (checkBox1.Checked)
                {
                    if (File.Exists(applicationPath + "\\" + filename))
                    {
                        File.Delete(applicationPath + "\\" + filename);
                    }
                }
                await Task.WhenAll();
            }
            else if (filename.Contains("helppack"))
            {
                string arguments = " /a \"" + filename + "\" /qb Targetdir=\"" + applicationPath + "\\Libre Office\"";
                Process process = new Process();
                process.StartInfo.FileName = "msiexec";
                process.StartInfo.Arguments = arguments;
                process.Start();
                process.WaitForExit();
                if (File.Exists(applicationPath + "\\Libre Office\\" + filename))
                {
                    File.Delete(applicationPath + "\\Libre Office\\" + filename);
                }
                if (checkBox1.Checked)
                {
                    if (File.Exists(applicationPath + "\\" + filename))
                    {
                        File.Delete(applicationPath + "\\" + filename);
                    }
                }
            }
        }
        private void NewMethod()
        {
            if (File.Exists(applicationPath + "\\Libre Office\\program\\soffice.exe"))
            {
                CheckIsFile64bit(File.OpenRead(applicationPath + "\\Libre Office\\program\\soffice.exe"));
                groupBox1.ForeColor = SystemColors.HotTrack;
                menuStrip1.Enabled = true;
            }
            if (File.Exists(applicationPath + "\\Setup.cfg"))
            {
                button3.Visible = true;
                label1.Visible = true;
            }
        }
        private void Button3_Click(object sender, EventArgs e)
        {
            UserChoice userChoice = new UserChoice(dummy);
            userChoice.ShowDialog();
        }
        private void RegisterLibreOfficeAsStandardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Files.Regfile.RegCreate(applicationPath);
        }
        private void RemoveLibreOfficeAsStandardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Files.Regfile.RegDel(applicationPath);
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void ExtrasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Microsoft.Win32.Registry.GetValue("HKEY_Current_User\\Software\\LibreOffice.PORTABLE\\Capabilities", "ApplicationName", null) != null)
            {
                registerLibreOfficeAsStandardToolStripMenuItem.Enabled = false;
                removeLibreOfficeAsStandardToolStripMenuItem.Enabled = true;
            }
            else if (Microsoft.Win32.Registry.GetValue("HKEY_Current_User\\Software\\LibreOffice.PORTABLE\\Capabilities", "ApplicationName", null) == null)
            {
                registerLibreOfficeAsStandardToolStripMenuItem.Enabled = true;
                removeLibreOfficeAsStandardToolStripMenuItem.Enabled = false;
            }
        }
        private void OnTheDesktopToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < progName.GetLength(0); i++)
            {
                if (File.Exists(applicationPath + "\\Libre Office\\program\\s" + progName[i].ToLower() + ".exe"))
                {
                    IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                    IWshRuntimeLibrary.IWshShortcut link = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(deskDir + "\\Libre " + progName[i] + " Portable.lnk");
                    link.IconLocation = applicationPath + "\\Libre Office\\program\\s" + progName[i].ToLower() + ".exe,0";
                    link.WorkingDirectory = applicationPath;
                    link.TargetPath = applicationPath + "\\Libre " + progName[i] + " Portable Launcher.exe";
                    link.Save();
                }
            }
        }
        private void OnTheDesktopToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
            IWshRuntimeLibrary.IWshShortcut link = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(deskDir + "\\Portable Libre Office Updater.lnk");
            link.IconLocation = applicationPath + "\\Portable Libre Office Updater.exe,0";
            link.WorkingDirectory = applicationPath;
            link.TargetPath = applicationPath + "\\Portable Libre Office Updater.exe";
            link.Save();
        }

        private void InTheStartmenuToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(startMenu + "\\LibreOffice Portable"))
            {
                Directory.CreateDirectory(startMenu + "\\LibreOffice Portable");
            }
            IWshRuntimeLibrary.WshShell shell1 = new IWshRuntimeLibrary.WshShell();
            IWshRuntimeLibrary.IWshShortcut link1 = (IWshRuntimeLibrary.IWshShortcut)shell1.CreateShortcut(startMenu + "\\LibreOffice Portable\\Portable Libre Office Updater.lnk");
            link1.IconLocation = applicationPath + "\\Portable Libre Office Updater.exe,0";
            link1.WorkingDirectory = applicationPath;
            link1.TargetPath = applicationPath + "\\Portable Libre Office Updater.exe";
            link1.Save();
        }
        private void InTheStartmenuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < progName.GetLength(0); i++)
            {
                if (File.Exists(applicationPath + "\\Libre Office\\program\\s" + progName[i].ToLower() + ".exe"))
                {
                    if (!Directory.Exists(startMenu + "\\LibreOffice Portable"))
                    {
                        Directory.CreateDirectory(startMenu + "\\LibreOffice Portable");
                    }
                    IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
                    IWshRuntimeLibrary.IWshShortcut link = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(startMenu + "\\LibreOffice Portable\\Libre " + progName[i] + " Portable.lnk");
                    link.IconLocation = applicationPath + "\\Libre Office\\program\\s" + progName[i].ToLower() + ".exe,0";
                    link.WorkingDirectory = applicationPath;
                    link.TargetPath = applicationPath + "\\Libre " + progName[i] + " Portable Launcher.exe";
                    link.Save();
                }
            }
        }
        private void RemoveAllShortcutsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < progName.GetLength(0); i++)
            {
                if (File.Exists(startMenu + "\\LibreOffice Portable\\Libre " + progName[i] + " Portable.lnk"))
                {
                    File.Delete(startMenu + "\\LibreOffice Portable\\Libre " + progName[i] + " Portable.lnk");
                }
                else if (File.Exists(startMenu + "\\LibreOffice Portable\\Libre " + progName[i] + " Portable Launcher.lnk"))
                {
                    File.Delete(startMenu + "\\LibreOffice Portable\\Libre " + progName[i] + " Portable Launcher.lnk");
                }
                if (File.Exists(deskDir + "\\Libre " + progName[i] + " Portable.lnk"))
                {
                    File.Delete(deskDir + "\\Libre " + progName[i] + " Portable.lnk");
                }
                else if (File.Exists(deskDir + "\\Libre " + progName[i] + " Portable Launcher.lnk"))
                {
                    File.Delete(deskDir + "\\Libre " + progName[i] + " Portable Launcher.lnk");
                }
            }
            if (File.Exists(startMenu + "\\LibreOffice Portable\\Portable Libre Office Updater.lnk"))
            {
                File.Delete(startMenu + "\\LibreOffice Portable\\Portable Libre Office Updater.lnk");
            }
            if (File.Exists(deskDir + "\\Portable Libre Office Updater.lnk"))
            {
                File.Delete(deskDir + "\\Portable Libre Office Updater.lnk");
            }
        }
        private void CheckIsFile64bit(FileStream filestream)
        {
            //Snippet is from *Thanks*
            //"https://dotnet-snippets.de/snippet/erkennen-ob-eine-exe-oder-dll-als-64bit-kompiliert-wurde/1181"
            if (File.Exists(applicationPath + "\\UserData\\Branch.log"))
            {
                branch = File.ReadAllText(applicationPath + "\\UserData\\Branch.log");
            }
            byte[] _data = new byte[4];
            filestream.Seek(0x3c, SeekOrigin.Begin);
            filestream.Read(_data, 0, 4);
            int _offset = BitConverter.ToInt32(_data, 0);

            if (_offset > 0x3c)
            {
                _data = new byte[4];
                filestream.Seek(_offset, SeekOrigin.Begin);
                filestream.Read(_data, 0, 4);

                if ((_data[0] == 0x50)
                    && (_data[1] == 0x45)
                    && (_data[2] == 0x00)
                    && (_data[3] == 0x00))
                {
                    _data = new byte[20];
                    filestream.Read(_data, 0, 20);
                    //int _machine = BitConverter.ToInt16(_data, 0);
                    _data = new byte[2];
                    filestream.Read(_data, 0, 2);
                    int _magicNumber = BitConverter.ToInt16(_data, 0);
                    if (_magicNumber == 0x010b)
                    {
                        groupBox1.Text = "Libre Office - Installed: " + branch + " " + FileVersionInfo.GetVersionInfo(applicationPath + "\\Libre Office\\program\\soffice.exe").FileVersion + " x86";
                        filestream.Close();
                        Refresh();
                    }
                    else if (_magicNumber == 0x020b)
                    {
                        groupBox1.Text = "Libre Office - Installed: " + branch + " " + FileVersionInfo.GetVersionInfo(applicationPath + "\\Libre Office\\program\\soffice.exe").FileVersion + " x64";
                        filestream.Close();
                        Refresh();
                    }
                }
                else
                {
                    groupBox1.Text = "Libre Office - Installed: " + branch + " " + FileVersionInfo.GetVersionInfo(applicationPath + "\\Libre Office\\program\\soffice.exe").FileVersion;
                    filestream.Close();
                    Refresh();
                }
            }
            else
            {
                groupBox1.Text = "Libre Office - Installed: " + branch + " " + FileVersionInfo.GetVersionInfo(applicationPath + "\\Libre Office\\program\\soffice.exe").FileVersion;
                filestream.Close();
                Refresh();
            }
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (File.Exists(applicationPath + "\\Setup.cfg"))
            {
                try
                {
                    string[] delFiles = File.ReadAllLines(applicationPath + "\\Setup.cfg");
                    for (int i = 0; i < delFiles.Length; i++)
                    {
                        if (delFiles[i].Contains("DeleteFiles="))
                        {
                            delFiles[i] = "DeleteFiles=" + Convert.ToInt32(checkBox1.Checked);
                            File.WriteAllLines(applicationPath + "\\Setup.cfg", delFiles);
                        }
                    }
                    if (!delFiles.Contains("DeleteFiles=") & !delFiles.Contains("DeleteDownloadedFiles"))
                    {
                        File.AppendAllText(applicationPath + "\\Setup.cfg", Environment.NewLine);
                        File.AppendAllText(applicationPath + "\\Setup.cfg", "[DeleteDownloadedFiles]" + Environment.NewLine);
                        File.AppendAllText(applicationPath + "\\Setup.cfg", "DeleteFiles" + "=" + Convert.ToInt32(Form1.checkBox1.Checked) + Environment.NewLine);
                    }
                }
                catch (Exception)
                {

                }
            }
        }
        private void CheckUpdate()
        {
            GroupBox groupBoxupdate = new GroupBox
            {
                Location = new Point(groupBox1.Location.X, button1.Location.Y + button1.Size.Height + 5),
                Size = new Size(groupBox1.Width + 5 + groupBox4.Width, 90),
                BackColor = Color.Aqua
            };
            Label versionLabel = new Label
            {
                AutoSize = false,
                TextAlign = ContentAlignment.BottomCenter,
                Dock = DockStyle.None,
                Location = new Point(2, 30),
                Size = new Size(groupBoxupdate.Width - 4, 25)
            };
            versionLabel.Font = new Font(versionLabel.Font.Name, 10F, FontStyle.Bold);
            Label infoLabel = new Label
            {
                AutoSize = false,
                TextAlign = ContentAlignment.BottomCenter,
                Dock = DockStyle.None,
                Location = new Point(2, 10),
                Size = new Size(groupBoxupdate.Width - 4, 20),
                Text = "A new version is available"
            };
            infoLabel.Font = new Font(infoLabel.Font.Name, 9.25F);
            Label downLabel = new Label
            {
                TextAlign = ContentAlignment.MiddleRight,
                AutoSize = false,
                Size = new Size(100, 23),
                Text = "Update now"
            };
            Button laterButton = new Button
            {
                Size = new Size(40, 23),
                BackColor = Color.FromArgb(224, 224, 224),
                Text = "No"
            };
            Button updateButton = new Button
            {
                Location = new Point(groupBoxupdate.Width - Width - 10, 60),
                Size = new Size(40, 23),
                BackColor = Color.FromArgb(224, 224, 224),
                Text = "Yes"
            };
            updateButton.Location = new Point(groupBoxupdate.Width - updateButton.Width - 10, 60);
            laterButton.Location = new Point(updateButton.Location.X - laterButton.Width - 5, 60);
            downLabel.Location = new Point(laterButton.Location.X - downLabel.Width - 20, 60);
            groupBoxupdate.Controls.Add(updateButton);
            groupBoxupdate.Controls.Add(laterButton);
            groupBoxupdate.Controls.Add(downLabel);
            groupBoxupdate.Controls.Add(infoLabel);
            groupBoxupdate.Controls.Add(versionLabel);
            updateButton.Click += new EventHandler(UpdateButton_Click);
            laterButton.Click += new EventHandler(LaterButton_Click);
            void LaterButton_Click(object sender, EventArgs e)
            {
                groupBoxupdate.Dispose();
                Controls.Remove(groupBoxupdate);
                groupBox1.Enabled = true;
                groupBox4.Enabled = true;
            }
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            try
            {
                var request = WebRequest.Create("https://github.com/UndertakerBen/PorLibOfficeUpd/raw/main/Version.txt");
                var response = request.GetResponse();
                using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                {
                    var version = reader.ReadToEnd();
                    FileVersionInfo testm = FileVersionInfo.GetVersionInfo(applicationPath + "\\Portable Libre Office Updater.exe");
                    versionLabel.Text = testm.FileVersion + "  >>> " + version;
                    if (Convert.ToInt32(version.Replace(".", "")) > Convert.ToInt32(testm.FileVersion.Replace(".", "")))
                    {
                        Controls.Add(groupBoxupdate);
                        groupBox1.Enabled = false;
                        groupBox4.Enabled = false;
                    }
                    reader.Close();
                }
            }
            catch (Exception)
            {
                
            }
            void UpdateButton_Click(object sender, EventArgs e)
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                var request2 = WebRequest.Create("https://github.com/UndertakerBen/PorLibOfficeUpd/raw/main/Version.txt");
                var response2 = request2.GetResponse();
                using (StreamReader reader = new StreamReader(response2.GetResponseStream()))
                {
                    var version = reader.ReadToEnd();
                    reader.Close();
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    using (WebClient myWebClient2 = new WebClient())
                    {
                        myWebClient2.DownloadFile($"https://github.com/UndertakerBen/PorLibOfficeUpd/releases/download/v{version}/Portable.LibreOffice.Updater.v{version}.zip", @"Portable.LibreOffice.Updater.v" + version + ".zip");
                    }
                    System.IO.Compression.ZipFile.ExtractToDirectory(applicationPath + "\\Portable.LibreOffice.Updater.v" + version + ".zip", applicationPath + "\\Temp");
                    var files = Directory.GetFiles(applicationPath + "\\Bin");
                    foreach (string file in files)
                    {
                        if (File.Exists(applicationPath + "\\Temp\\Bin\\" + Path.GetFileName(file)) & File.Exists(applicationPath + "\\Bin\\" + Path.GetFileName(file)))
                        {
                            FileVersionInfo binLauncher = FileVersionInfo.GetVersionInfo(applicationPath + "\\Bin\\" + Path.GetFileName(file));
                            FileVersionInfo tempLauncher = FileVersionInfo.GetVersionInfo(applicationPath + "\\Temp\\Bin\\" + Path.GetFileName(file));
                            if (Convert.ToInt32(tempLauncher.FileVersion.Replace(".", "")) > Convert.ToInt32(binLauncher.FileVersion.Replace(".", "")))
                            {
                                File.Copy(applicationPath + "\\Temp\\Bin\\" + Path.GetFileName(file), applicationPath + "\\Bin\\" + Path.GetFileName(file), true);
                                if (File.Exists(applicationPath + "\\" + Path.GetFileName(file)))
                                {
                                    FileVersionInfo istLauncher = FileVersionInfo.GetVersionInfo(applicationPath + "\\" + Path.GetFileName(file));
                                    if (Convert.ToInt32(tempLauncher.FileVersion.Replace(".", "")) > Convert.ToInt32(istLauncher.FileVersion.Replace(".", "")))
                                    {
                                        File.Copy(applicationPath + "\\Temp\\Bin\\" + Path.GetFileName(file), applicationPath + "\\" + Path.GetFileName(file), true);
                                    }
                                }
                            }
                        }
                    }
                    File.AppendAllText(@"Update.cmd", "@echo off" + "\n" +
                        "timeout /t 1 /nobreak" + "\n" +
                        "copy /Y \"" + applicationPath + "\\Temp\\Portable Libre Office Updater.exe\" " + "\"" + applicationPath + "\\Portable Libre Office Updater.exe\"\n" +
                        "call cmd /c Start /b \"\" " + "\"" + applicationPath + "\\Portable Libre Office Updater.exe\"\n" +
                        "rd /s /q \"" + applicationPath + "\\Temp\"\n" +
                        "del /f /q \"" + applicationPath + "\\Update.cmd\" && exit\n" +
                        "exit\n");
                    File.Delete(applicationPath + "\\Portable.LibreOffice.Updater.v" + version + ".zip");
                    string arguments = " /c call Update.cmd";
                    Process process = new Process();
                    process.StartInfo.FileName = "cmd.exe";
                    process.StartInfo.Arguments = arguments;
                    process.Start();
                    Close();
                }
            }
        }

        private void InfoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FileVersionInfo updVersion = FileVersionInfo.GetVersionInfo(applicationPath + "\\Portable Libre Office Updater.exe");
            FileVersionInfo launcherVersion = FileVersionInfo.GetVersionInfo(applicationPath + "\\Bin\\Libre Office Portable Launcher.exe");
            MessageBox.Show("Updater Version - " + updVersion.FileVersion + "\nLauncher Version - " + launcherVersion.FileVersion, "Version Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
