using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

using Microsoft.Speech.Recognition;
using Microsoft.Speech.Synthesis;
using Microsoft.Speech;

using Moda.Korean.TwitterKoreanProcessorCS;
using System.Xml;
using System.IO;
using System.Management.Automation;
namespace SR
{
    public partial class Form1 : Form
    {

        private List<orderStruct> orderList;
        private orderStruct ost = new orderStruct();
        ProcessStartInfo cmd = new ProcessStartInfo();
        Process process = new Process();
        PowerShell ps = PowerShell.Create();
        bool isSTTActive;
        int sttTimer = 0;
        Grammar g;

        public Form1()
        {
            InitializeComponent();
            initTTS();
            orderList = new List<orderStruct>();
            isSTTActive = false;
            setting_set();
            cmdSetting();
        }
        public void initRS()
        {
            try
            {
                SpeechRecognitionEngine sre = new SpeechRecognitionEngine(new CultureInfo("ko-KR"));
                sre.LoadGrammar(g);
                sre.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(sre_SpeechRecognized);
                sre.SetInputToDefaultAudioDevice();
                sre.RecognizeAsync(RecognizeMode.Multiple);
            }
            catch (Exception e)
            {
                richTextBox1.Text = "init RS Error : " + e.ToString();
            }
        }

        SpeechSynthesizer tts;

        public void initTTS()
        {
            try
            {
                tts = new SpeechSynthesizer();
                tts.SelectVoice("Microsoft Server Speech Text to Speech Voice (ko-KR, Heami)");
                tts.SetOutputToDefaultAudioDevice();
                tts.Volume = 100;
            }
            catch (Exception e)
            {
                richTextBox1.Text = "init TTS Error : " + e.ToString();
            }
        }

        void sre_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            if (isSTTActive == false && sttTimer > 0) return;
            if (tts.State != SynthesizerState.Ready) return;
            sttTimer = 3;
            for (int i = 0; i < orderList.Count; i++)
            {
                if (!orderList[i].isInActive) continue;
                for (int j = 0; j < orderList[i].voiceInputs.Count; j++)
                {
                    if (e.Result.Text == orderList[i].voiceInputs[j])
                    {
                        richTextBox1.Text = orderList[i].voiceOutput + "\n" + richTextBox1.Text;
                        richTextBox2.Text = e.Result.Text + "\n" + richTextBox2.Text;
                        tts.SpeakAsync(orderList[i].voiceOutput);
                        //doPowerShellProgram(orderList[i].cmdString);
                        cmdOrder(orderList[i].cmdString, orderList[i].args);
                        break;
                    }
                }
            }
            //switch (e.Result.Text)
            //{
            //    case "송환의":
            //        tts.SpeakAsync("하하하하");
            //        break;
            //
            //    case "컴퓨터":
            //        tts.SpeakAsync("네 컴퓨터입니다");
            //        break;
            //
            //    case "안녕":
            //        tts.SpeakAsync("반갑습니다 음성인식 테스터입니다");
            //        break;
            //
            //    case "종료":
            //        {
            //            tts.Speak("프로그램을 종료합니다");
            //            Application.Exit();
            //            break;
            //        }
            //
            //    case "계산기":
            //        {
            //            tts.SpeakAsync("계산기를 실행합니다");
            //            doProgram("c:\\windows\\system32\\calc.exe", "");
            //            break;
            //        }
            //
            //    case "메모장":
            //        {
            //            tts.SpeakAsync("메모장을 실행합니다");
            //            doProgram("c:\\windows\\system32\\notepad.exe", "");
            //            break;
            //        }
            //
            //    case "콘솔":
            //        {
            //            tts.SpeakAsync("콘솔을 실행합니다");
            //            doProgram("c:\\windows\\system32\\cmd.exe", "");
            //            break;
            //        }
            //
            //    case "그림판":
            //        {
            //            tts.SpeakAsync("그림판을 실행합니다");
            //            doProgram("c:\\windows\\system32\\mspaint.exe", "");
            //            break;
            //        }
            //
            //    case "계산기 닫기":
            //        {
            //            tts.SpeakAsync("계산기를 종료합니다");
            //            closeProcess("calc");
            //            break;
            //        }
            //}
        }

        // 프로세스 실행
        private static void doProgram(string filename, string arg)
        {
            ProcessStartInfo psi;
            if (arg.Length != 0)
                psi = new ProcessStartInfo(filename, arg);
            else
                psi = new ProcessStartInfo(filename);
            Process.Start(psi);
        }
        private void cmdSetting()
        {
            cmd.FileName = @"cmd";
            cmd.WindowStyle = ProcessWindowStyle.Normal;
            cmd.CreateNoWindow = true;

            cmd.UseShellExecute = false;
            cmd.RedirectStandardOutput = true;
            cmd.RedirectStandardInput = true;
            cmd.RedirectStandardError = true;
            process.EnableRaisingEvents = false;
            process.StartInfo = cmd;
            process.Start();
        }
        private void cmdOrder(string cmdstring, string arg)
        {

            process.StandardInput.Write(cmdstring + Environment.NewLine);
            //process.StandardInput.Close();
            //string result = process.StandardOutput.ReadToEnd();
            //process.WaitForExit();
            //process.Close();
        }

        private void doPowerShellProgram(string order)
        {
            ps.AddCommand(order);
            ps.Invoke();
        }

        // 프로세스 종료
        private static void closeProcess(string filename)
        {
            Process[] myProcesses;
            // Returns array containing all instances of Notepad.
            myProcesses = Process.GetProcessesByName(filename);
            foreach (Process myProcess in myProcesses)
            {
                myProcess.CloseMainWindow();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            XmlDocument xml = new XmlDocument();
            openFileDialog1.Filter = "XML 문서(*.xml)|*.xml|모든파일(*.*)|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                if (openFileDialog1.FileNames != null)
                {
                    foreach (string path in openFileDialog1.FileNames)
                    {
                        addOrder(path);
                    }

                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (isSTTActive)
            {
                isSTTActive = false;
                button3.Text = "음성인식 활성화";
                button1.Enabled = true;
                button2.Enabled = true;
                button4.Enabled = true;
                button5.Enabled = true;
                button6.Enabled = true;
                richTextBox1.Text = "음성인식이 비활성화되었습니다.";
                richTextBox2.Text = "";
                timer1.Enabled = false;
            }
            else
            {
                isSTTActive = true;
                button3.Text = "음성인식 비활성화";
                button1.Enabled = false;
                button2.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
                button6.Enabled = false;
                richTextBox1.Text = "음성인식이 활성화되었습니다.";
                richTextBox2.Text = "";
                timer1.Enabled = true;
            }
            if (orderList.Count > 0)
                grammarSet();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null) return;
            int count = listBox1.SelectedIndices.Count;
            for (int i = 0; i < count; i++)
            {
                orderStruct temp = (orderStruct)listBox1.SelectedItem;
                temp.isInActive = true;
                listBox2.Items.Add(listBox1.SelectedItem);
                listBox1.Items.Remove(listBox1.SelectedItem);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            if (listBox2.SelectedItem == null) return;

            int count = listBox2.SelectedIndices.Count;
            for (int i = 0; i < count; i++)
            {
                orderStruct temp = (orderStruct)listBox2.SelectedItem;
                temp.isInActive = false;
                listBox1.Items.Add(listBox2.SelectedItem);
                listBox2.Items.Remove(listBox2.SelectedItem);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (listBox1.SelectedItem != null)
            {
                orderStruct temp = (orderStruct)listBox1.SelectedItem;
                listBox1.Items.Remove(listBox1.SelectedItem);
                orderList.Remove(temp);
            }
            if (listBox2.SelectedItem != null)
            {
                orderStruct temp = (orderStruct)listBox2.SelectedItem;
                listBox2.Items.Remove(listBox2.SelectedItem);
                orderList.Remove(temp);
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void setting_set()
        {
            if (!isSTTActive)
            {
                button3.Text = "음성인식 활성화";
                button1.Enabled = true;
                button2.Enabled = true;
                button4.Enabled = true;
                button5.Enabled = true;
                richTextBox1.Text = "음성인식이 비활성화되었습니다.";
                richTextBox2.Text = "";
            }
            else
            {
                button3.Text = "음성인식 비활성화";
                button1.Enabled = false;
                button2.Enabled = false;
                button4.Enabled = false;
                button5.Enabled = false;
                richTextBox1.Text = "음성인식이 활성화되었습니다.";
                richTextBox2.Text = "";
            }
            //grammarSet();
        }
        private string orderRename(string str)
        {
            int count = 0;
            string temp = str;
            foreach (orderStruct it in orderList)
            {
                if (it.name == temp)
                {
                    count++;
                    temp = str + "(" + count.ToString() + ")";
                }
            }
            return temp;
        }
        private void addOrder(string path)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(path);
            orderStruct neworder;
            XmlNodeList root = xml.SelectNodes("orders");
            foreach (XmlNode it in root)
            {
                foreach (XmlNode it2 in it.SelectNodes("order"))
                {
                    neworder = new orderStruct();
                    List<string> test = new List<string>();
                    neworder.voiceInputs = new List<string>();
                    neworder.name = orderRename(it2.SelectSingleNode("ordername").InnerText);
                    neworder.voiceOutput = it2.SelectSingleNode("voiceOutput").InnerText;
                    neworder.cmdString = it2.SelectSingleNode("cmdString").InnerText;
                    neworder.description = it2.SelectSingleNode("description").InnerText;
                    neworder.args = it2.SelectSingleNode("cmdArgs").InnerText;
                    foreach (XmlNode it3 in it2.SelectSingleNode("inputOrders").SelectNodes("inputOrder"))
                    {
                        neworder.voiceInputs.Add(it3.InnerText);
                    }
                    orderList.Add(neworder);
                    listBox1.Items.Add(orderList[orderList.Count - 1]);
                }
            }
        }
        private void deleteOrder(string name)
        {
            string ordername = "";
            if (listBox1.SelectedItem != null)
            {
                ordername = listBox1.SelectedItem.ToString();
            }
            if (listBox1.SelectedItem == null) return;
            listBox2.Items.Add(listBox1.SelectedItem);
            listBox1.Items.Remove(listBox1.SelectedItem);
            if (listBox2.SelectedItem == null) return;
            listBox1.Items.Add(listBox2.SelectedItem);
            listBox2.Items.Remove(listBox2.SelectedItem);
        }
        private orderStruct findOrder(string name)
        {
            int i;
            for (i = 0; i < orderList.Count; i++)
            {
                if (name == orderList[i].name)
                {
                    break;
                }
            }
            return orderList[i];
        }
        private bool isThereOrder(string name)
        {
            foreach (orderStruct it in orderList)
            {
                if (name == it.name)
                    return true;
            }
            return false;
        }

        private void saveSettingAsXml(string path)
        {
            XmlDocument xml = new XmlDocument();
            XmlNode root = xml.CreateElement("orders");
            XmlNode order = xml.CreateElement("order"); ;
            foreach (orderStruct it in orderList)
            {
                order = xml.CreateElement("order");
                XmlNode ordername = xml.CreateElement("ordername");
                ordername.InnerText = it.name;
                XmlNode voiceoutput = xml.CreateElement("voiceOutput");
                voiceoutput.InnerText = it.voiceOutput;
                XmlNode cmdstring = xml.CreateElement("cmdString");
                cmdstring.InnerText = it.cmdString;
                XmlNode args = xml.CreateElement("cmdArgs");
                args.InnerText = it.args;
                XmlNode description = xml.CreateElement("description");
                description.InnerText = it.description;
                XmlNode inputorders = xml.CreateElement("inputOrders");
                foreach (string it2 in it.voiceInputs)
                {
                    XmlNode inputorer = xml.CreateElement("inputOrder");
                    inputorer.InnerText = it2;
                    inputorders.AppendChild(inputorer);
                }
                order.AppendChild(ordername);
                order.AppendChild(voiceoutput);
                order.AppendChild(cmdstring);
                order.AppendChild(args);
                order.AppendChild(description);
                order.AppendChild(inputorders);
                root.AppendChild(order);
            }
            xml.AppendChild(root);
            xml.Save(path);
        }
        private void grammarSet()
        {
            int size = 0;
            foreach (orderStruct it in orderList)
            {
                size += it.voiceInputs.Count;
            }
            string[] strs = new string[size];
            int count = 0;
            foreach (orderStruct it in orderList)
            {
                foreach (string it2 in it.voiceInputs)
                {
                    strs[count] = it2;
                    count++;
                }
            }

            GrammarBuilder gb = new GrammarBuilder(new Choices(strs));
            g = new Grammar(gb);
            initRS();
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (orderList.Count == 0)
            {
                MessageBox.Show("저장할 명령이 없습니다.");
                return;
            }
            saveFileDialog1.Filter = "XML 문서(*.xml)|*.xml|모든파일(*.*)|*.*";
            string path;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if ((path = saveFileDialog1.FileName) != null)
                {
                    saveSettingAsXml(path);
                }
            }


        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            sttTimer--;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string str = "";

            if (listBox1.SelectedItem != null)
            {
                orderStruct temp = (orderStruct)listBox1.SelectedItem;
                str += temp.description + "\n\n";
                str += "<실행 명령어>\n";
                foreach (string it in temp.voiceInputs)
                {
                    str += it + "\n";
                }
                MessageBox.Show(str);
                return;
            }
            else if (listBox2.SelectedItem != null)
            {
                orderStruct temp = (orderStruct)listBox2.SelectedItem;
                str += temp.description + "\n\n";
                str += "<실행 명령어>\n";
                foreach (string it in temp.voiceInputs)
                {
                    str += it + "\n";
                }
                MessageBox.Show(str);
                return;
            }
            MessageBox.Show("선택된 명령이 없습니다.");
            return;
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }

    class orderStruct
    {
        public List<string> voiceInputs;
        public string voiceOutput;
        public string cmdString;
        public string description;
        public string name;
        public string args;
        public bool isInActive;
        public orderStruct(int x = 0)
        {
            voiceInputs = new List<string>();
            voiceOutput = null;
            cmdString = null;
            description = null;
            name = null;
            isInActive = false;
            args = null;
        }
        public override String ToString()
        {
            return this.name;
        }
        public string ListBoxDisplayValue
        {
            get
            {
                return this.name;
            }
        }
        public string DisplayValue
        {
            get
            {
                return this.description;
            }
        }
    }
}