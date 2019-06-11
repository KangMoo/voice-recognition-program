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
using System.Media;


using CommandLine;
using Google.Apis.Auth.OAuth2;
using Google.Cloud.Speech.V1;
using Grpc.Auth;
using System.Threading;
using System.Threading.Tasks;

using System.Text.RegularExpressions;

namespace SR
{

    public partial class Form1 : Form
    {

        // [START speech_transcribe_streaming_mic]
        static async Task<object> StreamingMicRecognizeAsync(double seconds)
        {
            if (NAudio.Wave.WaveIn.DeviceCount < 1)
            {
                gapiTxt = "No Microphone!";
                //Console.WriteLine("No microphone!");
                return -1;
            }
            var speech = SpeechClient.Create();
            var streamingCall = speech.StreamingRecognize();
            // Write the initial request with the config.
            await streamingCall.WriteAsync(
                new StreamingRecognizeRequest()
                {
                    StreamingConfig = new StreamingRecognitionConfig()
                    {
                        Config = new RecognitionConfig()
                        {
                            Encoding =
                            RecognitionConfig.Types.AudioEncoding.Linear16,
                            SampleRateHertz = 16000,
                            LanguageCode = "ko-KR",
                        },
                        InterimResults = true,
                    }
                });

            // Print responses as they arrive.
            Task printResponses = Task.Run(async () =>
            {
                while (await streamingCall.ResponseStream.MoveNext(
                    default(CancellationToken)))
                {
                    foreach (var result in streamingCall.ResponseStream
                        .Current.Results)
                    {
                        foreach (var alternative in result.Alternatives)
                        {
                            gapiTxt = alternative.Transcript;
                            //Console.WriteLine(alternative.Transcript);
                        }
                    }
                }
            });
            // Read from the microphone and stream to API.
            object writeLock = new object();
            bool writeMore = true;
            var waveIn = new NAudio.Wave.WaveInEvent();
            waveIn.DeviceNumber = 0;
            waveIn.WaveFormat = new NAudio.Wave.WaveFormat(16000, 1);
            waveIn.DataAvailable +=
                (object sender, NAudio.Wave.WaveInEventArgs args) =>
                {
                    lock (writeLock)
                    {
                        if (!writeMore) return;
                        streamingCall.WriteAsync(
                            new StreamingRecognizeRequest()
                            {
                                AudioContent = Google.Protobuf.ByteString
                                    .CopyFrom(args.Buffer, 0, args.BytesRecorded)
                            }).Wait();
                    }
                };
            waveIn.StartRecording();

            gapiTxtUpdate = true;
            _Speech_On.Play();
            //Console.WriteLine("Speak now.");
            await Task.Delay(TimeSpan.FromSeconds(seconds));
            // Stop recording and shut down.
            waveIn.StopRecording();
            lock (writeLock) writeMore = false;
            await streamingCall.WriteCompleteAsync();
            await printResponses;
            gapiTxtUpdate = false;
            _Speech_Sleep.Play();
            return 0;
        }
        // [END speech_transcribe_streaming_mic]

        static string gapiTxt = "";
        static bool gapiTxtUpdate = false;
        static SoundPlayer _Speech_On = new SoundPlayer(SR_Test.Properties.Resources.Speech_On);
        static SoundPlayer _Speech_Sleep = new SoundPlayer(SR_Test.Properties.Resources.Speech_Sleep);

        const double _recognizeTimer = 5;
        private List<orderClass> orderList;
        ProcessStartInfo cmd = new ProcessStartInfo();
        Process process = new Process();
        PShell ps = new PShell();
        bool isSTTActive;
        float sttTimer = 0;
        Grammar g;
        bool isInorderList = false;
        static string richboxtextsave;

        public Form1()
        {
            InitializeComponent();
            initTTS();
            orderList = new List<orderClass>();
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
            if (gapiTxtUpdate == true) return;

            sttTimer = 3;

            if (e.Result.Text == "썬데이" || e.Result.Text == "선데이")
            {
                richTextBox2.Text = "썬데이" + "\n" + richTextBox2.Text;
                inputmatch(sender, e);
            }
        }

        //
        private async void inputmatch(object sender, SpeechRecognizedEventArgs e)
        {
            gapiTxt = "";
            var task1 = await StreamingMicRecognizeAsync(_recognizeTimer);

            string inputString = gapiTxt;
            orderUnit oU = new orderUnit();
            bool brk = false;
            foreach (orderClass it in orderList)
            {
                foreach (orderUnit it2 in it.units)
                {
                    string temp = compareOrder(inputString, it2);
                    if (temp != "")
                    {
                        string result = doPowerShellProgram(temp);
                        string textOutput = makeOutputString(result, it2.textOutput, false);
                        string outVoice = makeOutputString(result, it2.voiceOutput, true);

                        richTextBox1.Text = textOutput + "\n" + richTextBox1.Text;
                        richTextBox1.Update();
                        tts.Speak(outVoice);
                        break;
                    }
                }
                if (brk == true) break;
            }

        }

        //명령어가 매칭이 맞을경우 PowerShell 명령어 생성후 리턴/ 아니면 ""리턴
        private string compareOrder(string inputString, orderUnit oU)
        {
            var stems = TwitterKoreanProcessorCS.Stem(TwitterKoreanProcessorCS.Tokenize(inputString));
            List<KoreanToken> inputTokens = new List<KoreanToken>();
            List<KoreanToken> orderSplit = new List<KoreanToken>();
            List<string> vars = new List<string>();

            //형태소 분석기를 통해 추려내기
            foreach (KoreanToken it in stems)
            {
                if (it.Pos == KoreanPos.Space) continue;
                inputTokens.Add(it);
            }

          stems = TwitterKoreanProcessorCS.Stem(TwitterKoreanProcessorCS.Tokenize(oU.input));
            foreach (KoreanToken it in stems)
            {
                if (it.Pos == KoreanPos.Space) continue;
                else if (it.Text[0] == '$')
                {
                    if (it.Text == "$Noun")
                        it.Pos = KoreanPos.Noun;
                    else if (it.Text == "$Verb")
                        it.Pos = KoreanPos.Verb;
                    else if (it.Text == "$Adjective")
                        it.Pos = KoreanPos.Adjective;
                    else if (it.Text == "$Adverb")
                        it.Pos = KoreanPos.Adverb;
                    else if (it.Text == "$Determiner")
                        it.Pos = KoreanPos.Determiner;
                    else if (it.Text == "$Exclamation")
                        it.Pos = KoreanPos.Exclamation;
                    else if (it.Text == "$Josa")
                        it.Pos = KoreanPos.Josa;
                    else if (it.Text == "$Eomi")
                        it.Pos = KoreanPos.Eomi;
                    else if (it.Text == "$PreEomi")
                        it.Pos = KoreanPos.PreEomi;
                    else if (it.Text == "$Conjunction")
                        it.Pos = KoreanPos.Conjunction;
                    else if (it.Text == "$NounPrefix")
                        it.Pos = KoreanPos.NounPrefix;
                    else if (it.Text == "$VerbPrefix")
                        it.Pos = KoreanPos.VerbPrefix;
                    else if (it.Text == "$Suffix")
                        it.Pos = KoreanPos.Suffix;
                    else if (it.Text == "$Unknown")
                        it.Pos = KoreanPos.Unknown;
                    else if (it.Text == "$Korean")
                        it.Pos = KoreanPos.Korean;
                    else if (it.Text == "$Foreign")
                        it.Pos = KoreanPos.Foreign;
                    else if (it.Text == "$Number")
                        it.Pos = KoreanPos.Number;
                    else if (it.Text == "$KoreanParticle")
                        it.Pos = KoreanPos.KoreanParticle;
                    else if (it.Text == "$Alpha")
                        it.Pos = KoreanPos.Alpha;
                    else if (it.Text == "$Punctuation")
                        it.Pos = KoreanPos.Punctuation;
                    else if (it.Text == "$Hashtag")
                        it.Pos = KoreanPos.Hashtag;
                    else if (it.Text == "$ScreenName")
                        it.Pos = KoreanPos.ScreenName;
                    else if (it.Text == "$Email")
                        it.Pos = KoreanPos.Email;
                    else if (it.Text == "$URL")
                        it.Pos = KoreanPos.URL;
                    else if (it.Text == "$CashTag")
                        it.Pos = KoreanPos.CashTag;
                }
                orderSplit.Add(it);
            }

            int count = 0;
            for (int i = 0; i < inputTokens.Count(); i++)
            {

                if (orderSplit[count].Pos == inputTokens[i].Pos && orderSplit[count].Text[0] == '$')
                {
                    vars.Add(inputTokens[i].Text);
                    count++;

                }
                else if (orderSplit[count].Text == inputTokens[i].Text)
                {
                    count++;
                }
                if (count >= orderSplit.Count()) break;
            }

            if (count >= orderSplit.Count())
            {
                return mkPSOrder(vars, oU);
            }
            return "";
        }

        //PowerShell 명령어 만드는 함수
        private string mkPSOrder(List<string> vars, orderUnit oU)
        {
            string test = oU.PSOrder;
            for(int i =0; i<vars.Count();i++)
            {
                test = test.Replace("$" + i.ToString(), vars[i]);
            }
            return test;
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

        //PowerShellProgram 실햄 함수
        private string doPowerShellProgram(string order)
        {
            string output = ps.RunScript(order);
            if (output.Count() >= 4)
            {
                string temp = output.Substring(output.Count() - 4, 4);
                if (output.Substring(output.Count() - 4, 4) == "\r\n\r\n")
                    output = output.Remove(output.Count() - 4, 4);
            }
            return output;
        }

        private string makeOutputString(string result, string outVoice,bool erasernt)
        {
            outVoice = outVoice.Replace("$R", result);

            if (erasernt)
            {
                outVoice = outVoice.Replace("\r", "");
                outVoice = outVoice.Replace("\n", "");
                outVoice = outVoice.Replace("\t", "");
            }
            return outVoice;
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

        //음성인식 활성화 버튼
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
                richTextBox1.Text = "음성인식이 비활성화되었습니다." + "\n" + richTextBox1.Text;
                //richTextBox2.Text = "";
                timer1.Enabled = false;

                _Speech_Sleep.Play();
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
                richTextBox1.Text = "음성인식이 활성화되었습니다." + "\n" + richTextBox1.Text;
                //richTextBox2.Text = "";
                timer1.Enabled = true;
                
                _Speech_On.Play();
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

        //NotActive -> Active
        private void button4_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null) return;
            int count = listBox1.SelectedIndices.Count;
            for (int i = 0; i < count; i++)
            {
                orderClass temp = (orderClass)listBox1.SelectedItem;
                temp.isInActive = true;
                listBox2.Items.Add(listBox1.SelectedItem);
                listBox1.Items.Remove(listBox1.SelectedItem);
            }
        }

        //Active -> NotActive
        private void button5_Click(object sender, EventArgs e)
        {

            if (listBox2.SelectedItem == null) return;

            int count = listBox2.SelectedIndices.Count;
            for (int i = 0; i < count; i++)
            {
                orderClass temp = (orderClass)listBox2.SelectedItem;
                temp.isInActive = false;
                listBox1.Items.Add(listBox2.SelectedItem);
                listBox2.Items.Remove(listBox2.SelectedItem);
            }
        }

        //Delete Order
        private void button2_Click(object sender, EventArgs e)
        {

            if (listBox1.SelectedItem != null)
            {
                orderClass temp = (orderClass)listBox1.SelectedItem;
                listBox1.Items.Remove(listBox1.SelectedItem);
                orderList.Remove(temp);
            }
            if (listBox2.SelectedItem != null)
            {
                orderClass temp = (orderClass)listBox2.SelectedItem;
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
        private void addOrder(string path)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(path);
            orderClass neworder;
            XmlNodeList root = xml.SelectNodes("orders");
            foreach (XmlNode it in root)
            {
                foreach (XmlNode it2 in it.SelectNodes("order"))
                {
                    neworder = new orderClass();
                    neworder.name = it2.Attributes["name"].Value;
                    neworder.description = it2.Attributes["description"].Value;
                    neworder.isInActive = false;
                    foreach (XmlNode it3 in it2.SelectNodes("orderUnit"))
                    {
                        orderUnit newUnit = new orderUnit();

                        //input
                        newUnit.input = it3.SelectSingleNode("input").InnerText;

                        //psorder
                        newUnit.PSOrder = it3.SelectSingleNode("PSOrder").InnerText;

                        newUnit.textOutput = it3.SelectSingleNode("textOutput").InnerText;

                        //voiceoutput
                        newUnit.voiceOutput = it3.SelectSingleNode("voiceOutput").InnerText;

                        
                        neworder.units.Add(newUnit);
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


        private void saveSettingAsXml(string path)
        {


            XmlDocument xml = new XmlDocument();
            XmlNode root = xml.CreateElement("orders");
            XmlNode order;

            foreach (orderClass it in orderList)
            {

                order = xml.CreateElement("order");
                XmlAttribute nameattribute = xml.CreateAttribute("name");
                nameattribute.Value = it.name;
                XmlAttribute descriptionattribute = xml.CreateAttribute("description");
                descriptionattribute.Value = it.description;

                order.Attributes.Append(nameattribute);
                order.Attributes.Append(descriptionattribute);

                foreach (orderUnit it2 in it.units)
                {
                    XmlNode orderUnit = xml.CreateElement("orderUnit");


                    //input
                    XmlNode input = xml.CreateElement("input");
                    input.InnerText = it2.input;
                    //output
                    XmlNode psorder = xml.CreateElement("PSOrder");
                    psorder.InnerText = it2.PSOrder;

                    XmlNode voiceOutput = xml.CreateElement("voiceOutput");
                    voiceOutput.InnerText = it2.voiceOutput;

                    XmlNode textOutput = xml.CreateElement("textOutput");
                    textOutput.InnerText = it2.textOutput;

                    orderUnit.AppendChild(input);
                    orderUnit.AppendChild(psorder);
                    orderUnit.AppendChild(textOutput);
                    orderUnit.AppendChild(voiceOutput);
                    

                    order.AppendChild(orderUnit);
                }
                root.AppendChild(order);
            }
            xml.AppendChild(root);
            xml.Save(path);
        }
        private void grammarSet()
        {
            string[] strs = new string[2];
            strs[0] = "썬데이";
            strs[1] = "선데이";
            GrammarBuilder gb = new GrammarBuilder(new Choices(strs));
            g = new Grammar(gb);
            initRS();
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        //Save order setting
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


            if (gapiTxtUpdate)
            {
                if (richboxtextsave == "")
                    richboxtextsave = richTextBox2.Text;

                richTextBox2.Text = gapiTxt + "\n" + richboxtextsave;
            }
            else if (richboxtextsave != "")
            {
                richTextBox2.Text = gapiTxt + "\n" + richboxtextsave;
                richboxtextsave = "";
            }

        }

        //description button
        private void button7_Click(object sender, EventArgs e)
        {
            string str = "";

            if (listBox1.SelectedItem != null)
            {
                orderClass temp = (orderClass)listBox1.SelectedItem;
                str += temp.description + "\n\n";
                str += "<실행 명령어>\n";
                foreach (orderUnit it in temp.units)
                {
                    str += it.voiceOutput + "\n";
                }
                MessageBox.Show(str);
                return;
            }
            else if (listBox2.SelectedItem != null)
            {
                orderClass temp = (orderClass)listBox2.SelectedItem;
                str += temp.description + "\n\n";
                str += "<실행 명령어>\n";
                foreach (orderUnit it in temp.units)
                {
                    str += it.input + "\n";
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

        private void TestButton_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox2.Clear();
        }

    }

    class orderUnit
    {
        public string input = "";
        public string PSOrder = "";
        public string voiceOutput = "";
        public string textOutput = "";
        public orderUnit() { }
        public orderUnit(string input, string PSOrder, string voiceOutput)
        {
            this.input = input;
            this.PSOrder = PSOrder;
            this.voiceOutput = voiceOutput;
        }
    }
    class orderClass
    {
        public string name = "";
        public string description = "";
        public bool isInActive = false;
        public List<orderUnit> units = new List<orderUnit>();
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
    /// <summary>
    /// GoolgeSpeech
    /// </summary>
    class Options
    {
        [Value(0, HelpText = "A path to a sound file.  Encoding must be "
            + "Linear16 with a sample rate of 16000.", Required = true)]
        public string FilePath { get; set; }
    }

    class StorageOptions
    {
        [Value(0, HelpText = "A path to a sound file. "
            + "Can be a local file path or a Google Cloud Storage path like "
            + "gs://my-bucket/my-object. "
            + "Encoding must be "
            + "Linear16 with a sample rate of 16000.", Required = true)]
        public string FilePath { get; set; }
    }

    [Verb("sync", HelpText = "Detects speech in an audio file.")]
    class SyncOptions : StorageOptions
    {
        [Option('w', HelpText = "Report the time offsets of individual words.")]
        public bool EnableWordTimeOffsets { get; set; }

        [Option('p', HelpText = "Add punctuation to the transcription.")]
        public bool EnableAutomaticPunctuation { get; set; }

        [Option('m', HelpText = "Select a transcription model.")]
        public String SelectModel { get; set; }

        [Option('c', HelpText = "Set number of channels")]
        public int NumberOfChannels { get; set; }
    }

    [Verb("with-context", HelpText = "Detects speech in an audio file."
        + " Add additional context on stdin.")]
    class OptionsWithContext : StorageOptions { }

    [Verb("async", HelpText = "Creates a job to detect speech in an audio "
        + "file, and waits for the job to complete.")]
    class AsyncOptions : StorageOptions
    {
        [Option('w', HelpText = "Report the time offsets of individual words.")]
        public bool EnableWordTimeOffsets { get; set; }
    }

    [Verb("sync-creds", HelpText = "Detects speech in an audio file.")]
    class SyncOptionsWithCreds
    {
        [Value(0, HelpText = "A path to a sound file.  Encoding must be "
            + "Linear16 with a sample rate of 16000.", Required = true)]
        public string FilePath { get; set; }

        [Value(1, HelpText = "Path to Google credentials json file.", Required = true)]
        public string CredentialsFilePath { get; set; }
    }

    [Verb("stream", HelpText = "Detects speech in an audio file by streaming "
        + "it to the Speech API.")]
    class StreamingOptions : Options { }

    [Verb("listen", HelpText = "Detects speech in a microphone input stream.")]
    class ListenOptions
    {
        [Value(0, HelpText = "Number of seconds to listen for.", Required = false)]
        public int Seconds { get; set; } = 3;
    }

    [Verb("rec", HelpText = "Detects speech in an audio file. Supports other file formats.")]
    class RecOptions : Options
    {
        [Option('b', Default = 16000, HelpText = "Sample rate in bits per second.")]
        public int BitRate { get; set; }

        [Option('e', Default = RecognitionConfig.Types.AudioEncoding.Linear16,
            HelpText = "Audio file encoding format.")]
        public RecognitionConfig.Types.AudioEncoding Encoding { get; set; }
    }

    public class Recognize
    {

        static object Rec(string filePath, int bitRate,
            RecognitionConfig.Types.AudioEncoding encoding)
        {
            var speech = SpeechClient.Create();
            var response = speech.Recognize(new RecognitionConfig()
            {
                Encoding = encoding,
                SampleRateHertz = bitRate,
                LanguageCode = "ko-KR",
            }, RecognitionAudio.FromFile(filePath));
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine(alternative.Transcript);
                }
            }
            return 0;
        }

        // [START speech_transcribe_sync]
        static object SyncRecognize(string filePath)
        {
            var speech = SpeechClient.Create();
            var response = speech.Recognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 16000,
                LanguageCode = "ko-KR",
            }, RecognitionAudio.FromFile(filePath));
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine(alternative.Transcript);
                }
            }
            return 0;
        }
        // [END speech_transcribe_sync]


        // [START speech_sync_recognize_words]
        static object SyncRecognizeWords(string filePath)
        {
            var speech = SpeechClient.Create();
            var response = speech.Recognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 16000,
                LanguageCode = "ko-KR",
                EnableWordTimeOffsets = true,
            }, RecognitionAudio.FromFile(filePath));
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine($"Transcript: { alternative.Transcript}");
                    Console.WriteLine("Word details:");
                    Console.WriteLine($" Word count:{alternative.Words.Count}");
                    foreach (var item in alternative.Words)
                    {
                        Console.WriteLine($"  {item.Word}");
                        Console.WriteLine($"    WordStartTime: {item.StartTime}");
                        Console.WriteLine($"    WordEndTime: {item.EndTime}");
                    }
                }
            }
            return 0;
        }
        // [END speech_sync_recognize_words]

        // [START speech_transcribe_auto_punctuation]
        static object SyncRecognizePunctuation(string filePath)
        {
            var speech = SpeechClient.Create();
            var response = speech.Recognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 8000,
                LanguageCode = "ko-KR",
                EnableAutomaticPunctuation = true,
            }, RecognitionAudio.FromFile(filePath));
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine(alternative.Transcript);
                }
            }
            return 0;
        }
        // [END speech_transcribe_auto_punctuation]

        // [START speech_transcribe_model_selection]
        static object SyncRecognizeModelSelection(string filePath, string model)
        {
            var speech = SpeechClient.Create();
            var response = speech.Recognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 16000,
                LanguageCode = "ko-KR",
                // The `model` value must be one of the following:
                // "video", "phone_call", "command_and_search", "default"
                Model = model
            }, RecognitionAudio.FromFile(filePath));
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine(alternative.Transcript);
                }
            }
            return 0;
        }
        // [END speech_transcribe_model_selection]

        // [START speech_transcribe_multichannel_beta]
        static object SyncRecognizeMultipleChannels(string filePath, int channelCount)
        {
            Console.WriteLine("Starting multi-channel");
            var speech = SpeechClient.Create();

            // Create transcription request
            var response = speech.Recognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                LanguageCode = "ko-KR",
                // Configure request to enable multiple channels
                EnableSeparateRecognitionPerChannel = true,
                AudioChannelCount = channelCount
            }, RecognitionAudio.FromFile(filePath));

            // Print out the results.
            foreach (var result in response.Results)
            {
                // There can be several transcripts for a chunk of audio.
                // Print out the first (most likely) one here.
                var alternative = result.Alternatives[0];
                Console.WriteLine($"Transcript: {alternative.Transcript}");
                Console.WriteLine($"Channel Tag: {result.ChannelTag}");
            }
            return 0;
        }
        // [END speech_transcribe_multichannel_beta]

        /// <summary>
        /// Reads a list of phrases from stdin.
        /// </summary>
        static List<string> ReadPhrases()
        {
            Console.Write("Reading phrases from stdin.  Finish with blank line.\n> ");
            var phrases = new List<string>();
            string line = Console.ReadLine();
            while (!string.IsNullOrWhiteSpace(line))
            {
                phrases.Add(line.Trim());
                Console.Write("> ");
                line = Console.ReadLine();
            }
            return phrases;
        }

        static object RecognizeWithContext(string filePath, IEnumerable<string> phrases)
        {
            var speech = SpeechClient.Create();
            var config = new RecognitionConfig()
            {
                SpeechContexts = { new SpeechContext() { Phrases = { phrases } } },
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 16000,
                LanguageCode = "ko-KR",
            };
            var audio = IsStorageUri(filePath) ?
                RecognitionAudio.FromStorageUri(filePath) :
                RecognitionAudio.FromFile(filePath);
            var response = speech.Recognize(config, audio);
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine(alternative.Transcript);
                }
            }
            return 0;
        }

        static object SyncRecognizeWithCredentials(string filePath, string credentialsFilePath)
        {
            GoogleCredential googleCredential;
            using (Stream m = new FileStream(credentialsFilePath, FileMode.Open))
                googleCredential = GoogleCredential.FromStream(m);
            var channel = new Grpc.Core.Channel(SpeechClient.DefaultEndpoint.Host,
                googleCredential.ToChannelCredentials());
            var speech = SpeechClient.Create(channel);
            var response = speech.Recognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 16000,
                LanguageCode = "ko-KR",
            }, RecognitionAudio.FromFile(filePath));
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine(alternative.Transcript);
                }
            }
            return 0;
        }

        // [START speech_transcribe_sync_gcs]
        static object SyncRecognizeGcs(string storageUri)
        {
            var speech = SpeechClient.Create();
            var response = speech.Recognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 16000,
                LanguageCode = "ko-KR",
            }, RecognitionAudio.FromStorageUri(storageUri));
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine(alternative.Transcript);
                }
            }
            return 0;
        }
        // [END speech_transcribe_sync_gcs]

        // [START speech_transcribe_async]
        static object LongRunningRecognize(string filePath)
        {
            var speech = SpeechClient.Create();
            var longOperation = speech.LongRunningRecognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 16000,
                LanguageCode = "ko-KR",
            }, RecognitionAudio.FromFile(filePath));
            longOperation = longOperation.PollUntilCompleted();
            var response = longOperation.Result;
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine(alternative.Transcript);
                }
            }
            return 0;
        }
        // [END speech_transcribe_async]

        // [START speech_transcribe_async_gcs]
        static object AsyncRecognizeGcs(string storageUri)
        {
            var speech = SpeechClient.Create();
            var longOperation = speech.LongRunningRecognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 16000,
                LanguageCode = "ko-KR",
            }, RecognitionAudio.FromStorageUri(storageUri));
            longOperation = longOperation.PollUntilCompleted();
            var response = longOperation.Result;
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine($"Transcript: { alternative.Transcript}");
                }
            }
            return 0;
        }
        // [END speech_transcribe_async_gcs]

        // [START speech_transcribe_async_word_time_offsets_gcs]
        static object AsyncRecognizeGcsWords(string storageUri)
        {
            var speech = SpeechClient.Create();
            var longOperation = speech.LongRunningRecognize(new RecognitionConfig()
            {
                Encoding = RecognitionConfig.Types.AudioEncoding.Linear16,
                SampleRateHertz = 16000,
                LanguageCode = "ko-KR",
                EnableWordTimeOffsets = true,
            }, RecognitionAudio.FromStorageUri(storageUri));
            longOperation = longOperation.PollUntilCompleted();
            var response = longOperation.Result;
            foreach (var result in response.Results)
            {
                foreach (var alternative in result.Alternatives)
                {
                    Console.WriteLine($"Transcript: { alternative.Transcript}");
                    Console.WriteLine("Word details:");
                    Console.WriteLine($" Word count:{alternative.Words.Count}");
                    foreach (var item in alternative.Words)
                    {
                        Console.WriteLine($"  {item.Word}");
                        Console.WriteLine($"    WordStartTime: {item.StartTime}");
                        Console.WriteLine($"    WordEndTime: {item.EndTime}");
                    }
                }
            }
            return 0;
        }
        // [END speech_transcribe_async_word_time_offsets_gcs]

        /// <summary>
        /// Stream the content of the file to the API in 32kb chunks.
        /// </summary>
        // [START speech_transcribe_streaming]
        static async Task<object> StreamingRecognizeAsync(string filePath)
        {
            var speech = SpeechClient.Create();
            var streamingCall = speech.StreamingRecognize();
            // Write the initial request with the config.
            await streamingCall.WriteAsync(
                new StreamingRecognizeRequest()
                {
                    StreamingConfig = new StreamingRecognitionConfig()
                    {
                        Config = new RecognitionConfig()
                        {
                            Encoding =
                            RecognitionConfig.Types.AudioEncoding.Linear16,
                            SampleRateHertz = 16000,
                            LanguageCode = "ko-KR",
                        },
                        InterimResults = true,
                    }
                });
            // Print responses as they arrive.
            Task printResponses = Task.Run(async () =>
            {
                while (await streamingCall.ResponseStream.MoveNext(
                    default(CancellationToken)))
                {
                    foreach (var result in streamingCall.ResponseStream
                        .Current.Results)
                    {
                        foreach (var alternative in result.Alternatives)
                        {
                            Console.WriteLine(alternative.Transcript);
                        }
                    }
                }
            });
            // Stream the file content to the API.  Write 2 32kb chunks per
            // second.
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open))
            {
                var buffer = new byte[32 * 1024];
                int bytesRead;
                while ((bytesRead = await fileStream.ReadAsync(
                    buffer, 0, buffer.Length)) > 0)
                {
                    await streamingCall.WriteAsync(
                        new StreamingRecognizeRequest()
                        {
                            AudioContent = Google.Protobuf.ByteString
                            .CopyFrom(buffer, 0, bytesRead),
                        });
                    await Task.Delay(500);
                };
            }
            await streamingCall.WriteCompleteAsync();
            await printResponses;
            return 0;
        }
        // [END speech_transcribe_streaming]



        static bool IsStorageUri(string s) => s.Substring(0, 4).ToLower() == "gs:/";
    }
}