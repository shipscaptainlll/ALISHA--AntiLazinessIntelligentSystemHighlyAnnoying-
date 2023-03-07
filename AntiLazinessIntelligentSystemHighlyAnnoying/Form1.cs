using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.Speech.Recognition;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

namespace AntiLazinessIntelligentSystemHighlyAnnoying
{
    public partial class Form1 : Form
    {
        Stopwatch stopWatch;
        SpeechRecognitionEngine speechRecognitionEngine;
        SpeechSynthesizer speechSynthesizer;
        string assistantName = "ALISHA";

        public Form1()
        {
            InitializeComponent();
            LoadSpeechRecognition();
            InitializeStopwatch();
            LoadSpeechSynthesizing();
            
        }

        void InitializeStopwatch()
        {
            stopWatch = new Stopwatch();
        }

        void LoadSpeechRecognition()
        {
            speechRecognitionEngine = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("en-US"));

            var choiceLibrary = GetChoiceLibrary();
            var grammarBuilder = new GrammarBuilder(choiceLibrary);
            var grammar = new Grammar(grammarBuilder);
            speechRecognitionEngine.LoadGrammarAsync(grammar);

            speechRecognitionEngine.SpeechRecognized +=
                  new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);

            speechRecognitionEngine.SetInputToDefaultAudioDevice();

            speechRecognitionEngine.RecognizeAsync(RecognizeMode.Multiple);
        }

        void LoadSpeechSynthesizing()
        {
            speechSynthesizer = new SpeechSynthesizer();

            speechSynthesizer.SelectVoice("Microsoft Zira Desktop");

            speechSynthesizer.Volume = 3;
        }

        Choices GetChoiceLibrary()
        {
            Choices myChoices = new Choices();

            myChoices.Add("hey " + assistantName);
            myChoices.Add("hi " + assistantName);
            myChoices.Add("hello " + assistantName);

            myChoices.Add(assistantName + " open excel");
            myChoices.Add(assistantName + " open google");
            myChoices.Add(assistantName + " open paint");
            myChoices.Add(assistantName + " start stopwatch");
            myChoices.Add(assistantName + " pause stopwatch");
            myChoices.Add(assistantName + " stop stopwatch");
            myChoices.Add(assistantName + " reset stopwatch");
            myChoices.Add(assistantName + " save progress");

            return myChoices;
        }


        void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            
            richTextBox1.Text = e.Result.Text;
            
            switch (e.Result.Text)
            {
                case "ALISHA open excel": 
                    Process.Start("C:/Program Files/Microsoft Office/root/Office16/EXCEL.EXE"); 
                    speechSynthesizer.SpeakAsync("opening excel programm"); 
                    break;
                case "ALISHA open google": 
                    Process.Start("https://www.google.com");
                    speechSynthesizer.SpeakAsync("opening google browser");
                    break;
                case "ALISHA open paint": 
                    Process.Start("C:/Windows/System32/mspaint.exe");
                    speechSynthesizer.SpeakAsync("opening paint programm");
                    break;
                case "ALISHA start stopwatch":
                    if (stopWatch.IsRunning) { speechSynthesizer.SpeakAsync("sorry but stopwatch are already running"); }
                    else { stopWatch.Start(); speechSynthesizer.SpeakAsync("turning on stopwatch"); }
                    break;
                case "ALISHA pause stopwatch": case "ALISHA stop stopwatch":
                    if (!stopWatch.IsRunning) { speechSynthesizer.SpeakAsync("sorry but stopwatch are already paused"); }
                    else { stopWatch.Stop(); speechSynthesizer.SpeakAsync("pausing stopwatch initiated"); }
                    break;
                case "ALISHA reset stopwatch":
                    stopWatch.Reset(); speechSynthesizer.SpeakAsync("reseting stopwatches");
                    break;
                case "ALISHA save progress":
                    speechSynthesizer.SpeakAsync("sorry but I don't know how to do it");
                    break;
            }
            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            stopWatch.Start();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            stopWatch.Stop();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            stopWatch.Reset();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.label1.Text = string.Format("{0:hh\\:mm\\:ss}", stopWatch.Elapsed);
        }

        void UsingExcel()
        {
            string excelFilePath = @"C:\Users\nikit\OneDrive\Gamedev\HelloThere.xlsx";
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook;
            Worksheet worksheet;

            workbook = excelApp.Workbooks.Open(excelFilePath);
            worksheet = workbook.ActiveSheet;
            Range cellRange = worksheet.UsedRange;
            int countRecords = cellRange.Rows.Count;
            int add = countRecords + 11;
            worksheet.Cells[add, 1] = "Total rows " + countRecords;
            //cellRange.Value = "Pizza!";
            //workbook.Save();
            //workbook.SaveAs(@"C:\Users\nikit\Music\hello.xlsx");
            workbook.Close(true, Type.Missing, Type.Missing);
            excelApp.Quit();
        }
    }
}
