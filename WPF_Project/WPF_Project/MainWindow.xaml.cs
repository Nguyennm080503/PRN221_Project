using System.Speech.Recognition;
using System.Windows;
using System.Windows.Forms;

namespace WPF_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SpeechRecognitionEngine recognizer;
        private string recognizedText;
        //private WaveIn recorder;
        //private BufferedWaveProvider bufferedWaveProvider;
        //private SavingWaveProvider savingWaveProvider;
        //private WaveOut player;
        //SmbPitchShiftingSampleProvider pitch;
        //float semitone = (float)Math.Pow(2, 1.0 / 12);
        //float upOneTone;
        //float downOneTone;

        public MainWindow()
        {
            InitializeComponent();
            //upOneTone = semitone * semitone;
            //downOneTone = 1.0f / upOneTone;
        }


        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
           
        }

        private void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            tb_Text.Text = tb_Text.Text +  e.Result.Text.ToString() + Environment.NewLine;
            recognizedText = e.Result.Text;
            

        }

        
        //private void Button_Click(object sender, RoutedEventArgs e)
        //{
        //    if (tb_url.Text == "")
        //    {
        //        System.Windows.MessageBox.Show("Input url before start");
        //    }
        //    else
        //    {
        //        btnStart.IsEnabled = false;
        //        btnStop.IsEnabled = true;
        //        btnDownTone.IsEnabled = true;
        //        btnUpTone.IsEnabled = true;

        //        // set up the recorder
        //        recorder = new WaveIn();
        //        recorder.DataAvailable += RecorderOnDataAvailable;

        //        // set up our signal chain
        //        bufferedWaveProvider = new BufferedWaveProvider(recorder.WaveFormat);
        //        savingWaveProvider = new SavingWaveProvider(bufferedWaveProvider, tb_url.Text);

        //        // set up playback
        //        player = new WaveOut();
        //        pitch = new SmbPitchShiftingSampleProvider(savingWaveProvider.ToSampleProvider());
        //        pitch.PitchFactor = 1;
        //        lblTone.Content = pitch.PitchFactor.ToString();
        //        player.Init(pitch);

        //        // begin playback & record
        //        player.Play();
        //        recorder.StartRecording();
        //    }
        //}

        //private void RecorderOnDataAvailable(object sender, WaveInEventArgs waveInEventArgs)
        //{
        //    bufferedWaveProvider.AddSamples(waveInEventArgs.Buffer, 0, waveInEventArgs.BytesRecorded);
        //}

        //private void btnStop_Click(object sender, EventArgs e)
        //{
        //    if(tb_url.Text == "")
        //    {
        //        System.Windows.MessageBox.Show("Input url before start");
        //    }
        //    else
        //    {
        //        btnStart.IsEnabled = true;
        //        btnStop.IsEnabled = false;
        //        btnDownTone.IsEnabled = false;
        //        btnUpTone.IsEnabled = false;

        //        // stop recording
        //        recorder.StopRecording();
        //        // stop playback
        //        player.Stop();
        //        // finalise the WAV file
        //        savingWaveProvider.Dispose();
        //    }
        //}

        //private void btnUpTone_Click(object sender, EventArgs e)
        //{
        //    pitch.PitchFactor += 0.02f;
        //    lblTone.Content = pitch.PitchFactor.ToString();
        //}

        //private void btnDownTone_Click(object sender, EventArgs e)
        //{
        //    pitch.PitchFactor -= 0.02f;
        //    lblTone.Content = pitch.PitchFactor.ToString();
        //}

        //private void btn_browser_Click(object sender, RoutedEventArgs e)
        //{
        //    OpenFileDialog open = new OpenFileDialog();
        //    open.Filter = "Wave File (*.wav)| *.wav;";
        //    if (open.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
        //    tb_url.Text = open.FileName;
        //}


        private void btn_Speech_Click(object sender, RoutedEventArgs e)
        {
            //btn_Speech.IsEnabled = false;
            //btn_Stop.IsEnabled = true;
            //btn_Reset.IsEnabled = true;
            try 
            {
                recognizer = new SpeechRecognitionEngine();
                recognizer.SetInputToDefaultAudioDevice();
                Grammar grammar = new DictationGrammar();
                recognizer.LoadGrammar(grammar);
                recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(recognizer_SpeechRecognized);
                recognizer.RecognizeAsync(RecognizeMode.Multiple);
                
            }
            catch(Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Error");
            }
            
        }

        private void btn_Stop_Click(object sender, RoutedEventArgs e)
        {
            if (recognizer != null)
            {
                recognizer.RecognizeAsyncCancel();
                recognizer.Dispose();
            }
            if (recognizedText == "Show product lists")
            {
                ControlVoice controlVoiceWindow = new ControlVoice();
                controlVoiceWindow.Show();
            }
        }

        private void btn_Reset_Click(object sender, RoutedEventArgs e)
        {
            if (recognizer != null)
            {
                recognizer.RecognizeAsyncCancel();
                recognizer.Dispose();
            }
            tb_Text.Clear();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ControlVoice controlVoiceWindow = new ControlVoice();
            controlVoiceWindow.Show();
        }
    }
}