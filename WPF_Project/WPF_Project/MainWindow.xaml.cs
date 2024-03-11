
using NAudio.Wave;
using NAudio.Wave.SampleProviders;
using System.Windows;
using System.Windows.Forms;

namespace WPF_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private WaveIn recorder;
        private BufferedWaveProvider bufferedWaveProvider;
        private SavingWaveProvider savingWaveProvider;
        private WaveOut player;
        SmbPitchShiftingSampleProvider pitch;
        float semitone = (float)Math.Pow(2, 1.0 / 12);
        float upOneTone;
        float downOneTone;

        public MainWindow()
        {
            InitializeComponent();
            upOneTone = semitone * semitone;
            downOneTone = 1.0f / upOneTone;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (tb_url.Text == "")
            {
                System.Windows.MessageBox.Show("Input url before start");
            }
            else
            {
                btnStart.IsEnabled = false;
                btnStop.IsEnabled = true;
                btnDownTone.IsEnabled = true;
                btnUpTone.IsEnabled = true;

                // set up the recorder
                recorder = new WaveIn();
                recorder.DataAvailable += RecorderOnDataAvailable;

                // set up our signal chain
                bufferedWaveProvider = new BufferedWaveProvider(recorder.WaveFormat);
                savingWaveProvider = new SavingWaveProvider(bufferedWaveProvider, tb_url.Text);

                // set up playback
                player = new WaveOut();
                pitch = new SmbPitchShiftingSampleProvider(savingWaveProvider.ToSampleProvider());
                pitch.PitchFactor = 1;
                lblTone.Content = pitch.PitchFactor.ToString();
                player.Init(pitch);

                // begin playback & record
                player.Play();
                recorder.StartRecording();
            }
        }

        private void RecorderOnDataAvailable(object sender, WaveInEventArgs waveInEventArgs)
        {
            bufferedWaveProvider.AddSamples(waveInEventArgs.Buffer, 0, waveInEventArgs.BytesRecorded);
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if(tb_url.Text == "")
            {
                System.Windows.MessageBox.Show("Input url before start");
            }
            else
            {
                btnStart.IsEnabled = true;
                btnStop.IsEnabled = false;
                btnDownTone.IsEnabled = false;
                btnUpTone.IsEnabled = false;

                // stop recording
                recorder.StopRecording();
                // stop playback
                player.Stop();
                // finalise the WAV file
                savingWaveProvider.Dispose();
            }
        }

        private void btnUpTone_Click(object sender, EventArgs e)
        {
            pitch.PitchFactor += 0.02f;
            lblTone.Content = pitch.PitchFactor.ToString();
        }

        private void btnDownTone_Click(object sender, EventArgs e)
        {
            pitch.PitchFactor -= 0.02f;
            lblTone.Content = pitch.PitchFactor.ToString();
        }

        private void btn_browser_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Wave File (*.wav)| *.wav;";
            if (open.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
            tb_url.Text = open.FileName;
        }
    }
}