using System.Speech.Recognition;
using System.Windows;


namespace WPF_Project
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private SpeechRecognitionEngine recognizer;
        private string recognizedText;
        

        public MainWindow()
        {
            InitializeComponent();
        }


        private void Grid_Loaded(object sender, RoutedEventArgs e)
        {
           
        }

        private void recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            tb_Text.Text = tb_Text.Text +  e.Result.Text.ToString() + Environment.NewLine;
            recognizedText = e.Result.Text;
            

        }


        private void btn_Speech_Click(object sender, RoutedEventArgs e)
        {
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
            recognizer.Dispose();

            tb_Text.Clear();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ControlVoice controlVoiceWindow = new ControlVoice();
            controlVoiceWindow.Show();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ControlVoice controlVoice = new ControlVoice();
            controlVoice.ShowDialog();
        }
    }
}