using Microsoft.CognitiveServices.SpeechRecognition;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ArduinoVoiceControlLuis
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        static string appId = "2bdee1c9-82a5-48d5-9d44-a4703deca30e";
        static string luisSubscriptionKey = "39001e52b8ab4854ab752211b97ec933";
        static bool _continue;
        static SerialPort _serialPort;
        private readonly BackgroundWorker worker = new BackgroundWorker();
        private int speed;

        private List<LightMap> lightMaps = new List<LightMap>();

        private MicrophoneRecognitionClient micClient;

        public MainWindow()
        {
            InitializeComponent();

            //create a map of IO

            LightMap _lightMap = new LightMap();
            lightMaps = new List<LightMap>()
            {
                new LightMap()
                {

                    name = "Red",
                    keywords = new List<string>()
                    {
                        "energy", "war", "danger", "strength", "power","determination", "passion", "desire", "love", "christmas", "shy"
                    },
                    io = 1
                },
                new LightMap()
                {
                    name = "Amber",
                    keywords = new List<string>()
                    {
                        "freshness", "happiness", "positivity", "clarity", "energy", "optimism", "enlightenment", "remembrance", "intellect", "honor", "loyalty", "joy"
                    },
                    io = 2
                },
                new LightMap()
                {
                    name = "Green",
                    keywords = new List<string>()
                    {
                        "life","renewal","nature","growth","harmony","fertility","natural"
                    },
                    io = 3
                },

            };

            worker.DoWork += worker_DoWork;
            setSpeed();

            string name;
            string message;
            StringComparer stringComparer = StringComparer.OrdinalIgnoreCase;

            // Create a new SerialPort object with default settings.
            _serialPort = new SerialPort();

            // Allow the user to set the appropriate properties.
            _serialPort.PortName = "COM3";
            _serialPort.BaudRate = _serialPort.BaudRate;
            _serialPort.Parity = _serialPort.Parity;
            _serialPort.DataBits = _serialPort.DataBits;
            _serialPort.StopBits = _serialPort.StopBits;
            _serialPort.Handshake = _serialPort.Handshake;

            // Set the read/write timeouts
            _serialPort.ReadTimeout = 500;
            _serialPort.WriteTimeout = 500;

        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            TurnOnLight();
        }

        private void btnConnect_Click(object sender, RoutedEventArgs e)
        {

            btnConnect.IsEnabled = false;
            btnDisconnect.IsEnabled = true;
            btnStart.IsEnabled = true;
            btnSend.IsEnabled = true;

            _serialPort.Open();
            _continue = true;
        }

        private void btnDisconnect_Click(object sender, RoutedEventArgs e)
        {

            _continue = false;
            btnConnect.IsEnabled = true;
            btnDisconnect.IsEnabled = false;
            btnStart.IsEnabled = false;
            btnSend.IsEnabled = false;

            isListening = false;
            this.micClient.EndMicAndRecognition();

            _serialPort.Close();
        }

        private void btnSend_Click(object sender, RoutedEventArgs e)
        {
            if (_continue)
            {
                _continue = false;
                btnSend.Content = "Demo Start";
            }
            else
            {
                _continue = true;
                worker.RunWorkerAsync();
                btnSend.Content = "Demo Stop";
            }

        }

        public void TurnOnLight()
        {
            int i = 1;
            while (_continue)
            {
                _serialPort.Write("ON:"+i);
                System.Threading.Thread.Sleep(speed);
                _serialPort.Write("OFF:" + i);
                System.Threading.Thread.Sleep(speed);

                i++;
                if (i == 4)
                {
                    i = 1;
                }

            }
        }
        public void TurnOnLight(string color)
        {
            for (int i = 0; i < lightMaps.Count; i++)
            {
                if (color.ToUpper().Contains(lightMaps[i].name.ToUpper()))
                {
                    _serialPort.WriteLine(lightMaps[i].io.ToString());
                    Console.WriteLine(lightMaps[i].name + ">>>" + lightMaps[i].io);
                    System.Threading.Thread.Sleep(speed);
                }
            }
        }
        // Display Port values and prompt user to enter a port.
        public static string SetPortName(string defaultPortName)
        {
            string portName = defaultPortName;

            Console.WriteLine("Available Ports:");
            foreach (string s in SerialPort.GetPortNames())
            {
                Console.WriteLine("   {0}", s);
            }

            //Console.Write("Enter COM port value (Default: {0}): ", defaultPortName);
            //portName = Console.ReadLine();

            if (portName == "" || !(portName.ToLower()).StartsWith("com"))
            {
                portName = defaultPortName;
            }
            return portName;
        }
        // Display BaudRate values and prompt user to enter a value.
        public static int SetPortBaudRate(int defaultPortBaudRate)
        {
            string baudRate;

            Console.Write("Baud Rate(default:{0}): ", defaultPortBaudRate);
            baudRate = Console.ReadLine();

            if (baudRate == "")
            {
                baudRate = defaultPortBaudRate.ToString();
            }

            return int.Parse(baudRate);
        }

        // Display PortParity values and prompt user to enter a value.
        public static Parity SetPortParity(Parity defaultPortParity)
        {
            string parity;

            Console.WriteLine("Available Parity options:");
            foreach (string s in Enum.GetNames(typeof(Parity)))
            {
                Console.WriteLine("   {0}", s);
            }

            Console.Write("Enter Parity value (Default: {0}):", defaultPortParity.ToString(), true);
            parity = Console.ReadLine();

            if (parity == "")
            {
                parity = defaultPortParity.ToString();
            }

            return (Parity)Enum.Parse(typeof(Parity), parity, true);
        }
        // Display DataBits values and prompt user to enter a value.
        public static int SetPortDataBits(int defaultPortDataBits)
        {
            string dataBits;

            Console.Write("Enter DataBits value (Default: {0}): ", defaultPortDataBits);
            dataBits = Console.ReadLine();

            if (dataBits == "")
            {
                dataBits = defaultPortDataBits.ToString();
            }

            return int.Parse(dataBits.ToUpperInvariant());
        }

        // Display StopBits values and prompt user to enter a value.
        public static StopBits SetPortStopBits(StopBits defaultPortStopBits)
        {
            string stopBits;

            Console.WriteLine("Available StopBits options:");
            foreach (string s in Enum.GetNames(typeof(StopBits)))
            {
                Console.WriteLine("   {0}", s);
            }

            Console.Write("Enter StopBits value (None is not supported and \n" +
             "raises an ArgumentOutOfRangeException. \n (Default: {0}):", defaultPortStopBits.ToString());
            stopBits = Console.ReadLine();

            if (stopBits == "")
            {
                stopBits = defaultPortStopBits.ToString();
            }

            return (StopBits)Enum.Parse(typeof(StopBits), stopBits, true);
        }
        public static Handshake SetPortHandshake(Handshake defaultPortHandshake)
        {
            string handshake;

            Console.WriteLine("Available Handshake options:");
            foreach (string s in Enum.GetNames(typeof(Handshake)))
            {
                Console.WriteLine("   {0}", s);
            }

            Console.Write("Enter Handshake value (Default: {0}):", defaultPortHandshake.ToString());
            handshake = Console.ReadLine();

            if (handshake == "")
            {
                handshake = defaultPortHandshake.ToString();
            }

            return (Handshake)Enum.Parse(typeof(Handshake), handshake, true);
        }

        private void sliderSpeed_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            setSpeed();
        }
        private void setSpeed()
        {

            speed = (int)sliderSpeed.Value;
            lblSpeed.Text = speed.ToString();
        }


        #region 
        public bool IsMicrophoneClientShortPhrase { get; set; }
        public bool IsMicrophoneClientDictation { get; set; }
        public bool IsMicrophoneClientWithIntent { get; set; }
        private SpeechRecognitionMode Mode
        {
            get
            {
                if (this.IsMicrophoneClientDictation)
                {
                    return SpeechRecognitionMode.LongDictation;
                }

                return SpeechRecognitionMode.ShortPhrase;
            }
        }

        private string subscriptionKey = "0f9fac1ed80444b7ade2cc3d4e923ce8";

        public string SubscriptionKey
        {
            get
            {
                return this.subscriptionKey;
            }

            set
            {
                this.subscriptionKey = value;
            }
        }
        private string DefaultLocale
        {
            get { return "en-US"; }
        }

        private string LuisEndpointUrl
        {
            get { return ConfigurationManager.AppSettings["LuisEndpointUrl"]; }
        }
        private string AuthenticationUri
        {
            get
            {
                return ConfigurationManager.AppSettings["AuthenticationUri"];
            }
        }
        private bool UseMicrophone
        {
            get
            {
                return this.IsMicrophoneClientWithIntent ||
                    this.IsMicrophoneClientDictation ||
                    this.IsMicrophoneClientShortPhrase;
            }
        }
        #endregion
        bool isListening = false;
        private void btnStart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!isListening)
                {
                    isListening = true;
                    this.btnStart.IsEnabled = false;

                    this.CreateMicrophoneRecoClient();
                    this.micClient.StartMicAndRecognition();
                }
                else
                {

                    isListening = false;
                    this.btnStart.IsEnabled = true;

                    this.micClient.EndMicAndRecognition();
                }
            }catch(Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private void OnMicrophoneStatus(object sender, MicrophoneEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                WriteLine("--- Microphone status change received by OnMicrophoneStatus() ---");
                WriteLine("********* Microphone status: {0} *********", e.Recording);
                if (e.Recording)
                {
                    WriteLine("Please start speaking.");
                }

                WriteLine();
            });
        }

        private void WriteLine()
        {
            this.WriteLine(string.Empty);
        }

        private void WriteLine(string format, params object[] args)
        {
            var formattedStr = string.Format(format, args);
            Trace.WriteLine(formattedStr);
            Dispatcher.Invoke(() =>
            {
                txtLog.Text += (formattedStr + "\n");
                txtLog.ScrollToEnd();
            });
        }


        private void CreateMicrophoneRecoClient()
        {
            this.micClient = SpeechRecognitionServiceFactory.CreateMicrophoneClient(
                this.Mode,
                this.DefaultLocale,
                this.SubscriptionKey);
            this.micClient.AuthenticationUri = this.AuthenticationUri;

            // Event handlers for speech recognition results
            this.micClient.OnMicrophoneStatus += this.OnMicrophoneStatus;
            this.micClient.OnPartialResponseReceived += this.OnPartialResponseReceivedHandler;
            if (this.Mode == SpeechRecognitionMode.ShortPhrase)
            {
                this.micClient.OnResponseReceived += this.OnMicShortPhraseResponseReceivedHandler;
            }
            else if (this.Mode == SpeechRecognitionMode.LongDictation)
            {
                this.micClient.OnResponseReceived += this.OnMicDictationResponseReceivedHandler;
            }

            this.micClient.OnConversationError += this.OnConversationErrorHandler;
        }

        private void OnPartialResponseReceivedHandler(object sender, PartialSpeechResponseEventArgs e)
        {
            this.WriteLine("--- Partial result received by OnPartialResponseReceivedHandler() ---");
            this.WriteLine("{0}", e.PartialResult);
            this.WriteLine();
        }

        private void OnMicShortPhraseResponseReceivedHandler(object sender, SpeechResponseEventArgs e)
        {
            Dispatcher.Invoke((Action)(() =>
            {
                this.WriteLine("--- OnMicShortPhraseResponseReceivedHandler ---");

                // we got the final result, so it we can end the mic reco.  No need to do this
                // for dataReco, since we already called endAudio() on it as soon as we were done
                // sending all the data.
                this.micClient.EndMicAndRecognition();

                this.WriteResponseResult(e);
                if (isListening)
                {
                    this.micClient.StartMicAndRecognition();
                }

                //  btnStart.IsEnabled = true;
            }));
        }
        private void OnMicDictationResponseReceivedHandler(object sender, SpeechResponseEventArgs e)
        {
            this.WriteLine("--- OnMicDictationResponseReceivedHandler ---");
            if (e.PhraseResponse.RecognitionStatus == RecognitionStatus.EndOfDictation ||
                e.PhraseResponse.RecognitionStatus == RecognitionStatus.DictationEndSilenceTimeout)
            {
                Dispatcher.Invoke(
                    (Action)(() =>
                    {
                        // we got the final result, so it we can end the mic reco.  No need to do this
                        // for dataReco, since we already called endAudio() on it as soon as we were done
                        // sending all the data.
                        this.micClient.EndMicAndRecognition();

                        // this.btnStart.IsEnabled = true;
                    }));
            }

            this.WriteResponseResult(e);
            this.micClient.StartMicAndRecognition();
        }

        private void OnConversationErrorHandler(object sender, SpeechErrorEventArgs e)
        {
            Dispatcher.Invoke(() =>
            {
                btnStart.IsEnabled = true;
            });

            this.WriteLine("--- Error received by OnConversationErrorHandler() ---");
            this.WriteLine("Error code: {0}", e.SpeechErrorCode.ToString());
            this.WriteLine("Error text: {0}", e.SpeechErrorText);
            this.WriteLine();
        }

        private async void WriteResponseResult(SpeechResponseEventArgs e)
        {
            RecognizedPhrase _final = null;
            if (e.PhraseResponse.Results.Length == 0)
            {
                this.WriteLine("No phrase response is available.");
            }
            else
            {
                this.WriteLine("********* Final n-BEST Results *********");
                for (int i = 0; i < e.PhraseResponse.Results.Length; i++)
                {
                    this.WriteLine(
                        "[{0}] Confidence={1}, Text=\"{2}\"",
                        i,
                        e.PhraseResponse.Results[i].Confidence,
                        e.PhraseResponse.Results[i].DisplayText);
                    if (_final == null)
                    {
                        _final = e.PhraseResponse.Results[i];
                    }
                    else
                    {
                        if (_final.Confidence < e.PhraseResponse.Results[i].Confidence)
                        {
                            _final = e.PhraseResponse.Results[i];
                        }
                    }
                }

                if (_final != null)
                {
                    CheckLuisForAction(_final.DisplayText);
                }
                this.WriteLine();
            }
        }

        private async void CheckLuisForAction(string phrase)
        {
            using (var client = new HttpClient())
            {
                //client.BaseAddress = new Uri("http://localhost:81/");
                string uri = "https://westus.api.cognitive.microsoft.com/luis/v2.0/apps/"+appId+"?subscription-key="+luisSubscriptionKey+"&verbose=true&timezoneOffset=0&q=";
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                HttpResponseMessage response;


                response = await client.GetAsync(uri + phrase);

                if (response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();


                    this.WriteLine("--- Understanding Phrases by LUIS ---");
                    LuisResponse luisResponse = JsonConvert.DeserializeObject<LuisResponse>(content);

                    if (luisResponse.topScoringIntent != null)
                    {
                        this.WriteLine("Top Intent : " + luisResponse.topScoringIntent.intent.ToUpper());
                        if (luisResponse.topScoringIntent.intent.ToUpper() == "LIGHT ON" ||
                            luisResponse.topScoringIntent.intent.ToUpper() == "LIGHT OFF" ||
                            luisResponse.topScoringIntent.intent.ToUpper() == "LIGHT BLINK")
                        {
                            foreach (var entity in luisResponse.entities)
                            {
                                this.WriteLine("Detected Entity : " + entity.entity + "(" + entity.type.ToUpper()+")");
                                if (entity.type.ToUpper() == "MOOD")
                                {
                                    var map = lightMaps.Select(x => x).Where(x => x.keywords.Contains(entity.entity)).FirstOrDefault();
                                    if(map != null)
                                    {
                                        this.WriteLine("Color : " + map.name);
                                        if (luisResponse.topScoringIntent.intent.ToUpper() == "LIGHT ON")
                                            _serialPort.Write("ON:"+map.io);
                                        else if (luisResponse.topScoringIntent.intent.ToUpper() == "LIGHT OFF")
                                            _serialPort.Write("OFF:" + map.io);
                                        else if (luisResponse.topScoringIntent.intent.ToUpper() == "LIGHT BLINK")
                                        {
                                            _serialPort.Write("ON:" + map.io);
                                            System.Threading.Thread.Sleep(speed);
                                            _serialPort.Write("OFF:" + map.io);
                                        }
                                    }
                                }
                                else if (entity.type.ToUpper() == "COLOR")
                                {
                                    var map = lightMaps.Select(x => x).Where(x => x.name.ToLower().Contains(entity.entity)).FirstOrDefault();
                                    if (map != null)
                                    {
                                        this.WriteLine("Color : " + map.name);
                                        if (luisResponse.topScoringIntent.intent.ToUpper() == "LIGHT ON")
                                            _serialPort.Write("ON:" + map.io);
                                        else if (luisResponse.topScoringIntent.intent.ToUpper() == "LIGHT OFF")
                                            _serialPort.Write("OFF:" + map.io);
                                        else if (luisResponse.topScoringIntent.intent.ToUpper() == "LIGHT BLINK")
                                        {
                                            _serialPort.Write("ON:" + map.io);
                                            System.Threading.Thread.Sleep(speed);
                                            _serialPort.Write("OFF:" + map.io);
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
            }
        }
    }
}
