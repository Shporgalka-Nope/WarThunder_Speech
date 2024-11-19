using System.Globalization;
using System.Speech.Recognition;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using WindowsInput;
using System.Speech.Synthesis;

namespace WarThunder_Speech
{
    internal class Program
    {
        public static int currentThrottle = 0;
        public static InputSimulator isim = new InputSimulator();

        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr hWnd);

        static void Main()
        {
            var cultureInf = new System.Globalization.CultureInfo("en-US");
            using (SpeechRecognitionEngine recognizerEngine = new SpeechRecognitionEngine())
            {
                recognizerEngine.SetInputToDefaultAudioDevice();
                recognizerEngine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(speechMethod);
                recognizerEngine.SpeechDetected += new EventHandler<SpeechDetectedEventArgs>(speechDetected);
                //Set of german commands
                recognizerEngine.LoadGrammar(CreateKeyValuesGrammar(cultureInf, "", new string[] {"los los", "foward", "fowards" }));
                recognizerEngine.LoadGrammar(CreateKeyValuesGrammar(cultureInf, "", new string[] { "hinterward", "hinterwards", "ruckwards", "langsamer" }));
                recognizerEngine.LoadGrammar(CreateKeyValuesGrammar(cultureInf, "", new string[] { "links", "nach links", "rechts", "nach rechts" }));
                recognizerEngine.LoadGrammar(CreateKeyValuesGrammar(cultureInf, "", new string[] { "halten", "anhalten", "abstellen" }));
                //English commands (test use only)
                //recognizerEngine.LoadGrammar(CreateKeyValuesGrammar(cultureInf, "", new string[] { "slower", "foward", "faster", "stop", "left", "right", "stop turn" }));
                recognizerEngine.RecognizeAsync(RecognizeMode.Multiple);

                Process[] ps = Process.GetProcessesByName("aces");
                Process WT_Process = ps.FirstOrDefault();
                if (WT_Process != null)
                {
                    IntPtr h = WT_Process.MainWindowHandle;
                    SetForegroundWindow(h);
                }

                Console.WriteLine("Is now listening.");
                while (true)
                {
                    Console.ReadLine();
                }
            }
        }

        static Grammar CreateKeyValuesGrammar(CultureInfo cultureInf, string key, string[] values)
        {
            var grBldr = string.IsNullOrWhiteSpace(key) ? new GrammarBuilder() { Culture = cultureInf } : new GrammarBuilder(key) { Culture = cultureInf };
            grBldr.Append(new Choices(values));

            return new Grammar(grBldr);
        }

        private static void speechDetected(object sender, SpeechDetectedEventArgs e)
        {
            Console.WriteLine("[Speech Detected]");
        }
        private static void speechMethod(object sender, SpeechRecognizedEventArgs e)
        {
            Console.WriteLine($"Heard: {e.Result.Text}");
            switch ((e.Result.Text).ToLower())
            {
                case "eins laden":
                case "eins":
                    isim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_1);
                    Thread.Sleep(500);
                    isim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_1);
                    break;
                case "zwei laden":
                case "zwei":
                    isim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_2);
                    Thread.Sleep(500);
                    isim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_2);
                    break;
                case "foward":
                case "fowards":
                case "los los":
                    isim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_E);
                    Thread.Sleep(10);
                    isim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_E);
                    //isim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_S);
                    //isim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_W);
                    break;
                case "hinterward":
                case "hinterwards":
                case "ruckwards":
                case "langsamer":
                    isim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_Q);
                    Thread.Sleep(10);
                    isim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_Q);
                    break;
                case "links":
                case "nach links":
                    isim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_D);
                    isim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_A);
                    break;
                case "rechts":
                case "nach rechts":
                    isim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_A);
                    isim.Keyboard.KeyDown(WindowsInput.Native.VirtualKeyCode.VK_D);
                    break;
                case "halten":
                case "anhalten":
                case "abstellen":
                    isim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_A);
                    isim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.VK_D);
                    break;
            }
        }
    }
}
