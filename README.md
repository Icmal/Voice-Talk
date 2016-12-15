# Voice-Talk
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using artyom.Properties;
using System.Speech.Synthesis;
using System.Speech.Recognition;
//using System.Windows.Input; used for keyboards press button simulating, some errors occured after adding the different libraries
using System.Runtime.InteropServices;
using System.Media;
using System.Text.RegularExpressions;
using System.IO.Ports;


namespace artyom
{
    
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine _recognizer = new SpeechRecognitionEngine(new CultureInfo("en-US")); // WSR 8.0 engine is added
        SpeechSynthesizer JARVIS = new SpeechSynthesizer(); // Speech synthesizer, Ivona Voices are also added here
        Grammar shellcommandgrammar; // when the user wants to add commands to program on the Adding Commands Windows, this grammer is updated
        //------------------ these variables are used for shell commands----------------------------------------------
        String[] ArrayShellCommands;
        String[] ArrayShellResponse;
        String[] ArrayShellLocation;
        String userName = Environment.UserName;
        string scpath;
        string srpath;
        string slpath;
        int i;
        //------------------------------------------------------------------------------------------------------------
        Boolean flg = true;                      // opens and closes listening flag
        Boolean flg_shell = true;                // opens and closes shell commands flag
        Boolean flg_default = true;              // default commands are always open
        DateTime now = DateTime.Now;             // needed to read the date and time
        int mouse = 10; //mouse rate: range for mouse movements
        //--------------------------Mouse control-----------------------------------------------
        [ DllImport ("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)] // imports dynamic-link library
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo); //right-left clicks

        private const int LEFTDOWN = 0x02; // hex codes for clicking operations
        private const int LEFTUP = 0x04; // hex codes for clicking operations
        private const int RIGHTDOWN = 0x08; // hex codes for clicking operations
        private const int RIGHTUP = 0x10; // hex codes for clicking operations
        
        //-------------------------------------------------------------------------------------------------------------
        //-----------------------Eject CD-DVD------------------------------------------------------
        [DllImport("winmm.dll", EntryPoint = "mciSendStringA", CharSet = CharSet.Ansi)]
        protected static extern int mciSendString(string lpstrCommand,
        StringBuilder lpstrReturnString,
        int uReturnLength,
        IntPtr hwndCallback);
        //-------------------------------------------------------------------------------------------------------------
        int count = 1;      //this counter is needed by ALT TAB operation
        //----------------------------------Volume operations-----------------------------------------------------
        private const int VOLUME_ONOFF= 0x80000;
        private const int VOLUME_UP = 0xA0000;
        private const int VOLUME_DOWN = 0x90000;
        private const int WM_APPCOMMAND = 0x319;

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessageW(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);
        //------------------------------------------------------------------------------------------------------------
        //--------------------CAPSLOCK operation---------------------------------------------------------
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        //------------------------------------------------------------------------------------------------------------
        Form2 shell = new Form2();      // adding commands form, form2
        Form3 mail = new Form3();        // send mail form, form3
        Form4 search = new Form4();      //google search form, form4
        Form5 temperature = new Form5();  // temperature CONTROL form, form5
        Form6 reading = new Form6();      // reading form, form6
        Boolean search_flg = false;     //complete search ü search google a baglamak için kullanılan flag
        Boolean mail_flg = false;       //mail formundaki attach ve send i send mail komutuna baglamak için kullanılan flag
        Boolean key_flg = false;        //keyboardı açıp kapayan flag
        public Boolean temperature_flg = false; // Isı kontrolü için kullanılan flag
        Boolean serial_flg = false; // serial port
        Random rndm = new Random();     //yanıtlardaki çeşitlilik için kullanılıyor.
        
        public Form1()
        {
            InitializeComponent();
            this.FormClosing += Form1_FormClosing;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            Directory.CreateDirectory(@"C:\Users\" + userName + "\\Documents\\Jarvis Custom Commands"); // create directory 
            Settings.Default.ShellC = @"C:\Users\" + userName + "\\Documents\\Jarvis Custom Commands\\Shell Commands.txt"; // create text file
            Settings.Default.ShellR = @"C:\Users\" + userName + "\\Documents\\Jarvis Custom Commands\\Shell Response.txt"; // create text file
            Settings.Default.ShellL = @"C:\Users\" + userName + "\\Documents\\Jarvis Custom Commands\\Shell Location.txt"; // create text file
            Settings.Default.Save(); // saving the changes in the directory
            scpath = Settings.Default.ShellC; // gets the path for Shell Commands text file
            srpath = Settings.Default.ShellR; // gets the path for Shell Response text file
            slpath = Settings.Default.ShellL; // gets the path for Shell Location text file
            if (!File.Exists(scpath)) // if there is no shell commands text file in the path defined above
            {
                using (StreamWriter sw = File.CreateText(scpath)) { sw.Write("My Documents"); sw.Close(); } 
                // Add an example to first row of the shell commands text file by using StreamWriter
            }
            if (!File.Exists(srpath)) // if there is no shell response text file in the path defined above
            {
                using (StreamWriter sw = File.CreateText(srpath)) { sw.Write("Right away"); sw.Close(); }
                // Add an example to first row of the shell response text file by using StreamWriter
            }
            if (!File.Exists(slpath)) // if there is no shell location text file in the path defined above
            {
                using (StreamWriter sw = File.CreateText(slpath)) { sw.Write(@"C:\Users\"+ userName +"\\Documents"); sw.Close(); }
                // Add an example to first row of the shell location text file by using StreamWriter
            }
            ArrayShellCommands = File.ReadAllLines(scpath); // Read all lines in the shell commands text file
            ArrayShellResponse = File.ReadAllLines(srpath); // Read all lines in the shell response text file
            ArrayShellLocation = File.ReadAllLines(slpath); // Read all lines in the shell location text file
            try
            {
                shellcommandgrammar = new Grammar(new GrammarBuilder(new Choices(ArrayShellCommands))); // Gets the grammer form Shell inputs
                _recognizer.LoadGrammar(shellcommandgrammar); // Loads this Shell inputs to the program
            }
            catch
            {
                JARVIS.SpeakAsync("I've detected an in valid entry in your shell commands, possibly a blank line.");
            }
            _recognizer.SetInputToDefaultAudioDevice(); // Configurates this object to receive input from default audio device
            _recognizer.LoadGrammarAsync(new Grammar(new GrammarBuilder(new Choices(File.ReadAllLines(@"Default Commands.txt")))));
            // Grammar which is used in application is defined here
            if (flg) // flg is assigned for listening, if listening is on, flg == 1, otherwise flg == 0
            {
                if (flg_default) // flg_default is assigned for the default commands, this flag is always on
                    _recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Default_SpeechRecognized);
                if (flg_shell) // flg_shell is assigned for the shell commands, this flag can change by the inputs of the user
                    _recognizer.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Shell_SpeechRecognized);
            }
            _recognizer.RecognizeAsync(RecognizeMode.Multiple); // the recognizer continues performing asynchronous recognition operations
            JARVIS.SpeakAsync("Hello, Welcome to jarvis.");
            serialPort1.PortName = "COM3";
            serialPort1.BaudRate = 9600;
            
            if (!serialPort1.IsOpen)
            {
                try
                {
                    serialPort1.Open();
                    serialPort1.WriteLine("w");
                    serialPort1.Close();
                }
                catch (Exception)
                {

                }
            }
        }


       public void Form1_FormClosing(object sender, FormClosingEventArgs e)
{
    if (!serialPort1.IsOpen)
    {
        try
        {
            serialPort1.Open();
            serialPort1.WriteLine("r");
            serialPort1.Close();
        }
        catch (Exception)
        {

        }
    }
}

        private void Shell_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            string speech = e.Result.Text; // Audio input is converted to text and assigned this text to string 'speech'
            if (flg && flg_shell) // listening ON, Shell Commands ON
            {
                i = 0; // first row in the text files, i=1 second row...
                try
                {
                    foreach (string line in ArrayShellCommands)  // this is a for loop for each rows in the text files
                    {
                        
                        if (line == speech) // if first shell command equals to speech 
                        {
                            if (string.IsNullOrWhiteSpace(ArrayShellLocation[i])) // if Location text row  is empty
                            {
                                JARVIS.SpeakAsync(ArrayShellResponse[i]); // give a answer with voice only
                            }

                            else // if there is something to process in the Location row
                            {
                                Process.Start(ArrayShellLocation[i]);      // Process command opens everything in the given path
                                JARVIS.SpeakAsync(ArrayShellResponse[i]); // also informs the user by speech 
                            }

                        }
                        i += 1; // next row after process is done by current row
                    }
                }
                catch // if error occurs, as an example, space character in the line
                {
                    i += 1; // error occured at current row, go to next row
                    JARVIS.SpeakAsync("I'm sorry it appers the shell command " + speech + " on line " + (i+1) + " is accomplied by either ");
                    // Let the user know about the error, tell recognized speech and line number
                }
            }
        }

        private void Default_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {

            string speech = e.Result.Text;
            if (speech == "start listenning")
            {
                flg = true;
                
                JARVIS.Speak("listenning is activated sir.");
                if (!serialPort1.IsOpen)
                {
                    try
                    {
                        serialPort1.Open();
                        serialPort1.WriteLine("w");
                        serialPort1.Close();
                    }
                    catch (Exception)
                    {

                    }
                }

            }
            if (speech == "stop listenning")
            {
                flg = false;
                
                JARVIS.Speak("listenning is deactivated sir.");
                if (!serialPort1.IsOpen)
                {
                    try
                    {
                        serialPort1.Open();
                        serialPort1.WriteLine("q");
                        serialPort1.Close();
                    }
                    catch (Exception)
                    {

                    }
                }
                
            }
            #region flg
            if (flg) 
            {
                switch (speech)
                {
                    #region speech

                    case "Hello":           // speech command
                    case "Hello Jarvis":    // speech command
                    case "Hi Jarvis":       // speech command
                        
                        if (now.Hour >= 6 && now.Hour < 12) // if the current time is between 6 and 12 in the morning
                        {
                            if (rndm.Next() % 3 == 1)     // generated random number
                                JARVIS.SpeakAsync("Hello sir,such a lovely morning ha?"); // speech output
                            else if (rndm.Next() % 3 == 2) // generated random number
                                JARVIS.SpeakAsync("Good morning sir"); // speech output
                            else // generated random number
                                JARVIS.SpeakAsync("Hello sir.");  // speech output
                        }
                        if (now.Hour >= 12 && now.Hour <= 18) // if the current time is between 12 and 18
                        {
                            if (rndm.Next() % 3 == 1)  // generated random number
                                JARVIS.SpeakAsync("Hello sir,what can i do for you?"); // speech output
                            else if (rndm.Next() % 3 == 2)  // generated random number
                                JARVIS.SpeakAsync("Good Afternoon sir"); // speech output
                            else  // generated random number
                                JARVIS.SpeakAsync("Hello sir."); // speech output
                        }
                        if (now.Hour > 18 && now.Hour <= 22) // if the current time is between 18 and 22
                        {
                            if (rndm.Next() % 3 == 1) // generated random number
                                JARVIS.SpeakAsync("Hello sir,i hope your day was good"); // speech output
                            else if (rndm.Next() % 3 == 2) // generated random number
                                JARVIS.SpeakAsync("Good evening sir"); // speech output
                            else // generated random number
                                JARVIS.SpeakAsync("Hello sir."); // speech output
                        }
                        if (now.Hour > 22 && now.Hour < 6) // if the current time is between 22 and 6
                        {
                            if (rndm.Next() % 3 == 1) // generated random number
                                JARVIS.SpeakAsync("Hello sir,it's getting late"); // speech output
                            else if (rndm.Next() % 3 == 2) // generated random number
                                JARVIS.SpeakAsync("Good night sir"); // speech output
                            else // generated random number
                                JARVIS.SpeakAsync("Hello world."); // speech output
                        }
                        break;
                     
                    case "Are you there":
                        if (rndm.Next() % 3 == 1)
                            JARVIS.SpeakAsync("i am here,don't worry");
                        else if (rndm.Next() % 3 == 2)
                            JARVIS.SpeakAsync("yes ofcourse");
                        else
                            JARVIS.SpeakAsync("i am always with you");
                        break;
                    case "Goodbye":
                    case "Goodbye Jarvis":
                    case "Close":
                    case "Close Jarvis":
                        //SpeakAsync kullanmama nedenin asyc thread olusturuyo ve sıraya giriyo yazının okunu olarak asyc iyi hızlı okuyo ama threah oldugu için close işlemi önce devreye geliyo ve konusma duyulmadan program kapanıyor.
                         if (now.Hour >= 6 && now.Hour < 12)
                        {
                            if (rndm.Next() % 3 == 1)
                                JARVIS.Speak("Good bye sir,hope to enjoy this lovely morning");
                            else if (rndm.Next() % 3 == 2)
                                JARVIS.Speak("Good bye sir");
                            else
                                JARVIS.Speak("Farewell");
                        }
                        if (now.Hour >= 12 && now.Hour <= 18)
                        {
                            if (rndm.Next() % 3 == 1)
                                JARVIS.Speak("As you wish");
                            else if (rndm.Next() % 3 == 2)
                                JARVIS.Speak("Have a nice day sir");
                            else
                                JARVIS.Speak("Farewell");
                        }
                        if (now.Hour > 18 && now.Hour <= 22)
                        {
                            if (rndm.Next() % 3 == 1)
                                JARVIS.Speak("See you soon");
                            else if (rndm.Next() % 3 == 2)
                                JARVIS.Speak("take care till next time");
                            else
                                JARVIS.Speak("Farewell.");
                        }
                        if ((now.Hour > 22 && now.Hour <= 24)||(now.Hour<6))
                        {
                            if (rndm.Next() % 3 == 1)
                                JARVIS.Speak("it's getting late,you should sleep too");
                            else if (rndm.Next() % 3 == 2)
                                JARVIS.Speak("Good night sir");
                            else
                                JARVIS.Speak("Farewell");
                        }
                       
                        if (!serialPort1.IsOpen)
                        {
                            try
                            {
                                serialPort1.Open();
                                serialPort1.WriteLine("r");
                                serialPort1.Close();
                            }
                            catch (Exception)
                            {

                            }
                        }

                        Close();
                        break;

                    case "change your voice":
                        try
                        {
                            JARVIS.SelectVoice("IVONA 2 Salli");
                            JARVIS.Speak("My voice has changed");
                        }
                        catch (Exception) 
                        {
                            MessageBox.Show("Other voices is not available in this computer");
                        }
                         break;

                    case "orginal voice":
                         JARVIS.SelectVoice("IVONA 2 Brian");
                         JARVIS.Speak("I am back sir!");
                         break;

                    case "Terminate yourself":
                        JARVIS.Speak("Jarvis is terminating");
                        if (!serialPort1.IsOpen)
                        {
                            try
                            {
                                serialPort1.Open();
                                serialPort1.WriteLine("r");
                                serialPort1.Close();
                            }
                            catch (Exception)
                            {

                            }
                        }
                        Close();
                        break;
                    case "Who are you":
                        JARVIS.SpeakAsync("Hi my name is jarvis. i am a new generation personal assistant");
                        break;
                    case "What can i say":
                    case "What are the commands":
                        JARVIS.SpeakAsync("You can read here sir");

                        Process.Start(@"C:\Users\ACER\Desktop\Final\artyom\Default Commands.txt");
                        break;
                    #endregion

                    case "Update commands": // speech command 
                        UpdateCommands(); // calls a method called UpdateCommand
                        if (!serialPort1.IsOpen) // if serial port is not opened
                        {
                            try
                            {
                                serialPort1.Open(); // open serial port
                                serialPort1.WriteLine("y"); // send a character to Arduino
                                serialPort1.Close(); // close serial port
                            }
                            catch (Exception)
                            {
                                // if serial port is not connected, dont give error, continue to other tasks
                            }
                        }
                        break;
                    case "show commands": // speech command
                        JARVIS.SpeakAsync("just a sec"); // program informs the user
                        shell.Show(); // Send Mail Form is opened
                        if (!serialPort1.IsOpen) // if serial port is closed
                        {
                            try
                            {
                                serialPort1.Open();   // open the serial port
                                serialPort1.WriteLine("c"); // send a character to the arduino
                                serialPort1.Close(); // close the serial port
                            }
                            catch (Exception)
                            {
                                
                            }
                        }
                        break;

                    case "add commands": // speech command
                        JARVIS.SpeakAsync("adding these above commands"); // application informs the user
                        shell.button1_Click(sender, e); // performs click on the 'Add' button on the second form
                        break;

                    case "close commands":
                        if (!serialPort1.IsOpen)
                        {
                            try
                            {
                                serialPort1.Open();
                                serialPort1.WriteLine("o");
                                serialPort1.Close();
                            }
                            catch (Exception)
                            {

                            }
                        }
                        shell.Hide();
                        break;
                    case "ignore shell":
                        flg_shell = false;
                        JARVIS.Speak("ignoring shell commands");
                        break;
                    case "activate shell":
                        flg_shell = true;
                        JARVIS.Speak("activating shell commands");
                        break;
                    case "What time is it":
                        string time = now.GetDateTimeFormats('t')[0];
                        JARVIS.Speak(time);

                        try
                        {
                            serialPort1.Open();
                            serialPort1.WriteLine("Time: " + time);
                            serialPort1.Close();
                        }
                        catch (Exception)
                        {

                        }

                        break;
                    case "What day is it":
                        string day="";
                        
                        if (DateTime.Today.ToString("dddd") == "Pazartesi") {
                            JARVIS.Speak("Monday");
                            day = "Monday   "; }
                        else if (DateTime.Today.ToString("dddd") == "Salı")
                        {
                            JARVIS.Speak("Tuesday");
                            day = "Tuesday  ";
                        }
                        else if (DateTime.Today.ToString("dddd") == "Çarşamba")
                        {
                            JARVIS.Speak("Wednesday");
                            day = "Wednesday";
                        }
                        else if (DateTime.Today.ToString("dddd") == "Perşembe")
                        {
                            JARVIS.Speak("Thursday");
                            day = "Thursday ";
                        }
                        else if (DateTime.Today.ToString("dddd") == "Cuma")
                        {
                            JARVIS.Speak("Friday");
                            day = "Friday   ";
                        }
                        else if (DateTime.Today.ToString("dddd") == "Cumartesi")
                        {
                            JARVIS.Speak("Saturday");
                            day = "Saturday ";
                        }
                        else if (DateTime.Today.ToString("dddd") == "Pazar")
                        {
                            JARVIS.Speak("Sunday");
                            day = "Sunday   ";
                        }
                        else
                            JARVIS.Speak(DateTime.Today.ToString("dddd"));
                        try
                        {
                            serialPort1.Open();                           
                            serialPort1.WriteLine("Day: "+ day);
                            serialPort1.Close();
                        }
                        catch (Exception) { }

                        break;
                    case "Whats the date":
                    case "Whats todays date":
                        JARVIS.SpeakAsync(DateTime.Now.ToShortDateString());
                        try
                        {
                            serialPort1.Open();
                            serialPort1.WriteLine("Date: "+DateTime.Now.ToShortDateString());
                            serialPort1.Close();
                        }
                        catch (Exception) { }
                        break;


                    #region mouse commands

                    case "mouse up": // speech command
                        Cursor.Position = new Point(Cursor.Position.X, Cursor.Position.Y - mouse); // set cursor up
                        break;
                    case "mouse down": // speech command
                        Cursor.Position = new Point(Cursor.Position.X, Cursor.Position.Y + mouse); // set cursor down
                        break;
                    case "mouse right":// speech command
                        Cursor.Position = new Point(Cursor.Position.X+mouse, Cursor.Position.Y ); // set cursor right
                        break;
                    case "mouse left": // speech command
                        Cursor.Position = new Point(Cursor.Position.X-mouse, Cursor.Position.Y); // set cursor left
                        break;
                    case "mouse left up":
                       Cursor.Position = new Point(0,0);
                        break;
                    case "mouse left down":
                        Cursor.Position = new Point(0, Screen.PrimaryScreen.Bounds.Width);
                        JARVIS.Speak("mouse left down");
                        break;
                    case "mouse right up":
                        Cursor.Position = new Point(Screen.PrimaryScreen.Bounds.Right, 0);
                        break;
                    case "mouse right down":
                        Cursor.Position = new Point(Screen.PrimaryScreen.Bounds.Right, Screen.PrimaryScreen.Bounds.Bottom);
                        break;
                    case "mouse center":
                        Cursor.Position = new Point(Screen.PrimaryScreen.Bounds.Right / 2, Screen.PrimaryScreen.Bounds.Bottom / 2);
                        break;    
                    case "mouse go close button":
                        Cursor.Position = new Point(Screen.PrimaryScreen.Bounds.Right-10,10);
                        break;
                    case "mouse go minimize button":
                        Cursor.Position = new Point(Screen.PrimaryScreen.Bounds.Right - 90, 10);
                        break;
                    case "mouse go restore button": // speech command 
                        Cursor.Position = new Point(Screen.PrimaryScreen.Bounds.Right - 45, 10); // cursor goes to restore button
                        break;

                    case "change mouse rate with ten": // speech command
                        mouse = 10; // mouse rate can be changed by using this kind of calculations
                        JARVIS.SpeakAsync("mouse rate is changed"); // speech output, also the value of the mouse can be given here
                        break;
                    case "change mouse rate with five":
                        mouse = 5;
                        JARVIS.SpeakAsync("mouse rate is changed");
                        break;
                    case "increase mouse rate":
                        mouse += 5;
                        JARVIS.Speak("new mouse rate is" + mouse);
                        break;
                    case "decrease mouse rate":
                        mouse -= 5;
                        JARVIS.Speak("new mouse rate is" + mouse);
                        break;
                    case "increase mouse rate ten times":
                        mouse *= 10;
                        JARVIS.Speak("new mouse rate is" + mouse);
                        break;
                    case "decrease mouse rate ten times":
                        mouse /= 10;
                        JARVIS.Speak("new mouse rate is" + mouse);
                        break;
                    case "increase mouse rate five times":
                        mouse *= 5;
                        JARVIS.Speak("new mouse rate is" + mouse);
                        break;
                    case "decrease mouse rate five times":
                        mouse /= 5;
                        JARVIS.SpeakAsync("new mouse rate is" + mouse);
                        break;
                    case "mouse left click":
                    case "do left click":
                    case "press left click":
                        mouse_event(LEFTDOWN | LEFTUP, X, Y , 0, 0);
                        break;
                    case "mouse right click":
                    case "do right click":
                    case "press right click":
                        mouse_event(RIGHTDOWN | RIGHTUP, X, Y, 0, 0);
                        break;
                    #endregion 

                    case "eject cd":
                    case "eject dvd":
                        int ret = mciSendString("set cdaudio door open", null, 0, IntPtr.Zero);
                        break;

                    case "alt tab":
                    case "switch the window":
                        SendKeys.Send("%{TAB " + count + "}");
                        count += 1;
                        break;
                    case "terminate current program":
                    case "close the window":
                        SendKeys.Send("%{F4}");
                        break;
                    case "mute the volume":
                    case "turn off the volume":
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle, (IntPtr)VOLUME_ONOFF);
                        break;
                    case "turn up the volume":
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle, (IntPtr)VOLUME_UP);
                        break;
                    case "turn down the volume":
                        SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle, (IntPtr)VOLUME_DOWN);
                        break;
                    case "turn on the volume":
                         SendMessageW(this.Handle, WM_APPCOMMAND, this.Handle, (IntPtr)VOLUME_ONOFF);
                        break;
                    case "change your background":
                        OpenFileDialog op = new OpenFileDialog();
                        op.ShowDialog();
                        this.BackgroundImage = new Bitmap(op.FileName);
                        break;
                        

                    #region form3 mail sender

                    case "i wanna send mail":
                        JARVIS.Speak("opening mail form");
                        mail_flg = true;
                        mail.Show();
                        try
                        {
                            serialPort1.Open();
                            serialPort1.WriteLine("f");
                            serialPort1.Close();
                        }
                        catch (Exception) { }

                        break;
                        
                    case "send my mail":
                        if (mail_flg)
                        {
                            mail.b_send_Click(sender, e);
                            mail_flg = false;
                        }
                        else
                            JARVIS.SpeakAsync("invalid command");
                        break;

                    case "attach file to my mail":
                        if(mail_flg)
                            mail.b_attach_Click(sender, e);
                        else
                            JARVIS.SpeakAsync("invalid command");
                        break;
                        
                    case "close mail form":
                        mail.Hide();
                        try
                        {
                            serialPort1.Open();
                            serialPort1.WriteLine("g");
                            serialPort1.Close();
                        }
                        catch (Exception) { }
                        break;

                    #endregion 


                    #region form4 google search
                    case "search google":
                        JARVIS.SpeakAsync("what are you looking for?");
                        search_flg = true;
                        search.textBox1.Select();
                        search.Show();

                        try
                        {
                            serialPort1.Open();
                            serialPort1.WriteLine("j");
                            serialPort1.Close();
                        }
                        catch (Exception)
                        { }
                        
                        break;
                    case "complete search":
                        if (search_flg)
                            search.button1_Click(sender, e);
                        else
                            JARVIS.SpeakAsync("invalid command");
                        search_flg = false;
                        break;

                    case "thanks for search":
                        try
                        {
                            serialPort1.Open();
                            serialPort1.WriteLine("k");
                            serialPort1.Close();
                        }
                        catch (Exception)
                        { }
                        search.Hide();
                        break;
                    #endregion



                    #region form5 Temperature Control
                    case "Temperature":
                        JARVIS.SpeakAsync("Opening the control panel sir");
                        temperature_flg = true;
                        temperature.Show();
           
            try
            {
                serialPort1.Open();
                serialPort1.WriteLine("b");
                serialPort1.Close();
            }
            catch (Exception) { }

                        break;

                                case "First sensor":

                            temperature.button1.PerformClick();
                            string d = Convert.ToString(temperature.label1.Text);
                            string[] dd = Regex.Split(d, "C");
                            JARVIS.SpeakAsync(dd[0] + "Degrees");


                     /*      if (d.Contains("."))  // To convert float value to decimal value this function should be used
                           // Example; 26.54C will be read as 27C, 23.49C will be read as 23C 
                           d = d.Replace(".", ","); // Replace the '.' with the ','
                           decimal decval = System.Convert.ToDecimal(d); // here the output decval is the output
                     */    
                                    break;

                                case "Second sensor": // speech command
                                    
                                        temperature.button2.PerformClick(); // control the button on the form5
                                        string f = Convert.ToString(temperature.label2.Text); // take a string from second label on the form5
                                        string[] ff = Regex.Split(f, "C"); // split string from a special character
                                        JARVIS.SpeakAsync(ff[0] + "Degrees"); // Use the first part of the string and give a speech output
                                                                        
                                    break;

                                case "Third sensor": // speech command

                                    temperature.button3.PerformClick(); // control the button on the form5
                                    string g = Convert.ToString(temperature.label3.Text); // take a string from third label on the form5
                                    string[] gg = Regex.Split(g, "C"); // split string from a special character
                                    JARVIS.SpeakAsync(gg[0] + "Degrees"); // Use the first part of the string and give a speech output

                                    break;
                                case "Fourth sensor":

                                    temperature.button4.PerformClick();
                                    string h = Convert.ToString(temperature.label4.Text);
                                    string[] hh = Regex.Split(h, "C");
                                    JARVIS.SpeakAsync(hh[0] + "Degrees");
                                   
                                    break;

                                case "Fifth sensor":

                                    temperature.button5.PerformClick();
                                    string j = Convert.ToString(temperature.label5.Text);
                                    string[] jj = Regex.Split(j, "C");
                                    JARVIS.SpeakAsync(jj[0] + "Degrees");

                                    break;


                                case "close the control panel":
                                    JARVIS.SpeakAsync("Control panel is closing");
                                    temperature.Hide();
                                    try
                                    {
                                        serialPort1.Open();
                                        serialPort1.WriteLine("n");
                                        serialPort1.Close();
                                    }
                                    catch (Exception) { }
                                    break;

                    #endregion


                    #region Read text
                    case "open reading form":
                    case "activate reading form":

                        JARVIS.SpeakAsync("Opening reading form");
                        reading.Show();

                        try
                        {
                            serialPort1.Open();
                            serialPort1.WriteLine("l");
                            serialPort1.Close();
                        }
                        catch (Exception)
                        { }

                        break;

                    case "read in english": // speech command
                                                
                        string a = Convert.ToString(reading.textBox1.Text); // convert the text inside the textbox on Reading Form
                        JARVIS.SpeakAsync(a); // give an output in english
                        break;

                    case "read in russian": // speech command
                        try
                        {
                            JARVIS.SelectVoice("IVONA 2 Tatyana"); // change the speech synthesizer for Russian Languge 
                            // this speech synthesizer can also convert kiril alphabet to speech
                            string b = Convert.ToString(reading.textBox1.Text); // convert the text inside the textbox on Reading Form
                            JARVIS.Speak(b); // here, it is not a async speak. It is used, because with the SpeakAsync Program would 
                            // go to next row without waiting the reading operation
                            JARVIS.SelectVoice("IVONA 2 Brian"); // change the speech synthesizer for orginal voice
                        }
                        catch (Exception) 
                        {
                            MessageBox.Show("Russian language package is not installed in this computer, install it to continue to reading operating");
                        }



                        break;

                    case "read in turkish": // speech command
                        try
                        {
                            JARVIS.SelectVoice("IVONA 2 Filiz"); // change the speech synthesizer for Turkish Languge 
                            string c = Convert.ToString(reading.textBox1.Text); // convert the text inside the textbox on Reading Form
                            JARVIS.Speak(c); // here, it is not a async speak. It is used, because with the SpeakAsync Program would 
                            // go to next row without waiting the reading operation
                            JARVIS.SelectVoice("IVONA 2 Brian"); // change the speech synthesizer for orginal voice
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Turkish language package is not installed in this computer, install it to continue to reading operating");
                        }
                        break;

                    case "thanks for reading":
                        reading.Hide();
                        try
                        {
                            serialPort1.Open();
                            serialPort1.WriteLine("i");
                            serialPort1.Close();
                        }
                        catch (Exception)
                        { }
                        break;

                    #endregion


                    case "unlock keyboard":
                        JARVIS.SpeakAsync("activating keyboard");
                        key_flg = true;
                        break;
                    case "lock keyboard":
                        JARVIS.SpeakAsync("deactivating keyboard");
                        key_flg = false;
                        break;
                    case "page down":
                        SendKeys.Send("{PGDN}");
                        break;
                    case "page up":
                        SendKeys.Send("{PGUP}");
                        break;
                    case "page top":
                        SendKeys.Send("{HOME}");
                        break;
                    case "page bottom":
                        SendKeys.Send("{END}");
                        break;
                    case "windows key":
                    case "start button":
                        //Simulate Key Press
                        keybd_event(0x5B, 0x45, 0x1, (UIntPtr)0);
                        //Simulate Key Release
                        keybd_event(0x5B, 0x45, 0x1 | 0x2, (UIntPtr)0);
                        break;
                    case "print screen":
                        
                        SendKeys.Send("{PRTSC}");
                        break;
                    case "open task manager":
                        SendKeys.Send("+^{ESC}");
                        break;
                    case "run command":
                     
                        keybd_event(0x5B, 0x45, 0, (UIntPtr)0); // key press 'windows'
                        keybd_event(0x52, 0x93, 0, (UIntPtr)0); // key press 'r'
                        keybd_event(0x52, 0x93, 0x0002, (UIntPtr)0); // release 'r'
                        keybd_event(0x5B, 0x45, 0x0002, (UIntPtr)0); // release 'windows'

                        break;
                    case "show desktop":
                        
                        keybd_event(0x5B, 0x45, 0, (UIntPtr)0); // key press 'windows'
                        keybd_event(0x44, 0xA0, 0, (UIntPtr)0); // key press 'd'
                        keybd_event(0x44, 0xA0, 0x0002, (UIntPtr)0); // release 'd'
                        keybd_event(0x5B, 0x45, 0x0002, (UIntPtr)0); // release 'windows'

                        break;
                    
                    case "show my system skills":
                        JARVIS.Speak("Here they are sir");
                       //Simulate windows Key Press
                        keybd_event(0x5B, 0x45, 0x1, (UIntPtr)0);
                        //simulate break key press
                        keybd_event(0x13, 0x45, 0x1, (UIntPtr)0);
                        
                        //Simulate break Key Release
                        keybd_event(0x13, 0x45, 0x1 | 0x2, (UIntPtr)0);
                        //Simulate windows Key Release
                        keybd_event(0x5B, 0x45, 0x1 | 0x2, (UIntPtr)0);
                        break;
                     
  
                }    //asıl switch case in sonu
                #region keyboard commands
                if (key_flg)
                {
                    switch (speech)
                    {
                        case "a": // speech command
                        case "1st": // speech command
                            SendKeys.Send("{a}"); // output, press the button 'a' on the keyboard
                            break;
                        case "2nd":
                        case "b":
                            SendKeys.Send("{b}");
                            break;
                        case "3rd":
                        case "c":
                            SendKeys.Send("{c}");
                            break;
                        case "4th":
                        case "d":
                            SendKeys.Send("{d}");
                            break;
                        case "5th":
                        case "e":
                            SendKeys.Send("{e}");
                            break;
                        case "6th":
                        case "f":
                            SendKeys.Send("{f}");
                            break;
                        case "7th":
                        case "g":
                            SendKeys.Send("{g}");
                            break;
                        case "8th":
                        case "h":
                            SendKeys.Send("{h}");
                            break;
                        case "9th":
                        case "i":
                            SendKeys.Send("{i}");
                            break;
                        case "10th":
                        case "j":
                            SendKeys.Send("{j}");
                            break;
                        case "11th":
                        case "k":
                            SendKeys.Send("{k}");
                            break;
                        case "12th":
                        case "l":
                            SendKeys.Send("{l}");
                            break;
                        case "13th":
                        case "m":
                            SendKeys.Send("{m}");
                            break;
                        case "14th":
                        case "n":
                            SendKeys.Send("{n}");
                            break;
                        case "15th":
                        case "o":
                            SendKeys.Send("{o}");
                            break;
                        case "16th":
                        case "p":
                            SendKeys.Send("{p}");
                            break;
                        case "17th":
                        case "q":
                            SendKeys.Send("{q}");
                            break;
                        case "18th":
                        case "r":
                            SendKeys.Send("{r}");
                            break;
                        case "19th":
                        case "s":
                            SendKeys.Send("{s}");
                            break;
                        case "20th":
                        case "t":
                            SendKeys.Send("{t}");
                            break;
                        case "21th":
                        case "u":
                            SendKeys.Send("{u}");
                            break;
                        case "22th":
                        case "v":
                            SendKeys.Send("{v}");
                            break;
                        case "23th":
                        case "w":
                            SendKeys.Send("{w}");
                            break;
                        case "24th":
                        case "x":
                            SendKeys.Send("{x}");
                            break;
                        case "25th":
                        case "y":
                            SendKeys.Send("{y}");
                            break;
                        case "26th":
                        case "z":
                            SendKeys.Send("{z}");
                            break;
                        case "space":
                            SendKeys.Send(" ");
                            break;
                        case "backspace":
                            SendKeys.Send("{BACKSPACE}");
                            break;
                        case "enter":
                            SendKeys.Send("{ENTER}");
                            break;
                        case "delete":
                            SendKeys.Send("{delete}");
                            break;
                        case "down arrow":
                            SendKeys.Send("{DOWN}");
                            break;
                        case "up arrow":
                            SendKeys.Send("{UP}");
                            break;
                        case "left arrow":
                            SendKeys.Send("{LEFT}");
                            break;
                        case "right arrow":
                            SendKeys.Send("{RIGHT}");
                            break;
                        case "1":
                        case "number one":
                            SendKeys.Send("{1}");
                            break;
                        case "2":
                        case "number two":
                            SendKeys.Send("{2}");
                            break;
                        case "3":
                        case "number three":
                            SendKeys.Send("{3}");
                            break;
                        case "4":
                        case "number four":
                            SendKeys.Send("{4}");
                            break;
                        case "5":
                        case "number five":
                            SendKeys.Send("{5}");
                            break;
                        case "6":
                        case "number six":
                            SendKeys.Send("{6}");
                            break;
                        case "7":
                        case "number seven":
                            SendKeys.Send("{7}");
                            break;
                        case "8":
                        case "number eight":
                            SendKeys.Send("{8}");
                            break;
                        case "9":
                        case "number nine":
                            SendKeys.Send("{9}");
                            break;
                        case "0":
                        case "number zero":
                            SendKeys.Send("{0}");
                            break;
                        case "+":
                        case "plus":
                            SendKeys.Send("{+}");
                            break;
                        case "-":
                        case "minus":
                            SendKeys.Send("{-}");
                            break;
                        case "*":
                        case "multiply":
                            SendKeys.Send("{*}");
                            break;
                        case "/":
                        case "divide":
                            SendKeys.Send("{/}");
                            break;
                        case ".":
                        case "dot":
                            SendKeys.Send("{.}");
                            break;
                        case ",":
                        case "comma":
                            SendKeys.Send("{,}");
                            break;
                        case ";":
                        case "semicolon":
                            SendKeys.Send("{;}");
                            break;
                        case ":":
                        case "colon":
                            SendKeys.Send("{:}");
                            break;
                        case "_":
                        case "underline":
                            SendKeys.Send("{_}");
                            break;
                        case "#":
                        case "hashtag":
                            SendKeys.Send("{#}");
                            break;
                        case "$":
                        case "dollar":
                            SendKeys.Send("{$}");
                            break;
                        case "&":
                        case "ampersand":
                            SendKeys.Send("{&}");
                            break;
                        case "backslash":
                            SendKeys.Send("{\\}");
                            break;
                        case "slash":
                            SendKeys.Send("{/}");
                            break;
                      
                        case "at sign":
                            SendKeys.Send("{@}");
                            break;
                        case "question sign":
                            SendKeys.Send("{?}");
                            break;
                        case "exclamation sign":
                            SendKeys.Send("{!}");
                            break;
                        case "percent sign":
                             SendKeys.Send("{%}");
                            break;
                        case "equals sign":
                            SendKeys.Send("{=}");
                            break;
                        case "caret sign":
                            SendKeys.Send("{^}");
                            break;
                        case "left parenthesis":
                            SendKeys.Send("{(}");
                            break;
                        case "right parenthesis":
                            SendKeys.Send("{)}");
                            break;
                        case "left square":
                            SendKeys.Send("{[}");
                            break;
                        case "right square":
                            SendKeys.Send("{]}");
                            break;
                        case "left flower":
                            SendKeys.Send("{{}");
                            break;
                        case "right flower":
                            SendKeys.Send("{}}");
                            break;
                        case "single quotation":
                            SendKeys.Send("{'}");
                            break;
                        case "double quotation":
                            SendKeys.Send("{\"}");
                            break;
                        case "vertical bar":
                            SendKeys.Send("{|}");
                            break;
                        case "left angle":
                            SendKeys.Send("{<}");
                            break;
                        case "right angle":
                            SendKeys.Send("{>}");
                            break;
                        case "tab sign":
                        case "big space":
                            SendKeys.Send("{TAB}");
                            break;
                        case "escape":
                            SendKeys.Send("{ESC}");
                            break;
                        case "capslock":
                            //Simulate Key Press
                            keybd_event(0x14, 0x45, 0x1, (UIntPtr)0);
                            //0x1 keydown 0x2 keyup
                            //Simulate Key Release
                            keybd_event(0x14, 0x45, 0x1 | 0x2, (UIntPtr)0);
                            break;
                        case "numlock":
                            
                            //Simulate Key Press
                             keybd_event(0x90, 0x45, 0x1, (UIntPtr)0);
                            //Simulate Key Release
                             keybd_event(0x90, 0x45, 0x1 | 0x2, (UIntPtr)0);
                            break;
                        
                        case "F1":
                            SendKeys.Send("{F1}");
                            break;
                        case "F2":
                            SendKeys.Send("{F2}");
                            break;
                        case "F3":
                            SendKeys.Send("{F3}");
                            break;
                        case "F4":
                            SendKeys.Send("{F4}");
                            break;
                        case "F5":
                            SendKeys.Send("{F5}");
                            break;
                        case "F6":
                            SendKeys.Send("{F6}");
                            break;
                        case "F7":
                            SendKeys.Send("{F7}");
                            break;
                        case "F8":
                            SendKeys.Send("{F8}");
                            break;
                        case "F9":
                            SendKeys.Send("{F9}");
                            break;
                        case "F10":
                            SendKeys.Send("{F10}");
                            break;
                        case "F11":
                            SendKeys.Send("{F11}");
                            break;
                        case "F12":
                            SendKeys.Send("{F12}");
                            break;

                        case "http":
                            SendKeys.Send("http://");
                            break;
                        case "https":
                            SendKeys.Send("https://");
                            break;
                        
                        case "www":
                            SendKeys.Send("www.");
                            break;
                        case ".com":
                            SendKeys.Send(".com");
                            break;
                        case ".com.tr":
                            SendKeys.Send(".com.tr");
                        
                            break;
                        case ".org":
                            SendKeys.Send(".org");
                            break;
                        case ".org.tr":
                            SendKeys.Send(".org.tr");
                            break;
                        case "copy":
                            SendKeys.SendWait("^C");
                            break;
                        case "paste":
                            SendKeys.SendWait("^V");
                            break;
                        case "corp":
                            SendKeys.SendWait("^X");
                            break;
                        case "undo":
                            SendKeys.SendWait("^Z");
                            break;
                        case "bold":
                            SendKeys.SendWait("^B");
                            break;
                        case "italic":
                            SendKeys.SendWait("^I");
                            break;
                        case "select all":
                            SendKeys.SendWait("^A");
                            break;
                        case "program files":
                            SendKeys.Send("C:\\Program Files");
                            break;
                        case "desktop":
                            SendKeys.Send("Desktop");
                            break;
                        case "documents":
                            SendKeys.Send("Libraries\\Documents");
                            break;
                        case "downloads":
                            SendKeys.Send("C:\\Users\\"+Environment.UserName.ToString() +"\\Downloads");
                            break;
                        //bu böyle devam eder.....
                            
                    }//keyborad switch in sonu

                    }  //keyboard if in sonu
                #endregion
            

                
            }   //if(flg) nin sonu
            #endregion
        }  //default_speechRecognized in sonu


        private void UpdateCommands() // called method 
        {
            JARVIS.SpeakAsync("this may take a few seconds"); // Async answer to inform the user
            _recognizer.UnloadGrammar(shellcommandgrammar); // unload the shellcommandgrammar
            ArrayShellCommands = File.ReadAllLines(scpath); // read all lines in scpath
            ArrayShellResponse = File.ReadAllLines(srpath); // read all lines in srpath
            ArrayShellLocation = File.ReadAllLines(slpath); // read all lines in slpath
            try
            {
                shellcommandgrammar = new Grammar(new GrammarBuilder(new Choices(ArrayShellCommands)));
                _recognizer.LoadGrammar(shellcommandgrammar);
                // define the shellcommandgrammar again with the last added lines
            }
            catch
            {
                JARVIS.SpeakAsync("I've detected an in valid entry in your shell commands, possibly a blank line. Shell commands will stop work until it is fixed");
                // If problem occurs, let the user know about situation
                if (!serialPort1.IsOpen) // if serial port is not open
                {
                    try
                    {
                        serialPort1.Open(); // opens serial port
                        serialPort1.WriteLine("x"); // sends a character to Arduino
                        serialPort1.Close(); // closes serial port
                    }
                    catch (Exception)
                    {
                        // dont do anything if there are problems with serial communication
                    }
                }
            }
            JARVIS.SpeakAsync("all commands updated"); // let the user know, new commands are usable now
        }

        //mouse x ve y------------------
        public uint X { get; set; }

        public uint Y { get; set; }
        //------------------------------------------
    }
}
