using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using Ivi.Visa.Interop;
using System.Threading;
using System.Timers;
using System.Windows.Input;

using System.IO;
using System.Diagnostics;
using LED.Properties;
using System.Reflection;
using System.IO.Ports;

using NationalInstruments;
using NationalInstruments.DAQmx;

// IDEEN - VORSCHLÄGE
namespace LED
{

    public partial class F_LED_Signalgeber : Form
    {
        // Flags

        public static bool getanmeldung = false;
        bool bitStatusPrüfung = false;
        bool bitRegelung = false;
        bool bitupdown_manuell = false;
        bool bitSetIndicatorLamp = false;
        bool bitSetLaser = false;
        bool bitSetConfigurationDMM_U = false;
        bool bitSetConfigurationDMM_I = false;
        public static bool bitAbbruch = false;
        bool bitAbschaltung = false;
        public static bool getbitAbschaltung_visuell = false;
        bool bittimerTaste = false;
        public static bool bitMessageBox = false;
        public static bool bitSV_CC_Mode = false;
        bool bitPrfgAb = false;
        bool existSerialnummer = false;
        
        // Variablen allgemein

        public static int time = 0;
        public static int timemax = 0;
        int aPhasen = 0;                            // Anzahl der Phasen eines Prüfschritts
        int zPhasen = 1;                            // Zähler der Phasen eines Prüfschritts
        int zTaste = 0;

        //Variablen DigIO

        int OutDigIO = 0;
        int SumOutDigIO = 0;
        int PrüfOutDigIO = 0;
        int mask = 255;
        byte set;
        byte prüf;

        // Variablen Messequipment

       // string ADR_DMM_U = "01";                    // Adresse Spannungsmesser
       // string ADR_DMM_I = "02";                    // Adress Strommesser

        string ADR_DMM_U = "";
        string ADR_DMM_I = "";
        string ADR_SV_ACDC = "";                  // Adresse SV_ACDC

        string MessEquipment = "";
        public string[,] InformationMessequipmentDB;      // [Hersteller, Typ, Serial]
        public string[] Manufacturer;
        string IDN_String = "";
        string[] MessEquipmentIDN;                 // [Manufacturer, Model, Serial, Firmware]
        bool[] bitcheckMessEquipment;

        string RangeDMM_U = "0";                         // Messbereich Spannungsmesser
        string SetRangeDMM_U = "1000";                   // gesetzter Messbereich Spannungsmesser
        string RangeDMM_I = "0";                         // Messbereich Strommesser
        string SetRangeDMM_I = "3";                      // Messbereich Spannungsmesser       
        string Res = "";                                 // Resolution Spannungs- oder Strommesser
        string DETBAND = "200";

        string STB_SV_ACDC = "";
        byte STB = 0;
        bool AuswertungBit = false;
        int BitNumber = 0;

        // Variablen Parametersatz aus DB

        string LastSachnummer = "";
        string Sachnummer = "";
        string Typenbezeichnung = "";
        long ID_Parametersatz;
        public int aPrüfschritte = 0;        // Anzahl Prüfschritte
        public static int zPrüfschritt = 1;         // Zähler Prüfschritt
        string setUart = "";                        // String für Spannungsart, AC oder DC
        double LimesTestID = 0;
        
        // Variablen Messequipment aus DB

        int ID_DMM_U = 0;
        int ID_DMM_I = 0;
        int ID_SV_ACDC = 0;
        int ID_Therm = 0;
        string[] IdentificationEquipment;
        string Interface_DMM_U = "";
        string Interface_DMM_I = "";

        // Arrays Parametersatz

        double[] ID_Pschritt;
        double[] U1;
        double[] U2;
        double[] lowtolU1;
        double[] uptolU1;
        double[] lowtol2;
        double[] uptol2;
        double[] I1min;
        double[] I1max;
        double[] I2min;
        double[] I2max;
        string[] Rp;
        double[] Sweite1;
        double[] Sweite2;
        string[] Uart;
        string[] Wave;
        double[] f;
        string[] VP;
        string[] Lichtstärke;
        string[] Farbort;
        bool[] PrfgU;
        bool[] PrfgI;
        bool[] PrfgVi;
        bool[] PrfgAb;
        bool[] PrfgCo;

        // Membervariablen für Datensatz aus DB

        double m_U = 0;
        double m_Umin = 0;                              // untere Spannungsgrenze
        double m_Umax = 0;                              // obere Spannungsgrenze 
        double m_lowtol = 0;
        double m_uptol = 0;
        double m_I = 0;
        double m_Imin = 0;
        double m_Imax = 0;
        double m_SWeite = 0;                            // Schrittweite bei Spannungseinstellung

        //double m_U2 = 0;
        //double m_U2min = 0;                              // untere Spannungsgrenze
        //double m_U2max = 0;                              // obere Spannungsgrenze        
        //double m_lowtol2 = 0;
        //double m_uptol2 = 0;
        //double m_I2min = 0;
        //double m_I2max = 0;
        //double m_SWeite2 = 0;                            // Schrittweite bei Spannungseinstellung  
     
        string m_Rp = "";
        string m_Uart = "";
        double m_Freq = 0;
        bool m_PrfgU = false;
        bool m_PrfgI = false;
        bool m_PrfgVi = false;
        bool m_PrfgAb = false;
        bool m_PrfgCo = false;

        // Arrays Messergebnisse und Berwerung Prüfschritte

        double[,] ErgUI;
        string[,] ErgPrfg;                        // Ergebnisse PrfgU, PrfgI, PrfgCo, PrfgAb, PrfgVi 

        public static double[,] ErgLED136Mat;
        public static double[] ErgAbweichxy;

        // Variablen Bewertung Prüfschritt

        string ErgPrfgU = "";
        string ErgPrfgI = "";
        string ErgPrfgVi = "";
        string ErgPrfgAb = "";
        string ErgPrfgCo = "";

        // Variablen Messergebnisse

        string Serialnummer = "";
        string Auftragsnummer = "";
        DateTime Datum_Uhrzeit;
        int ID_Messequipment = 0;
        public static string ID_Prüfer = "";

        // Variablen Rücklesen Messergebnisse

        string ReadCurr = "";
        double Imess = 0;
        double Idiff = 0;
        string ReadVolt = "";
        double Umess = 0;
        double Udiff = 0;

        // Variablen SV setzen

        string setU = "";
        string setI = "";
        string setImax = "";
        string setf = "";
        string setW = "";

        // Config Serialport

        string[] Config;
        string strCOM = "";
        string PortName = "";
        int BaudRate = 0;
        int DataBits = 0;
        StopBits StopBits;
        Parity Parity;

        public static string GetPortName = "";
        public static int GetBaudRate = 0;
        public static int GetDataBits = 0;
        public static StopBits GetStopBits;
        public static Parity GetParity;

        byte[] readserportDigIO;

        // Variablen Temperaturmessung

        double minMessTemp = 0;
        double maxMessTemp = 50;
        double lowerTemp = 0;
        double upperTemp = 0;
        double PruefTemp = 0;

        // Variablen GPIB

        Ivi.Visa.Interop.FormattedIO488 ioDMM_U;
        Ivi.Visa.Interop.FormattedIO488 ioDMM_I;
        Ivi.Visa.Interop.FormattedIO488 ioSV_ACDC;

        // Variablen Objekte

        int minsizewidth = 545;
        int minsizeheight = 310;
        int maxsizewidth = 545;
        int maxsizeheight = 690;
        int posActFormx = 0;
        int posActFormy = 0;
        int ActFormWidth = 0;
        int ActFormHeight = 0;

        DialogResult DRM;
        DialogResult DRMC;
        DialogResult DRMCC;
        DialogResult DRMPP;
        public static DialogResult DRR;

        string MboxText = "";
        string MboxTitle = "";

        Form OpenForm;

        double m_ResSV = 0;
        int OldValue = 0;
        int NewValue = 0;

        bool left;
        bool right;
        bool up;
        bool down;
        bool mousebutton = false;
        bool keydown = false;
        bool pageup;
        bool pagedown;
    
        // ************************************************ Funktionen der Form F_LED_Signalgeber ************************************************************

        public F_LED_Signalgeber()
        {
            InitializeComponent();
        }

        public void F_LED_Signalgeber_Load(object sender, EventArgs e)
        {

            LoadID_MessequipmentDB();
            LoadInformationMessequipmentDB();

            ioSV_ACDC = new Ivi.Visa.Interop.FormattedIO488();
            ioDMM_U = new Ivi.Visa.Interop.FormattedIO488();
            ioDMM_I = new Ivi.Visa.Interop.FormattedIO488();

            ResourceManager mgr_SV_ACDC = new ResourceManager();
            ResourceManager mgr_DMM_U = new ResourceManager();
            ResourceManager mgr_DMM_I = new ResourceManager();

            ioSV_ACDC.IO = (IMessage)mgr_SV_ACDC.Open("GPIB1::" + ADR_SV_ACDC + "::INSTR", AccessMode.NO_LOCK, 2000, "");

            switch (Interface_DMM_U)
            {
                case "GPIB":                    
                    ioDMM_U.IO = (IMessage)mgr_DMM_U.Open("GPIB1::" + ADR_DMM_U + "::INSTR", AccessMode.NO_LOCK, 2000, "");
                    break;

                case "TCPIP":
                    ioDMM_U.IO = (IMessage)mgr_DMM_U.Open("TCPIP1::" + ADR_DMM_U + "::inst0::INSTR", AccessMode.NO_LOCK, 2000, "");
                    break;
            }

            switch (Interface_DMM_I)
            {
                case "GPIB":
                    ioDMM_I.IO = (IMessage)mgr_DMM_I.Open("GPIB1::" + ADR_DMM_I + "::INSTR", AccessMode.NO_LOCK, 2000, "");
                    break;

                case "TCPIP":
                    ioDMM_I.IO = (IMessage)mgr_DMM_I.Open("TCPIP1::" + ADR_DMM_I + "::inst0::INSTR", AccessMode.NO_LOCK, 2000, "");
                    break;
            }

            SendKeys.Send("%()");                                   //ALT-Taste drücken, um die Keycodes anzuzeigen
            SendKeys.Send("%()");                                   //ALT-Taste drücken, um Focus von Menüleiste zu entfernen

            InitMessequipment();
        }

        private void F_LED_Signalgeber_Activated(object sender, EventArgs e)
        {
            timerTaste.Start();
        }

        private void F_LED_Signalgeber_FormClosing(object sender, FormClosingEventArgs e)
        {
            notifyIconLED.Visible = false;
        }

        private void F_LED_Signalgeber_FormClosed(object sender, FormClosedEventArgs e)
        { 
            ResetObj();
            ResetOutDigIO();
            ResetDMM_U();
            ResetDMM_I();
            ResetSV_ACDC();
        }

        public void F_LED_Signalgeber_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (bittimerTaste == false)
            {
                timerTaste.Start();
                bittimerTaste = true;

                switch (e.KeyValue)
                {
                    case (char)27:    // Tastendruck "ESC" Presenter

                        if (btnCancel.Enabled == true)
                        {
                            btnAbbruch_Click(sender, e);
                        }
                        break;

                    case (char)33:   // Tastendruck "Pfeil-Taste oben" Presenter

                        if (btnUp.Enabled == true)
                        {
                          //  btnUp_Click(sender, e);
                        }
                        break;

                    case (char)34:   // Tastendruck "Pfeil-Taste unten" Presenter

                        if (btnDown.Enabled == true)
                        {
                           // btnDown_Click(sender, e);
                        }
                        break;

       
                    case (char)38:   // Tastendruck "Pfeil-Taste oben" 

                        if (btnUp.Enabled == true)
                        {
                           // btnUp_Click(sender, e);
                        }
                        break;

                    case (char)40:   // Tastendruck "Pfeil-Taste unten" 

                        if (btnDown.Enabled == true)
                        {
                          //  btnDown_Click(sender, e);
                        }
                        break;
          
                    case (char)65:    // Tastendruck "a" 

                        if (btnCancel.Enabled == true)
                        {
                            btnAbbruch_Click(sender, e);
                        }
                        break;

                    case (char)66:    // Tastendruck "B" Presenter 

                        if (btnLaser.Enabled == true)
                        {
                            SetLaser();
                        }
                        break;

                    case (char)68:   // Tastendruck "d" 

                        if (btnDown.Enabled == true)
                        {
                            btnDown_Click(sender, e);
                        }
                        break;

                    case (char)76:   // Tastendruck "l"  

                        SetLaser();
                        //e.Handled = true;
                        break;

                    case (char)82:   // Tastendruck "r"  

                        btnRepeat_Click(sender, e);
                        break;

                    case (char)83:   // Tastendruck "s" 

                        if (btnStartWeiterSpeichern.Enabled == true &&
                           (btnStartWeiterSpeichern.Text == "&Start" || btnStartWeiterSpeichern.Text == "&Weiter"))
                        {
                            btnStartWeiterSpeichern_Click(sender, e);
                        }
                        break;

                    case (char)85:   // Tastendruck "u"

                        if (btnUp.Enabled == true)
                        {
                            btnUp_Click(sender, e);
                        }
                        break;

                    case (char)87:   // Tastendruck "w" 

                        if (btnStartWeiterSpeichern.Enabled == true && btnStartWeiterSpeichern.Text == "&Weiter")
                        {
                            btnStartWeiterSpeichern_Click(sender, e);
                        }
                        break;

                    case (char)112:   // Tastendruck "F1" 

                        if (btnStartWeiterSpeichern.Enabled == true && btnStartWeiterSpeichern.Text == "&Start")
                        {
                            btnStartWeiterSpeichern_Click(sender, e);
                        }
                        break;

                    case (char)113:   // Tastendruck "F2" 

                        if (cboxiO.Enabled == true)
                        {
                            if (cboxiO.Checked == false)
                            {
                                cboxiO.Checked = true;
                            }
                            else
                            {
                                cboxiO.Checked = false;
                            }

                            cboxiO_CheckedChanged(sender, e);
                        }
                        break;

                    case (char)114:   // Tastendruck "F3" 

                        if (cboxniO.Enabled == true)
                        {
                            if (cboxniO.Checked == false)
                            {
                                cboxniO.Checked = true;
                            }
                            else
                            {
                                cboxniO.Checked = false;
                            }

                            cboxniO_CheckedChanged(sender, e);
                        }
                        break;

                    case (char)115:   // Tastendruck "F4" 

                        if (btnStartWeiterSpeichern.Enabled == true &&
                           (btnStartWeiterSpeichern.Text == "&Weiter" || btnStartWeiterSpeichern.Text == "&Speichern"))
                        {
                            btnStartWeiterSpeichern_Click(sender, e);
                        }
                        break;

                    case (char)116:    // Tastendruck "F5" 

                        if (btnCancel.Enabled == true)
                        {
                            btnAbbruch_Click(sender, e);
                        }
                        break;

                    case (char)117:    // Tastendruck "F6" 

                        if (btnLaser.Enabled == true)
                        {
                            SetLaser();
                        }
                        break;

                    case (char)174:   // Tastendruck "-" Presenter 

                        if (cboxiO.Enabled || cboxniO.Enabled)
                        {
                            ++zTaste;

                            switch (zTaste)
                            {
                                case 1:
                                    cboxiO.Checked = true;
                                    AppraisalVisualTest();
                                    break;

                                case 2:
                                    cboxiO.Checked = false;
                                    AppraisalVisualTest();
                                    break;

                                case 3:
                                    cboxniO.Checked = true;
                                    AppraisalVisualTest();
                                    break;

                                case 4:
                                    cboxniO.Checked = false;
                                    AppraisalVisualTest();
                                    zTaste = 0;
                                    break;

                                default:
                                    break;
                            }                            
                        }
                        break;

                    case (char)175:   // Tastendruck "+" Presenter 

                        if (btnStartWeiterSpeichern.Enabled == true)
                        {
                            btnStartWeiterSpeichern_Click(sender, e);
                        }
                        break;

                    default:
                        break;
                }
                textBox1.Text = Convert.ToString(e.KeyValue);
            }
        }

        private void ReadPosActivForm()
        {
            ActFormWidth = ActiveForm.Bounds.Width;
            ActFormHeight = ActiveForm.Bounds.Height;
            posActFormx = ActiveForm.Location.X;
            posActFormy = ActiveForm.Location.Y;
        }

        // ****************************************************** Übernahme Parameter aus DB *****************************************************************

        private void LoadID_MessequipmentDB()
        {
            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataReader reader;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   //"Data Source=C:\\Daten\\Datenbank LED\\DB_LED-Signalgeber.mdb";
                                   //"Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb";
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb;" +
                                   "Jet OLEDB:Database Password=Goniometer";

            cmd.Connection = con;
            cmd.CommandText = "select * from Messequipment where Freigabe = true";

            try
            {
                con.Open();
                reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    ID_Messequipment = Convert.ToInt16(reader["ID_Messequipment"]);
                    ID_DMM_U = Convert.ToInt16(reader["ID_DMM_U"]);
                    ID_DMM_I = Convert.ToInt16(reader["ID_DMM_I"]);
                    ID_SV_ACDC = Convert.ToInt16(reader["ID_SV_ACDC"]);
                    ID_Therm = Convert.ToInt16(reader["ID_Therm"]);
                }
                reader.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message, "LoadID_MessequipmentDB", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            F_Einstellungen F_Einstellungen = new F_Einstellungen();
            F_Einstellungen.GetIDMessequipment(ID_Messequipment);
        }

        private void LoadInformationMessequipmentDB()
        {
            int i = 0;
            IdentificationEquipment = new string[4] { "DMM_U", "DMM_I", "SV_ACDC", "Therm" };
            InformationMessequipmentDB = new string[4, 5];

            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataReader reader;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +                                  
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb;" +
                                   "Jet OLEDB:Database Password=Goniometer";
            cmd.Connection = con;

            while (i <= 3)
            {
                switch (IdentificationEquipment[i])
                {
                    case "DMM_U":
                        cmd.CommandText = "select * from DMM_Spannung where ID_DMM_U like '" + ID_DMM_U + "'";
                        break;

                    case "DMM_I":
                        cmd.CommandText = "select * from DMM_Strom where ID_DMM_I like '" + ID_DMM_I + "'";
                        break;

                    case "SV_ACDC":
                        cmd.CommandText = "select * from SV_ACDC where ID_SV_ACDC like '" + ID_SV_ACDC + "'";
                        break;

                    case "Therm":
                        cmd.CommandText = "select * from Thermometer where ID_Therm like '" + ID_Therm + "'";
                        break;

                    default:
                        break;
                }

                try
                {
                    con.Open();
                    reader = cmd.ExecuteReader();

                    while (reader.Read())
                    {
                        InformationMessequipmentDB[i, 0] = Convert.ToString((reader["Hersteller"]));
                        InformationMessequipmentDB[i, 0] = InformationMessequipmentDB[i, 0].ToUpper();
                        InformationMessequipmentDB[i, 1] = Convert.ToString((reader["Typ"]));
                        InformationMessequipmentDB[i, 1] = InformationMessequipmentDB[i, 1].ToUpper();
                        InformationMessequipmentDB[i, 2] = Convert.ToString((reader["Serial"]));
                        InformationMessequipmentDB[i, 2] = InformationMessequipmentDB[i, 2].ToUpper();
                        InformationMessequipmentDB[i, 3] = Convert.ToString((reader["Interface"]));
                        InformationMessequipmentDB[i, 3] = InformationMessequipmentDB[i, 3].ToUpper();
                        InformationMessequipmentDB[i, 4] = Convert.ToString((reader["Adr"]));
                        InformationMessequipmentDB[i, 4] = InformationMessequipmentDB[i, 4].ToUpper();
                    }
                    reader.Close();
                    con.Close();

                    i = i + 1;
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show("Die Verbindung ist bereits offen. \n" +
                                    "Source: " + ex.Source + "\n" +
                                    "Message: " + ex.Message + "\n",
                                    "CheckIDMessequipment " + MessEquipment, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (OleDbException ex)
                {
                    MessageBox.Show("Beim Öffnen der Verbindung ist ein Fehler auf Verbindungsebene aufgetreten.\n" +
                                    "Source: " + ex.Source + "\n" +
                                    "Message: " + ex.Message + "\n",
                                    "CheckIDMessequipment " + MessEquipment, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            Interface_DMM_U = InformationMessequipmentDB[0, 3];
            Interface_DMM_I = InformationMessequipmentDB[1, 3];

            ADR_DMM_U = InformationMessequipmentDB[0, 4];   
            ADR_DMM_I = InformationMessequipmentDB[1, 4];
            ADR_SV_ACDC = InformationMessequipmentDB[2, 4];
        }

        private void LoadSachnummerDB()
        {
            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataReader reader;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   //"Data Source=C:\\Daten\\Datenbank LED\\DB_LED-Signalgeber.mdb";
                                   //"Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb";
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb;" +
                                   "Jet OLEDB:Database Password=Goniometer";

            cmd.Connection = con;
            cmd.CommandText = "select * from Signalgeber where Freigabe = true";

            try
            {
                con.Open();
                reader = cmd.ExecuteReader();
                coboxSachnummer.Items.Clear();

                while (reader.Read())
                {
                    coboxSachnummer.Items.Add(reader["Sachnummer"]);
                }
                reader.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message, "LoadSachnummerDB", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadIDParametersatzDB()
        {
            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataReader reader;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   //"Data Source=C:\\Daten\\Datenbank LED\\DB_LED-Signalgeber.mdb";
                                   //"Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb";
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb;" +
                                   "Jet OLEDB:Database Password=Goniometer";
            cmd.Connection = con;
            cmd.CommandText = "select * from Signalgeber where Sachnummer = '" + Sachnummer + "'";

            try
            {
                con.Open();
                reader = cmd.ExecuteReader();
                coboxSachnummer.Items.Clear();

                reader.Read();

                ID_Parametersatz = Convert.ToInt32(reader["ID_Parametersatz"]);
                Typenbezeichnung = Convert.ToString(reader["Typenbezeichnung"]);

                reader.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message, "LoadIDParametersatzDB", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadParametersatzDB()
        {
            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataReader reader;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   //"Data Source=C:\\Daten\\Datenbank LED\\DB_LED-Signalgeber.mdb";
                                   //"Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb";
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb;" +
                                   "Jet OLEDB:Database Password=Goniometer";
            cmd.Connection = con;
            cmd.CommandText = "select * from Parametersätze where ID_Parametersatz like '" + ID_Parametersatz + "'";


            try
            {
                con.Open();
                reader = cmd.ExecuteReader();
                reader.Read();

                //Anzahl der Prüfschritte laden
                aPrüfschritte = Convert.ToInt32(reader["Prüfschritte"]);

                // Parameter für Temp. laden
                lowerTemp = Convert.ToDouble(reader["lowerTemp"]);
                tbxlowerTemp.Text = Convert.ToString(lowerTemp);
                upperTemp = Convert.ToDouble(reader["upperTemp"]);
                tbxupperTemp.Text = Convert.ToString(upperTemp);

                // ID aller benötigten Prüfschritte einlesen 

                ID_Pschritt = new double[aPrüfschritte + 1];

                for (int i = 1; i <= aPrüfschritte; i++)
                {
                    ID_Pschritt[i] = Convert.ToDouble(reader["Pschritt" + i]);
                }

                reader.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message, "LoadParametersatzDB", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadPrüfschritteDB()
        {

            U1 = new double[aPrüfschritte + 1];
            lowtolU1 = new double[aPrüfschritte + 1];
            uptolU1 = new double[aPrüfschritte + 1];
            I1min = new double[aPrüfschritte + 1];
            I1max = new double[aPrüfschritte + 1];
            Sweite1 = new double[aPrüfschritte + 1];                  //
            U2 = new double[aPrüfschritte + 1];
            lowtol2 = new double[aPrüfschritte + 1];
            uptol2 = new double[aPrüfschritte + 1];
            I2min = new double[aPrüfschritte + 1];
            I2max = new double[aPrüfschritte + 1];                      //
            Sweite2 = new double[aPrüfschritte + 1];
            Rp = new string[aPrüfschritte + 1];
            Uart = new string[aPrüfschritte + 1];
            Wave = new string[aPrüfschritte + 1];
            f = new double[aPrüfschritte + 1];
            VP = new string[aPrüfschritte + 1];
            Lichtstärke = new string[aPrüfschritte + 1];
            Farbort = new string[aPrüfschritte + 1];
            PrfgU = new bool[aPrüfschritte + 1];
            PrfgI = new bool[aPrüfschritte + 1];
            PrfgVi = new bool[aPrüfschritte + 1];
            PrfgAb = new bool[aPrüfschritte + 1];
            PrfgCo = new bool[aPrüfschritte + 1];

            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataReader reader;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   //"Data Source=C:\\Daten\\Datenbank LED\\DB_LED-Signalgeber.mdb";
                                   //"Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb";
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb;" +
                                   "Jet OLEDB:Database Password=Goniometer";
            try
            {
                for (int i = 1; i <= aPrüfschritte; i++)
                {
                    cmd.Connection = con;
                    cmd.CommandText = "select * from Pschritte where ID_Pschritt like '" + ID_Pschritt[i] + "'";

                    con.Open();
                    reader = cmd.ExecuteReader();
                    reader.Read();

                    U1[i] = Convert.ToDouble(reader["U1"]);
                    lowtolU1[i] = Convert.ToDouble(reader["lowtolU1"]);
                    uptolU1[i] = Convert.ToDouble(reader["uptolU1"]);
                    I1min[i] = Convert.ToDouble(reader["I1min"]);
                    I1max[i] = Convert.ToDouble(reader["I1max"]);
                    Sweite1[i] = Convert.ToDouble(reader["Sweite1"]);
                    U2[i] = Convert.ToDouble(reader["U2"]);
                    lowtol2[i] = Convert.ToDouble(reader["lowtolU2"]);
                    uptol2[i] = Convert.ToDouble(reader["uptolU2"]);
                    I2min[i] = Convert.ToDouble(reader["I2min"]);
                    I2max[i] = Convert.ToDouble(reader["I2max"]);
                    Sweite2[i] = Convert.ToDouble(reader["Sweite2"]);
                    Rp[i] = Convert.ToString(reader["Rp"]);
                    Rp[i] = Rp[i].ToUpper();
                    Uart[i] = Convert.ToString(reader["Uart"]);
                    Uart[i] = Uart[i].ToUpper();
                    Wave[i] = Convert.ToString(reader["Wave"]);
                    f[i] = Convert.ToDouble(reader["f"]);
                    VP[i] = Convert.ToString(reader["VP"]);
                    Lichtstärke[i] = Convert.ToString(reader["Lichtstärke"]);
                    Farbort[i] = Convert.ToString(reader["Farbort"]);
                    PrfgU[i] = Convert.ToBoolean(reader["PrfgU"]);
                    PrfgI[i] = Convert.ToBoolean(reader["PrfgI"]);
                    PrfgVi[i] = Convert.ToBoolean(reader["PrfgVi"]);
                    PrfgAb[i] = Convert.ToBoolean(reader["PrfgAb"]);
                    PrfgCo[i] = Convert.ToBoolean(reader["PrfgCo"]);

                    reader.Close();
                    con.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message, "LoadPrüfschritteDB", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadLimesTestIDDB()
        {
            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataReader reader;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse.mdb";
            
            cmd.Connection = con;
            cmd.CommandText = "SELECT * FROM TestData WHERE TestName = '" + Serialnummer + "'";

            try
            {
                con.Open();
                reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    LimesTestID = Convert.ToDouble(reader["TestID"]);
                    textBox6.Text = Convert.ToString(LimesTestID);
                }
                reader.Close();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("" + ex.Message, "LoadLimesTestID", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CheckSerialnummerExist()
        {
            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            OleDbDataReader reader;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb;" +
                                   "Jet OLEDB:Database Password=Goniometer";
            cmd.Connection = con;

            if (Typenbezeichnung == "LED 136 Matrix-BG")
                cmd.CommandText = "select * ErgLED136MatrixBG where Serialnummer = '" + Serialnummer + "'"; 
            else
                cmd.CommandText = "select * from Messergebnisse where Serialnummer = '" + Serialnummer + "'";

            try
            {
                con.Open();
                reader = cmd.ExecuteReader();

                reader.Read();

                string ExistSerial = Convert.ToString(reader["Serialnummer"]);

                if (Serialnummer == ExistSerial )
                    existSerialnummer = true;

                reader.Close();
                con.Close();
            }
            catch
            {
                existSerialnummer = false;
            }
        }

        // ******************************************************* Funktionen Messequipment ******************************************************************

        // Initialisierung Messequipment

        private void InitDMM_U()
        {
            MessEquipment = "DMM_U";
            MessEquipmentIDN = new string[4];

            try
            {
                // Reset DMM_U
                ioDMM_U.WriteString("*RST", true);
                // Abfrage ID
                ioDMM_U.WriteString("*IDN?", true);
                IDN_String = ioDMM_U.ReadString();
                IDN_String = IDN_String.ToUpper();
                MessEquipmentIDN = IDN_String.Split(new char[] { ',' });

                CheckMessequipment(MessEquipment, MessEquipmentIDN);

                ErrorDMM_U();
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("Die Initialisierung des Spannungsmessers ist fehlgeschlagen. \n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Initialisierung (InitDMM_U)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }

        private void InitDMM_I()
        {
            MessEquipment = "DMM_I";

            try
            {
                // Reset DMM_I
                ioDMM_I.WriteString("*RST", true);
                // Abfrage ID
                ioDMM_I.WriteString("*IDN?", true);
                IDN_String = ioDMM_I.ReadString();
                IDN_String = IDN_String.ToUpper();
                MessEquipmentIDN = IDN_String.Split(new char[] {  ',' });

                CheckMessequipment(MessEquipment, MessEquipmentIDN);

                ErrorDMM_I();
            }

            catch (ArgumentException ex)
            {
                MessageBox.Show("Die Initialisierung des Strommessers ist fehlgeschlagen. \n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Initialisierung (InitDMM_I)", MessageBoxButtons.OK, MessageBoxIcon.Error);

                Close();
            }
        }

        private void InitSV_ACDC()
        {
            MessEquipment = "SV_ACDC";
            string Manufacturer = "";

            try
            {
                ioSV_ACDC.WriteString("*IDN?", true);
                IDN_String = ioSV_ACDC.ReadString();
                Manufacturer = IDN_String.Substring(0, 14);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                MessageBox.Show("Fehler bei der Bearbeitung des IDN-String.\n\n" + ex, "InitSV_ACDC");
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("Die Textzeichenfolge enthält ein ungültiges Ersatzzeichenpaar. \n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Initialisierung (Init_SV_ACDC)", MessageBoxButtons.OK, MessageBoxIcon.Error);

                InitSV_ACDC();
            }
            catch (SystemException ex)
            {
                MessageBox.Show("Die Initialisierung der Stromversorgung ist fehlgeschlagen. \n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Initialisierung (Init_SV_ACDC)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                InitSV_ACDC();
            }

            if (Manufacturer == "EAC-S 50/15V 3")
            {
                try
                {
                    // Reset Stromversorgung
                    ioSV_ACDC.WriteString("*RST", true);

                    // Go to Remote
                    ioSV_ACDC.WriteString("GTR", true);

                    // Abfrage ID
                    ioSV_ACDC.WriteString("*IDN?", true);
                    IDN_String = ioSV_ACDC.ReadString();
                    MessEquipmentIDN[0] = "SCHULTZ ELECTRONIC";
                    MessEquipmentIDN[1] = IDN_String;
                    MessEquipmentIDN[1] = MessEquipmentIDN[1].Remove(18);
                    MessEquipmentIDN[2] = "0710000316";

                    CheckMessequipment(MessEquipment, MessEquipmentIDN);

                }

                catch (ArgumentException ex)
                {
                    MessageBox.Show("Die Initialisierung der Stromversorgung ist fehlgeschlagen. \n" +
                                    "Source: " + ex.Source + "\n" +
                                    "Message: " + ex.Message + "\n",
                                    "Fehler Initialisierung (Init_SV_ACDC)", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    InitSV_ACDC();
                }
            }
            else if (Manufacturer == "HBS-Electronic")
            {
                try
                {
                    // Reset Stromversorgung
                    ioSV_ACDC.WriteString("*RST", true);

                    // Abfrage ID
                    ioSV_ACDC.WriteString("*IDN?", true);
                    IDN_String = ioSV_ACDC.ReadString();
                    MessEquipmentIDN = IDN_String.Split(new char[] { ' ', ',' });

                    CheckMessequipment(MessEquipment, MessEquipmentIDN);

                }

                catch (ArgumentException ex)
                {
                    MessageBox.Show("Die Initialisierung der Stromversorgung ist fehlgeschlagen. \n" +
                                    "Source: " + ex.Source + "\n" +
                                    "Message: " + ex.Message + "\n",
                                    "Fehler Initialisierung (Init_SV_ACDC)", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    InitSV_ACDC();
                }
            }
            else
            {
                MessageBox.Show("Es ist keine der in der Datenbank freigegeben \n" +
                                "Stromversorgungen am Rechner angeschlossen." +
                                "Bitte schließen Sie eine freigegebene Stromversorgung " +
                                "an und starten das Prüfprogramm erneut.", "Fehler Stromversorgung(InitSV_ACDC)", MessageBoxButtons.OK, MessageBoxIcon.Error);

                Close();
            }
        }

        private void Init_Ni_USBTC01()
        {
            MessEquipment = "Therm";
            MessEquipmentIDN = new string[4];

            try
            {
                Task analogInTask = new Task();

                AIChannel TempAIChannel;

                TempAIChannel = analogInTask.AIChannels.CreateThermocoupleChannel("Dev1/ai0", "Temperature", minMessTemp, maxMessTemp,
                                                                                   AIThermocoupleType.J, AITemperatureUnits.DegreesC);

                AnalogSingleChannelReader reader = new AnalogSingleChannelReader(analogInTask.Stream);

                Device dev = DaqSystem.Local.LoadDevice("dev1");

                PruefTemp = reader.ReadSingleSample();
                dev.SelfTest();
                MessEquipmentIDN[0] = "NATIONAL INSTRUMENTS";
                MessEquipmentIDN[1] = dev.ProductType.ToString();        // Antwort: USB-TC01
                MessEquipmentIDN[2] = dev.SerialNumber.ToString("X");    // Antwort: 174FCD3  

                CheckMessequipment(MessEquipment, MessEquipmentIDN);
            }
            catch (SystemException ex)
            {
                MessageBox.Show("Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Thermometer (Init_Ni_USBTC01)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitDigIO()
        {
            SetOutDigIO(-(SumOutDigIO));
            serPortDigIO_AusgabeRückanwort();
        }

        // Überprüfung Messequipment

        private void InitMessequipment()
        {
            bitcheckMessEquipment = new bool[4];
            Manufacturer = new string[4];

         //   LoadID_MessequipmentDB();
         //   LoadInformationMessequipmentDB();

            ResetSV_ACDC();
            InitDMM_U();
            InitDMM_I();
            InitSV_ACDC();
            Init_Ni_USBTC01();
            SetConfigSerPort();
            InitDigIO();

            if (!(bitcheckMessEquipment[0] && bitcheckMessEquipment[1] &&
                  bitcheckMessEquipment[2] && bitcheckMessEquipment[3]))
            {
                tsmiDatei.Enabled = true;
                tsmiDateiAnmelden.Enabled = false;
            }
            else
            {
                tsmiDatei.Enabled = true;
                tsmiDateiAnmelden.Enabled = true;
            }
        }

        private void CheckMessequipment(string MessEquipment, string[] MessEquipmentIDN)
        {
            int z = 0;
            this.MessEquipment = MessEquipment;

            try
            {
                switch (this.MessEquipment)
                {
                    case "DMM_U":

                        if (this.MessEquipmentIDN[0].ToUpper() == "HEWLETT-PACKARD")
                        {
                            Manufacturer[0] = MessEquipmentIDN[0].ToUpper();

                            if (!(this.MessEquipmentIDN[0] == InformationMessequipmentDB[z, 0] &&            // Hersteller DMM_U
                                  this.MessEquipmentIDN[1] == InformationMessequipmentDB[z, 1]))             // Typ DMM_U 
                            {
                                MessageBox.Show("Es wird das DMM mit den geforderten Parametern \n" +
                                                "Hersteller: " + InformationMessequipmentDB[z, 0] + "\n" +
                                                "Typ:        " + InformationMessequipmentDB[z, 1] + "\n" +
                                                "Serial:     " + InformationMessequipmentDB[z, 2] + "\n" +
                                                "verwendet!!!" + "\n" +
                                                "Bevor Sie fortfahren können, müssen Sie das geforderte DMM" + "\n" +
                                                "an den Rechner anschliessen oder einen anderen Datensatz in" + "\n" +
                                                "der Datenbank freigeben." + "\n" +
                                                "Danach ist eine Systeminitialisierung durchzuführen und  Anmeldung am Prüfprogramm möglich.",
                                                "Fehler Check Spannungsmesser (CheckMessequipment)", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

                                bitcheckMessEquipment[0] = false;
                            }
                            else
                            {
                                bitcheckMessEquipment[0] = true;
                            }
                        }
                        else if (this.MessEquipmentIDN[0].ToUpper() == "KEYSIGHT TECHNOLOGIES")
                        {
                            Manufacturer[0] = MessEquipmentIDN[0].ToUpper();

                            if (!(this.MessEquipmentIDN[0] == InformationMessequipmentDB[z, 0] &&            // Hersteller DMM_U
                                  this.MessEquipmentIDN[1] == InformationMessequipmentDB[z, 1] &&            // Typ DMM_U
                                  this.MessEquipmentIDN[2] == InformationMessequipmentDB[z, 2]))             // Serial DMM_U
                            {
                                MessageBox.Show("Es wird das DMM mit den geforderten Parametern \n" +
                                                "Hersteller: " + InformationMessequipmentDB[z, 0] + "\n" +
                                                "Typ:        " + InformationMessequipmentDB[z, 1] + "\n" +
                                                "Serial:     " + InformationMessequipmentDB[z, 2] + "\n" +
                                                "verwendet!!!" + "\n" +
                                                "Bevor Sie fortfahren können, müssen Sie das geforderte DMM" + "\n" +
                                                "an den Rechner anschliessen oder einen anderen Datensatz in" + "\n" +
                                                "der Datenbank freigeben." + "\n" +
                                                "Danach ist eine Anmeldung am Prüfprogramm möglich.",
                                                "Fehler Check Spannungsmesser (CheckMessequipment)", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

                                bitcheckMessEquipment[0] = false;
                            }
                            else
                            {
                                bitcheckMessEquipment[0] = true;
                            }
                        }
                        else
                        {
                            bitcheckMessEquipment[0] = false;
                        }
                        break;

                    case "DMM_I":

                        z = 1;

                        if (this.MessEquipmentIDN[0].ToUpper() == "HEWLETT-PACKARD")
                        {
                            Manufacturer[1] = MessEquipmentIDN[0].ToUpper();

                            if (!(this.MessEquipmentIDN[0] == InformationMessequipmentDB[z, 0] &&            // Hersteller DMM_I
                                  this.MessEquipmentIDN[1] == InformationMessequipmentDB[z, 1]))             // Typ DMM_I
                            {
                                MessageBox.Show("Es wird das DMM mit den geforderten Parametern \n" +
                                                "Hersteller: " + InformationMessequipmentDB[z, 0] + "\n" +
                                                "Typ:        " + InformationMessequipmentDB[z, 1] + "\n" +
                                                "Serial:     " + InformationMessequipmentDB[z, 2] + "\n" +
                                                "verwendet!!!" + "\n" +
                                                "Bevor Sie fortfahren können, müssen Sie das geforderte DMM" + "\n" +
                                                "an den Rechner anschliessen oder einen anderen Datensatz in" + "\n" +
                                                "der Datenbank freigeben." + "\n" +
                                                "Danach ist eine Anmeldung am Prüfprogramm möglich.",
                                                "Fehler Check Strommesser (CheckMessequipment)", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

                                bitcheckMessEquipment[1] = false;
                            }
                            else
                            {
                                bitcheckMessEquipment[1] = true;
                            }
                        }
                        else if (this.MessEquipmentIDN[0].ToUpper() == "KEYSIGHT TECHNOLOGIES")
                        {
                            Manufacturer[1] = MessEquipmentIDN[0].ToUpper();

                            if (!(this.MessEquipmentIDN[0] == InformationMessequipmentDB[z, 0] &&            // Hersteller DMM_I
                                  this.MessEquipmentIDN[1] == InformationMessequipmentDB[z, 1] &&            // Typ DMM_I
                                  this.MessEquipmentIDN[2] == InformationMessequipmentDB[z, 2]))             // Serial DMM_I
                            {
                                MessageBox.Show("Es wird das DMM mit den geforderten Parametern \n" +
                                                "Hersteller: " + InformationMessequipmentDB[z, 0] + "\n" +
                                                "Typ:        " + InformationMessequipmentDB[z, 1] + "\n" +
                                                "Serial:     " + InformationMessequipmentDB[z, 2] + "\n" +
                                                "verwendet!!!" + "\n" +
                                                "Bevor Sie fortfahren können, müssen Sie das geforderte DMM" + "\n" +
                                                "an den Rechner anschliessen oder einen anderen Datensatz in" + "\n" +
                                                "der Datenbank freigeben." + "\n" +
                                                "Danach ist eine Anmeldung am Prüfprogramm möglich.",
                                                "Fehler Check Strommesser (CheckMessequipment)", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

                                bitcheckMessEquipment[1] = false;

                            }
                            else
                            {
                                bitcheckMessEquipment[1] = true;
                            }
                        }
                        else
                        {
                            bitcheckMessEquipment[1] = false;
                        }
                        break;

                    case "SV_ACDC":

                        z = 2;

                        if (this.MessEquipmentIDN[0].ToUpper() == "SCHULTZ ELECTRONIC")
                        {
                            Manufacturer[2] = MessEquipmentIDN[0].ToUpper();

                            if (this.MessEquipmentIDN[1] != InformationMessequipmentDB[z, 1])             // Typ SV
                            {
                                MessageBox.Show("Es wird nicht die Stromversorgung mit den geforderten Parametern \n" +
                                                "Hersteller: " + InformationMessequipmentDB[z, 0] + "\n" +
                                                "Typ:        " + InformationMessequipmentDB[z, 1] + "\n" +
                                                "Serial:     " + InformationMessequipmentDB[z, 2] + "\n" +
                                                "verwendet!!!" + "\n" +
                                                "Bevor Sie fortfahren können, müssen Sie die geforderte" + "\n" +
                                                "Stromversorgung an den Rechner anschliessen oder einen" + "\n" +
                                                "anderen Datensatz in der Datenbank freigeben." + "\n" +
                                                "Danach ist eine Anmeldung am Prüfprogramm möglich.",
                                                "Fehler Check SV (CheckMessequipment)", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

                                bitcheckMessEquipment[2] = false;
                            }
                            else
                            {
                                bitcheckMessEquipment[2] = true;
                            }
                        }

                        else if (this.MessEquipmentIDN[0].ToUpper() == "HBS-ELECTRONIC")
                        {
                            Manufacturer[2] = MessEquipmentIDN[0].ToUpper();

                            if (!(this.MessEquipmentIDN[0].ToUpper() == InformationMessequipmentDB[z, 0] &&  // Hersteller SV_ACDC
                                  this.MessEquipmentIDN[1] == InformationMessequipmentDB[z, 1] &&            // Typ SV_ACDC
                                  this.MessEquipmentIDN[2] == InformationMessequipmentDB[z, 2]))             // Serial SV_ACDC
                            {
                                MessageBox.Show("Es wird nicht die Stromversorgung mit den geforderten Parametern \n" +
                                                "Hersteller: " + InformationMessequipmentDB[z, 0] + "\n" +
                                                "Typ:        " + InformationMessequipmentDB[z, 1] + "\n" +
                                                "Serial:     " + InformationMessequipmentDB[z, 2] + "\n" +
                                                "verwendet!!!" + "\n" +
                                                "Bevor Sie fortfahren können, müssen Sie die geforderte" + "\n" +
                                                "Stromversorgung an den Rechner anschliessen oder einen" + "\n" +
                                                "anderen Datensatz in der Datenbank freigeben." + "\n" +
                                                "Danach ist eine Anmeldung am Prüfprogramm möglich.",
                                                "Fehler Check SV_ACDC (CheckMessequipment)", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);



                                bitcheckMessEquipment[2] = false;
                            }
                            else
                            {
                                bitcheckMessEquipment[2] = true;
                            }
                        }
                        break;

                    case "Therm":

                        z = 3;

                        if (!(this.MessEquipmentIDN[0] == InformationMessequipmentDB[z, 0] &&            // Hersteller Thermometer
                              this.MessEquipmentIDN[1] == InformationMessequipmentDB[z, 1] &&            // Typ Thermometer
                              this.MessEquipmentIDN[2] == InformationMessequipmentDB[z, 2]))             // Serial Thermometer
                        {
                            MessageBox.Show("Es wird nicht das Thermometer mit den geforderten Parametern \n" +
                                            "Hersteller: " + InformationMessequipmentDB[z, 0] + "\n" +
                                            "Typ:        " + InformationMessequipmentDB[z, 1] + "\n" +
                                            "Serial:     " + InformationMessequipmentDB[z, 2] + "\n" +
                                            "verwendet!!!" + "\n" +
                                            "Bevor Sie fortfahren können, müssen Sie das geforderte" + "\n" +
                                            "Thermometer an den Rechner anschliessen oder einen" + "\n" +
                                            "anderen Datensatz in der Datenbank freigeben." + "\n" +
                                            "Danach ist eine Anmeldung am Prüfprogramm möglich.",
                                            "Fehler Check Thermometer (CheckMessequipment)", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

                            bitcheckMessEquipment[3] = false;
                        }
                        else
                        {
                            bitcheckMessEquipment[3] = true;
                        }
                        break;

                    default:
                        break;
                }
            }
            catch
            {
                MessageBox.Show("Die Überprüfung des Messequipments " + MessEquipment + " ist fehlgeschlagen.", "CheckMessequipment");
            }
        }

        // Fehlerabfrage Messequipment

        private void ErrorDMM_U()
        {
            string Readerror = "";
            int PositionError = -1;
            string ReadTerm = "";

            try
            {
                while (PositionError == -1)
                {
                    // Abfrage ID und Abfrage ob Fehler vorhanden
                    ioDMM_U.WriteString("SYST:ERR?", true);
                    Readerror = ioDMM_U.ReadString();
                    PositionError = Readerror.IndexOf("No error");

                    if (PositionError == -1)
                    {
                        MessageBox.Show("Error: " + Readerror, "Fehlerabfrage Spannungsmesser (ErrorDMM_U)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }

            catch (SystemException ex)
            {
                //Fehlerabfrage
                MessageBox.Show("Die Fehlerabfrage des Spannungsmessers ist fehlgeschlagen.\n" +
                                "Bitte überprüfen Sie die Kommunikation mit dem Messbus.\n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Read Error (ErrorDMM_U)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                InitDMM_U();
            }


            if (Manufacturer[0] == "HEWLETT-PACKARD")
            {
                try
                {
                    ioDMM_U.WriteString("ROUT:TERM?", true);
                    ReadTerm = ioDMM_U.ReadString();

                    do
                    {
                        ioDMM_U.WriteString("ROUT:TERM?", true);
                        ReadTerm = ioDMM_U.ReadString();
                        ReadTerm = ReadTerm.Substring(0, 4);

                        if (!(ReadTerm == "REAR"))
                        {
                            if (ReadTerm == "FRON") ReadTerm = "FRONT";


                            MessageBox.Show("Die Messbuchsen am Spannungsmesser stehen auf " + ReadTerm + ".\n" +
                                            "Bitte schalten Sie auf die rückwärtigen Messbuchsen am Spannungsmesser um!",
                                            "Fehler Abfrage Messbuchsen DMM_U", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }

                    } while (!(ReadTerm == "REAR"));
                }

                catch (SystemException ex)
                {
                    //Fehlerabfrage
                    MessageBox.Show("Die Abfrage der verwendeten Messbuchsen ist fehlgeschlagen.\n" +
                                    "Bitte überprüfen Sie die Kommunikation mit dem Messbus.\n" +
                                    "Source: " + ex.Source + "\n" +
                                    "Message: " + ex.Message + "\n",
                                    "Fehler Read Terminals DMM_U", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ErrorDMM_U();
                }
            }
        }

        private void ErrorDMM_I()
        {
            string Readerror = "";
            int PositionError = -1;
            string ReadTerm = "";

            try
            {
                while (PositionError == -1)
                {
                    // Abfrage ID und Abfrage ob Fehler vorhanden
                    ioDMM_I.WriteString("SYST:ERR?", true);
                    Readerror = ioDMM_I.ReadString();
                    PositionError = Readerror.IndexOf("No error");

                    if (PositionError == -1)
                    {
                        MessageBox.Show("Error: " + Readerror, "Fehlerabfrage Strommesser (ErrorDMM_I)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
            }

            catch (SystemException ex)
            {
                //Fehlerabfrage
                MessageBox.Show("Die Fehlerabfrage des Strommessers ist fehlgeschlagen. \n" +
                                "Bitte überprüfen Sie die Kommunikation mit dem Messbus.\n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Read Error (ErrorDMM_I)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                InitDMM_I();
            }

            if (Manufacturer[0] == "HEWLETT-PACKARD")
            {
                try
                {
                    ioDMM_I.WriteString("ROUT:TERM?", true);
                    ReadTerm = ioDMM_I.ReadString();

                    do
                    {
                        ioDMM_I.WriteString("ROUT:TERM?", true);
                        ReadTerm = ioDMM_I.ReadString();
                        ReadTerm = ReadTerm.Substring(0, 4);

                        if (!(ReadTerm == "REAR"))
                        {
                            if (ReadTerm == "FRON") ReadTerm = "FRONT";

                            MessageBox.Show("Die Messbuchsen am Strommesser stehen auf " + ReadTerm + ".\n" +
                                            "Bitte schalten Sie auf die rückwärtigen Messbuchsen am Stromsmesser um!",
                                            "Fehler Abfrage Messbuchsen DMM_I", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    } while (!(ReadTerm == "REAR"));
                }

                catch (SystemException ex)
                {
                    //Fehlerabfrage
                    MessageBox.Show("Die Abfrage der verwendeten Messbuchsen ist fehlgeschlagen.\n" +
                                    "Source: " + ex.Source + "\n" +
                                    "Message: " + ex.Message + "\n",
                                    "Fehler Read Terminals DMM_I", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    ErrorDMM_I();
                }
            }
        }

        private void ErrorSV_ACDC()
        {
            STB_SV_ACDC = "";
            STB = 0;
            AuswertungBit = false;
            BitNumber = 2;

            try
            {
                if (Manufacturer[2] == "SCHULTZ ELECTRONIC")
                {
                    do
                    {
                        // Abfrage STB- Register
                        ioSV_ACDC.WriteString("*STB?", true);
                        STB_SV_ACDC = ioSV_ACDC.ReadString();
                        STB_SV_ACDC = STB_SV_ACDC.Substring(8, 4);
                        STB = Convert.ToByte(STB_SV_ACDC);

                        switch (STB)
                        {
                            case 0:
                                break;
                            case 1:
                                MessageBox.Show("Syntax Error", "ErrorSV_ACDC", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                            case 2:
                                MessageBox.Show("Command Error", "ErrorSV_ACDC", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                            case 3:
                                MessageBox.Show("Range Error", "ErrorSV_ACDC", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                            case 4:
                                MessageBox.Show("Unit Error", "ErrorSV_ACDC", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                            case 5:
                                MessageBox.Show("Hardware Error", "ErrorSV_ACDC", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                            case 6:
                                MessageBox.Show("Read Error", "ErrorSV_ACDC", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                break;
                            default:
                                MessageBox.Show("STB: " + STB + " - " + "Fehler ist nicht in der Fehlerliste aufgeführt.\n" +
                                       "Bitte Stromversorgung überprüfen lassen.", "ErrorSV_ACDC", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                                break;
                        }
                    } while (!(STB == 0));
                }

                if (Manufacturer[2] == "HBS-ELECTRONIC")
                {
                    // Abfrage STB- Register
                    ioSV_ACDC.WriteString("*STB?", true);
                    STB_SV_ACDC = ioSV_ACDC.ReadString();
                    STB = Convert.ToByte(STB_SV_ACDC);

                    AuswertungBit = (STB & (1 << BitNumber)) > 0;

                    if (AuswertungBit)
                    {
                        MessageBox.Show("Query Error", "ErrorSV_ACDC", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

            }
            catch (SystemException ex)
            {
                //Fehlerabfrage
                MessageBox.Show("Die Abfrage des Statusregisters ist fehlgeschlagen. " + ex.Source + "  " + ex.Message, "Fehler Abfrage STB", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }

        //Reset Messequipment

        private void ResetDMM_U()
        {
            try
            {
                // Reset DMM_U (Spanungsmesser)
                ioDMM_U.WriteString("*RST", true);
            }
            catch (SystemException ex)
            {
                //Fehlerabfrage
                MessageBox.Show("Das Rücksetzen des Spannungsmessers ist fehlgeschlagen. \n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Reset Spannungsmesser (ResetDMM_U)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }

        private void ResetDMM_I()
        {
            try
            {
                // Reset DMM_I (Strommesser)
                ioDMM_I.WriteString("*RST", true);
            }
            catch (SystemException ex)
            {
                //Fehlerabfrage
                MessageBox.Show("Das Rücksetzen des Strommessers ist fehlgeschlagen. \n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Reset Strommesser (ResetDMM_I)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }

        private void ResetSV_ACDC()
        {
            try
            {
                if (Manufacturer[2] == "SCHULTZ ELECTRONIC")
                {
                    ioSV_ACDC.WriteString("RST", true);                    // Reset SV Schultz Electronic
                    ioSV_ACDC.WriteString("DCL", true);
                }

                if (Manufacturer[2] == "HBS-ELECTRONIC")
                {
                    ioSV_ACDC.WriteString("*RST", true);                      // Reset SV HBS Electronic
                }

            }
            catch (SystemException ex)
            {
                //Fehlerabfrage
                MessageBox.Show("Das Rücksetzen der Stromversorgung ist fehlgeschlagen. \n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Reset Stromversorgung (ResetSV)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ResetSV_ACDC();
            }
        }

        // Setzen der SV

        private void SetVoltageSV_ACDC(string setU)
        {
            setU = setU.Replace(",", ".");
            setUart = Convert.ToString(Uart[zPrüfschritt]);
            tbxsetU.Invoke(new Action(() => tbxsetU.Text = setU));

            switch (setUart)
            {
                case "AC":

                    try
                    {
                        if (Manufacturer[2] == "SCHULTZ ELECTRONIC")
                        {
                            ioSV_ACDC.WriteString("UAC," + setU, true);                                // Spannung setzen AC Schultz Electronic
                        }

                        if (Manufacturer[2] == "HBS-ELECTRONIC")
                        {
                            ioSV_ACDC.WriteString("SOUR:VOLTAC," + setU, true);                          // Spannung setzen AC HBS Electronic
                        }
                    }

                    catch (ArgumentException ex)
                    {
                        MessageBox.Show("Das Setzen der Spannung an der Spannungsversorgung ist fehlgeschlagen!" +
                                        "Source: " + ex.Source + "\n" +
                                        "Message: " + ex.Message + "\n",
                                        "Fehler Spannung setzen (SetVoltageSV)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                case "DC":

                    try
                    {
                        if (Manufacturer[2] == "SCHULTZ ELECTRONIC")
                        {
                            ioSV_ACDC.WriteString("UDC," + setU, true);                                // Spannung setzen DC Schultz Electronic
                        }

                        if (Manufacturer[2] == "HBS-ELECTRONIC")
                        {
                            ioSV_ACDC.WriteString("SOUR:VOLTDC," + setU, true);                           // Spannung setzen DC HBS Electronic
                        }
                    }

                    catch (ArgumentException ex)
                    {
                        MessageBox.Show("Das Setzen der Spannung an der Spannungsversorgung ist fehlgeschlagen!" +
                                        "Source: " + ex.Source + "\n" +
                                        "Message: " + ex.Message + "\n",
                                        "Fehler Spannung setzen (SetVoltageSV)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    break;

                default:
                    break;
            }

            ErrorSV_ACDC();
        }

        private void SetCurrentSV_ACDC(string setI)
        {
            setI = setI.Replace(",", ".");
            tbxsetI.Invoke(new Action(() => tbxsetI.Text = setI));

            try
            {
                if (Manufacturer[2] == "SCHULTZ ELECTRONIC")
                {
                    ioSV_ACDC.WriteString("IA," + setI, true);                                       // Strom setzen Schultz Electronic
                }

                if (Manufacturer[2] == "HBS-ELECTRONIC")
                {
                    ioSV_ACDC.WriteString("SOUR:CURR," + setI, true);                                // Strom setzen HBS Electronic
                }
            }

            catch (ArgumentException ex)
            {
                MessageBox.Show("Das Setzen des Stromes an der Spannungsversorgung ist fehlgeschlagen!" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Strom setzen (SetCurrentSV)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ErrorSV_ACDC();
        }

        private void SetCurrentMaxSV_ACDC()
        {
            if (Manufacturer[2] == "HBS-ELECTRONIC")
            {

                setImax = Convert.ToString(m_Imax * 1e-3 * 1.1);
                setImax = setImax.Replace(",", ".");
                tbxsetImax.Invoke(new Action(() => tbxsetImax.Text = setImax));

                try
                {
                    ioSV_ACDC.WriteString("SOUR:CURRMAX," + setImax, true);                               // Strom setzen HBS Electronic             
                }

                catch (ArgumentException ex)
                {
                    MessageBox.Show("Das Setzen des Stromes an der Spannungsversorgung ist fehlgeschlagen!" +
                                    "Source: " + ex.Source + "\n" +
                                    "Message: " + ex.Message + "\n",
                                    "Fehler Strom setzen (SetCurrentSV)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                ErrorSV_ACDC();
            }
        }

        private void SetFrequencySV_ACDC(double f)
        {
            if (Uart[zPrüfschritt] == "AC")
            {
                if (setf != Convert.ToString(f))
                {
                    setf = Convert.ToString(f);
                    setf = setf.Replace(",", ".");
                    tbxsetf.Invoke(new Action(() => tbxsetf.Text = setf));

                    try
                    {
                        if (Manufacturer[2] == "SCHULTZ ELECTRONIC")
                        {
                            ioSV_ACDC.WriteString("FA," + setf, true);                               // Frequenz setzen Schultz Electronic
                        }

                        if (Manufacturer[2] == "HBS-ELECTRONIC")
                        {
                            ioSV_ACDC.WriteString("SOUR:FREQ," + setf, true);                          // Frequenz setzen HBS Electronic
                        }
                    }

                    catch (SystemException ex)
                    {
                        MessageBox.Show("Das Setzen der Frequenz an der Spannungsversorgung ist fehlgeschlagen!" +
                                        "Source: " + ex.Source + "\n" +
                                        "Message: " + ex.Message + "\n",
                                        "Fehler Frequenz setzen (SetFrequencySV)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    ErrorSV_ACDC();
                }
            }
        }

        private void SetWaveSV_ACDC(string setW)
        {
            if (Manufacturer[2] == "SCHULTZ ELECTRONIC")
            {
                if (setW != Convert.ToString(Wave[zPrüfschritt]))
                {
                    setW = Convert.ToString(Wave[zPrüfschritt]);
                    tbxWave.Invoke(new Action(() => tbxWave.Text = setW));

                    try
                    {
                        ioSV_ACDC.WriteString("WAVE," + setW, true);
                    }

                    catch (SystemException ex)
                    {
                        MessageBox.Show("Das Setzen der Signalform an der Spannungsversorgung ist fehlgeschlagen!" +
                                        "Source: " + ex.Source + "\n" +
                                        "Message: " + ex.Message + "\n",
                                        "Fehler Signalform setzen (SetWaveSV)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    ErrorSV_ACDC();
                }
            }
        }

        private void SetOutputSV_ACDC(string setOut)
        {
            if (Manufacturer[2] == "HBS-ELECTRONIC" && setOut == "0")
            {
                setOut = "1";
            }
            else if (Manufacturer[2] == "HBS-ELECTRONIC" && setOut == "1")
            {
                setOut = "0";
            }

            tbxOut.Invoke(new Action(() => tbxOut.Text = setOut));

            try
            {
                if (Manufacturer[2] == "SCHULTZ ELECTRONIC")
                {
                    ioSV_ACDC.WriteString("SB," + setOut, true);                                     // Ausgang setzen Schultz Electronic
                }

                if (Manufacturer[2] == "HBS-ELECTRONIC")
                {
                    ioSV_ACDC.WriteString("OUTP:STAT " + setOut, true);                                 // Ausgang setzen HBS Electronic
                }
            }

            catch (SystemException ex)
            {
                MessageBox.Show("Das Freischalten des Ausganges der Stromversorgung ist fehlgeschlagen!" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Set Output (SetOutputSV)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            ErrorSV_ACDC();
        }

        private void MeasCurrSV_ACDC()
        {
            try
            {
                if (Manufacturer[2] == "SCHULTZ ELECTRONIC")
                {
                    ioSV_ACDC.WriteString("MIA," + setU, true);                                // Spannung setzen AC Schultz Electronic
                    ReadCurr = ioSV_ACDC.ReadString();
                    ReadCurr = ReadCurr.Substring(4, 5);
                    ReadCurr = ReadCurr.Replace(".", ",");
                    Imess = Convert.ToDouble(ReadCurr);
                    Imess = Math.Round(Imess, 3);               // Runden auf 3 Nachkommastellen 
                    tbxImess.Invoke(new Action(() => tbxImess.Text = Convert.ToString(Imess)));
                    tbxImess.Invoke(new Action(() => tbxImess.Text = Convert.ToString(Imess)));
                    tbxm_I.Invoke(new Action(() => tbxm_I.Text = Convert.ToString(m_I)));
                }
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("Das Setzen der Spannung an der Spannungsversorgung ist fehlgeschlagen!" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Spannung setzen (MEASCurrSV_ACDC)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }           
        }

        private void bgwMeasCurrSV_ACDC_DoWork(object sender, DoWorkEventArgs e)
        {
            MeasCurrSV_ACDC();
        }

        //Abfrage STB SV

        //private bool AbfrageACS_STB_SV_ACDC()
        //{
        //    string ACS_SV_ACDC = "";
        //    int ACS_STB = 0;            
        //    BitNumber = 3;

        //    try
        //    {
        //        if (Manufacturer[2] == "HBS-ELECTRONIC")
        //        {
        //            // Abfrage ACS-STB Byte
        //            ioSV_ACDC.WriteString("*ACS?", true);
        //            ACS_SV_ACDC = ioSV_ACDC.ReadString();
        //            ACS_STB = Convert.ToInt16(ACS_SV_ACDC);

        //            textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(ACS_STB)));
        //            textBox3.Invoke(new Action(() => textBox3.Text = Convert.ToString(ACS_SV_ACDC)));

        //            if ((ACS_STB & (1 << BitNumber)) > 0)
        //            {
        //                bitCC_Mode = true;
        //            }
        //            else
        //            {
        //                bitCC_Mode = false;
        //            }

        //            textBox1.Invoke(new Action(() => textBox1.Text = Convert.ToString(bitCC_Mode)));
        //        }

        //    }
        //    catch (SystemException ex)
        //    {
        //        //Fehlerabfrage                
        //        MessageBox.Show("Die Abfrage des Statusregisters ist fehlgeschlagen. " + ex.Source + "  " + ex.Message, "Fehler Abfrage ACS Status Byte", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        Close();
        //    }

        //    return bitCC_Mode;
        //}

        // Konfiguration Multimeter

        private void bgwSetConfigurationDMM_DoWork(object sender, DoWorkEventArgs e)
        {
            SetConfigurationDMM_U(m_Umax);
            SetConfigurationDMM_I(m_Imin, m_Imax);
        }

        private void SetConfigurationDMM_U(double m_Umax)
        {
            if (m_Umax <= 1000)
            {
                if (m_Umax <= 100)
                {
                    if (m_Umax <= 10)
                    {
                        if (m_Umax <= 1)
                        {
                            if (m_Umax <= 0.1)
                            {
                                if(m_Umax == 0)
                                {
                                    RangeDMM_U = "100";
                                    Res = "0.001";
                                }
                                else
                                {
                                RangeDMM_U = "0.1";
                                Res = "0.001";
                                }
                            }
                            else
                            {
                                RangeDMM_U = "1";
                                Res = "0.00001";
                            }
                        }
                        else
                        {
                            RangeDMM_U = "10";
                            Res = "0.0001";
                        }
                    }
                    else
                    {
                        RangeDMM_U = "100";
                        Res = "0.001";
                    }
                }
                else
                {
                    RangeDMM_U = "1000";
                    Res = "0.01";
                }
            }

            tbxRangeU.Invoke(new Action(() => tbxRangeU.Text = RangeDMM_U));
            tbxResU.Invoke(new Action(() => tbxResU.Text = Res));

            if ((!(bitSetConfigurationDMM_U)) || SetRangeDMM_U != RangeDMM_U)
            {
                SetRangeDMM_U = RangeDMM_U;

                try
                {
                    if (InformationMessequipmentDB[1, 0] == "HEWLETT-PACKARD")
                    {
                        ioDMM_U.WriteString("CONF:VOLT:" + Uart[zPrüfschritt] + " " + SetRangeDMM_U + "," + " " + Res);
                        ioDMM_U.WriteString("SENS:DET:BAND " + DETBAND);
                    }
                    else
                    {
                        ioDMM_U.WriteString("CONF:VOLT:" + Uart[zPrüfschritt] + " " + SetRangeDMM_U + "," + " " + Res);                       
                    }


                    bitSetConfigurationDMM_U = true;
                }
                catch (SystemException ex)
                {
                    MessageBox.Show("Die Konfiguration des Spannungsmessers ist fehlgeschlagen. \n" +
                                    "Bitte überprüfen Sie die Kommunikation mit dem Messbus.\n" +
                                    "Source: " + ex.Source + "\n" +
                                    "Message: " + ex.Message + "\n",
                                    "Fehler Konfiguration (SetConfigurationDMM_U)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            ErrorDMM_U();
        }

        private void SetConfigurationDMM_I(double m_Imin, double m_Imax)
        {
            if (m_Imax <= 1000 && m_Imax != 0)
            {
                RangeDMM_I = "1";
                Res = "0.00001";
            }
            else if (m_Imax <= 3000 && m_Imax != 0)
            {
                RangeDMM_I = "3";
                Res = "0.0001";
            }
            else if (m_Imin <= 1000 && m_Imin != 0 && m_Imax == 0)
            {
                RangeDMM_I = "1";
                Res = "0.00001";
            }
            else if (m_Imin <= 3000 && m_Imin != 0 && m_Imax == 0)
            {
                RangeDMM_I = "3";
                Res = "0.0001";
            }
            else
            {
                RangeDMM_I = "3";
                Res = "0.0001";
            }

            tbxRangeI.Invoke(new Action(() => tbxRangeI.Text = SetRangeDMM_I));
            tbxResI.Invoke(new Action(() => tbxResI.Text = Res));

            if ((!(bitSetConfigurationDMM_I)) || SetRangeDMM_I != RangeDMM_I)
            {
                SetRangeDMM_I = RangeDMM_I;
                try
                {


                    if (InformationMessequipmentDB[1, 0] == "HEWLETT-PACKARD")
                    {
                        ioDMM_I.WriteString("CONF:CURR:" + Uart[zPrüfschritt] + " " + RangeDMM_I + "," + " " + Res);
                        ioDMM_I.WriteString("SENS:DET:BAND " + DETBAND);
                    }
                    else
                    {
                        ioDMM_I.WriteString("CONF:CURR:" + Uart[zPrüfschritt]);
                        ioDMM_I.WriteString("CURR:" + Uart[zPrüfschritt] + ":RANG " + RangeDMM_I);
                    }

                    bitSetConfigurationDMM_I = true;
                }
                catch (SystemException ex)
                {
                    MessageBox.Show("Die Konfiguration des Strommessers ist fehlgeschlagen. \n" +
                                    "Bitte überprüfen Sie die Kommunikation mit dem Messbus.\n" +
                                    "Source: " + ex.Source + "\n" +
                                    "Message: " + ex.Message + "\n",
                                    "Fehler Konfiguration (SetConfigurationDMM_I)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            ErrorDMM_I();
        }

        // messen von Spannung, Strom, Temp.

        private double MeasureVoltage()
        {
            try
            {
                ioDMM_U.WriteString("INIT");
                ioDMM_U.WriteString("FETCH?");
                ReadVolt = ioDMM_U.ReadString();
                ReadVolt = ReadVolt.Replace(".", ",");
                Umess = Convert.ToDouble(ReadVolt);
                Umess = Math.Round(Umess, 3);               // Runden auf 3 Nachkommastellen 
                tbxUmess.Invoke(new Action(() => tbxUmess.Text = Convert.ToString(Umess)));

                if (((m_Umin > Umess || Umess > m_Umax) && m_PrfgU) || (m_PrfgAb && (U1[zPrüfschritt] > Umess || Umess > U2[zPrüfschritt])))
                    tbxUmess.Invoke(new Action(() => tbxUmess.BackColor = tbxUmess.BackColor = Color.LightPink));
                else if ((m_Umin <= Umess && Umess <= m_Umax && m_PrfgU) || (m_PrfgAb && (U1[zPrüfschritt] < Umess && Umess < U2[zPrüfschritt])))
                    tbxUmess.Invoke(new Action(() => tbxUmess.BackColor = tbxUmess.BackColor = Color.LightGreen));
                else if (m_Umin <= Umess && Umess <= m_Umax && !m_PrfgU)
                    tbxUmess.Invoke(new Action(() => tbxUmess.BackColor = tbxUmess.BackColor = SystemColors.Window));
            }
            catch (SystemException ex)
            {
                MessageBox.Show("Die Spannungsmessung ist fehlgeschlagen. \n" +
                                "Bitte überprüfen Sie die Kommunikation mit dem Messbus. \n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Spannungsmessung (MeasureVoltage)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            ErrorDMM_U();

            return Umess;
        }

        private double MeasureCurrent()
        {
            try
            {
                ioDMM_I.WriteString("INIT");
                ioDMM_I.WriteString("FETCH?");
                ReadCurr = ioDMM_I.ReadString();
                ReadCurr = ReadCurr.Replace(".", ",");
                Imess = Convert.ToDouble(ReadCurr);
                Imess = Imess * 1E3;                        // Umrechnung in mA   
                Imess = Math.Round(Imess, 3);               // Runden auf 3 Nachkommastellen 
                tbxImess.Invoke(new Action(() => tbxImess.Text = Convert.ToString(Imess)));

                if (((m_Imin > Imess || Imess > m_Imax && m_Imin != 0 && m_Imax != 0 || 
                     Imess < m_Imin && m_Imax == 0 || 
                     Imess > m_Imax && m_Imin == 0) && m_PrfgI) ||
                     Imess > m_Imax && m_PrfgAb)
                    tbxImess.Invoke(new Action(() => tbxImess.BackColor = tbxImess.BackColor = Color.LightPink));
                else if (((m_Imin <= Imess && Imess <= m_Imax && m_Imin != 0 && m_Imax != 0 || 
                           Imess >= m_Imin && m_Imax == 0 || 
                           Imess <= m_Imax && m_Imin == 0) && m_PrfgI) ||
                           Imess <= m_Imax && m_PrfgAb)
                    tbxImess.Invoke(new Action(() => tbxImess.BackColor = tbxImess.BackColor = Color.LightGreen));
                else if (m_Imin < Imess && Imess < m_Imax && m_Imin != 0 && m_Imax != 0 && !m_PrfgI)
                    tbxImess.Invoke(new Action(() => tbxImess.BackColor = tbxImess.BackColor = SystemColors.Window));
            }
            catch (SystemException ex)
            {
                MessageBox.Show("Die Strommessung ist fehlgeschlagen. \n" +
                                "Bitte überprüfen Sie die Kommunikation mit dem Messbus. \n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Strommessung (MeasureCurrent)", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            ErrorDMM_I();

            return Imess;
        }

        private void bgwMeasureVoltage_DoWork(object sender, DoWorkEventArgs e)
        {
            MeasureVoltage();
        }

        private void bgwMeasureCurrent_DoWork(object sender, DoWorkEventArgs e)
        {
            MeasureCurrent();
        }

        private void MeasureTemperature()
        {
            try
            {
                Task analogInTask = new Task();

                AIChannel TempAIChannel;

                TempAIChannel = analogInTask.AIChannels.CreateThermocoupleChannel("Dev1/ai0", "Temperature", minMessTemp, maxMessTemp, AIThermocoupleType.J, AITemperatureUnits.DegreesC);

                AnalogSingleChannelReader reader = new AnalogSingleChannelReader(analogInTask.Stream);

                PruefTemp = reader.ReadSingleSample();

                if (PruefTemp <= maxMessTemp)
                {
                    if (lowerTemp <= PruefTemp && PruefTemp <= upperTemp)
                    {
                        lblTempData.Text = "Temp: " + PruefTemp.ToString("0.00") + "°C";

                        if (PruefTemp <= lowerTemp + 1)
                        {
                            lblTempData.BackColor = Color.Yellow;
                        }

                        if (PruefTemp >= upperTemp - 1)
                        {
                            lblTempData.BackColor = Color.Yellow;
                        }

                        if (PruefTemp > lowerTemp + 1 && PruefTemp < upperTemp - 1)
                        {
                            lblTempData.BackColor = SystemColors.Control;
                        }

                    }
                    else
                    {
                        lblTempData.Text = "Temp: " + PruefTemp.ToString("0.00") + "°C";
                        lblTempData.BackColor = Color.Red;

                        MessageBox.Show("Die gemessene Temperatur von " + PruefTemp.ToString("0.00") +
                                        " °C befindet sich außerhalb der Prüftemperatur von " + lowerTemp + " bis " + upperTemp + " °C!!!\n" +
                                        "Die Prüfung kann nicht fortgesetzt werden!!!",
                                        "Fehler Prüftemperatur (MeasureTemperature)", MessageBoxButtons.OK, MessageBoxIcon.Error);

                        bitAbbruch = true;
                    }
                }
                else
                {
                    MessageBox.Show("Es befindet sich kein Temperatursensor am Thermometer!!!! \n" +
                                    "      Bitte Temperatursensor am Thermometer stecken!!!",
                                    "Fehler Temperatursensor (MeasureTemperature)", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    bitAbbruch = true;
                }
            }
            catch (SystemException ex)
            {
                MessageBox.Show("Es befindet sich kein Thermometer am Messbus!\n" +
                                "Bitte überprüfen Sie die Datenverbindung zwischen Thermometer und Rechner!\n" +
                                "Source: " + ex.Source + "\n" +
                                "Message: " + ex.Message + "\n",
                                "Fehler Thermometer (MeasureTemperature)", MessageBoxButtons.OK, MessageBoxIcon.Error);

                bitAbbruch = true;
            }
        }

        // ********************************************************* Objekte Kontextmenü *********************************************************************

        public void tsmiDateiAnmelden_Click(object sender, EventArgs e)
        {
            F_Login Authentifizierung = new F_Login();
            Authentifizierung.ShowDialog();
            SetObjAfterLogin(getanmeldung);
        }

        private void tsmiDateiAbmelden_Click(object sender, EventArgs e)
        {
            SetObjAfterMeasurement();
            SetObjAfterLogout();
            ResetVar();
            SetOutputSV_ACDC("1");
            ResetSV_ACDC();
            ResetDMM_U();
            ResetDMM_I();
            ResetOutDigIO();
        }

        private void tsmiDateiChangePW_Click(object sender, EventArgs e)
        {

            F_PwdÄndern ChangePW = new F_PwdÄndern();
            ChangePW.ShowDialog();

        }

        private void tsmiDateiBeenden_Click(object sender, EventArgs e)
        {

            Close();

        }

        private void tsmiExtrasSystemInit_Click(object sender, EventArgs e)
        {
            LoadID_MessequipmentDB();
            LoadInformationMessequipmentDB();
            InitMessequipment();
        }

        private void tsmiExtrasEinstellungen_Click(object sender, EventArgs e)
        {
            F_Einstellungen Einstellungen = new F_Einstellungen();
            Einstellungen.ShowDialog();
            SetConfigSerPort();
        }

        private void tsmiToolsConnectionExpert_Click(object sender, EventArgs e)
        {

            Process AgilentConnectionExpert = new Process();
            AgilentConnectionExpert.StartInfo.FileName = "AgilentConnectionExpert";
            AgilentConnectionExpert.Start();

        }

        private void tsmiToolsIOmonitor_Click(object sender, EventArgs e)
        {

            Process IOMonitor = new Process();
            IOMonitor.StartInfo.FileName = "IOMonitor";
            IOMonitor.Start();

        }

        private void tsmiToolsInteractiveIO_Click(object sender, EventArgs e)
        {

            Process InteractiveIO = new Process();
            InteractiveIO.StartInfo.FileName = "InteractiveIO";
            InteractiveIO.Start();

        }

        private void tsmiToolsIOMonitorViewer_Click(object sender, EventArgs e)
        {

            Process IOMonitorViewer = new Process();
            IOMonitorViewer.StartInfo.FileName = "IOMonitorViewer";
            IOMonitorViewer.Start();

        }

        private void tsmiToolsNIOpenMA_Click(object sender, EventArgs e)
        {

            Process NIMax = new Process();
            NIMax.StartInfo.FileName = "Measurement & Automation";
            NIMax.Start();

        }

        private void tsmiInfo_Click(object sender, EventArgs e)
        {

            new AboutBox1().Show();

        }

        private void tsmiDokuAgilentIOLSQSQ_Click(object sender, EventArgs e)
        {

            string path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "IO Libraries Quick Start Guide.pdf");

            if (!(File.Exists(path)))
            {

                File.WriteAllBytes(path, Resources.Resource1.IO_Libraries_Quick_Start_Guide);

            }
            Process.Start(path);


        }

        private void tsmiDokuHWSEAS_Click(object sender, EventArgs e)
        {

            string path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Bedienungsanleitung Wechselspannungsquelle SE-AS 15V 3.35A.pdf");

            if (!(File.Exists(path)))
            {

                File.WriteAllBytes(path, Resources.Resource1.Bedienungsanleitung_Wechselspannungsquelle_SE_AS_15V_3_35A);

            }

            Process.Start(path);

        }

        private void tsmiDokuHWhp34401A_Click(object sender, EventArgs e)
        {


            string path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Agilent 34401A Multimeter User Guide.pdf");

            if (!(File.Exists(path)))
            {

                File.WriteAllBytes(path, Resources.Resource1.Agilent_34401A_Multimeter_User_Guide);

            }

            Process.Start(path);

        }

        private void tsmiHilfe_Click(object sender, EventArgs e)
        {

            this.helpProvider1 = new System.Windows.Forms.HelpProvider();


        }

        private void contMStsmiAnmelden_Click(object sender, EventArgs e)
        {
            F_Login Authentifizierung = new F_Login();
            Authentifizierung.ShowDialog();
            SetObjAfterLogin(getanmeldung);
        }

        private void contMStsmiBeenden_Click(object sender, EventArgs e)
        {

            Close();

        }

        // ***************************************************** Objekte Form F_LED_Signalgeber **************************************************************

        private void cbxSachnummer_SelectedIndexChanged(object sender, EventArgs e)
        {
            Sachnummer = coboxSachnummer.Text.ToString();
            tbxSeriennummer.Enabled = true;
            tbxSeriennummer.Clear();
            tbxSeriennummer.Focus();
        }

        private void tbxSeriennummer_Enter(object sender, EventArgs e)
        {
            if (LastSachnummer != Sachnummer)
            {
                LoadIDParametersatzDB();
                LoadParametersatzDB();
                LoadPrüfschritteDB();
            }

            WriteParameterinVariables(1);
            CalcTol(1);
            LoadParameterInObjects();
            LoadSachnummerDB();
            coboxSachnummer.Text = Sachnummer;
            bgwSetConfigurationDMM.RunWorkerAsync();            
        }

        private void tbxSeriennummer_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (!((e.KeyValue >= 48 && e.KeyValue <= 57) || //Prüfung auf Tasten 0-9
             (e.KeyValue >= 96 && e.KeyValue <= 105) || //Prüfung auf NumPad 0-9
             (e.KeyValue == 189 || //Prüfung auf Taste -
              e.KeyValue == 109) || //Prüfung auf NumPad - 
             (e.KeyCode == Keys.Back || //Rücktaste
              e.KeyCode == Keys.Delete))) //Entf-Taste
            {
                e.SuppressKeyPress = true; //KeyPress unterdrücken
                e.Handled = true;
            }

            if (e.KeyCode == Keys.Enter)
            {
                Serialnummer = tbxSeriennummer.Text;


                if (Serialnummer != "" && Serialnummer.Length == 14)
                {
                    Auftragsnummer = Serialnummer.Substring(0, 10);
                    tbxAuftrag.Text = Auftragsnummer;
                    coboxSachnummer.Enabled = false;
                    tbxSeriennummer.Enabled = false;
                    btnStartWeiterSpeichern.Enabled = true;
                    btnStartWeiterSpeichern.Focus();    
                }
                else
                {
                    MessageBox.Show("Serialnummer entspricht nicht der Länge von 14 Zeichen!!!", "Fehler Serialnummer (tbxSeriennummer_KeyDown)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void UpDown_manuell(double m_SWeite)
        {
            bitupdown_manuell = true;

            SetObjDuringMeasurement();

            switch (m_Uart)
            {
                case "AC":

                    switch (m_Rp)
                    {
                        case "U":


                            bgwUAC_manuell.RunWorkerAsync();

                            ReadPosActivForm();

                            DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                    Environment.NewLine + Environment.NewLine +
                                                   "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                   "Bedienerhinweis",
                                                    MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.Cancel == DRM)
                                bitAbbruch = true;
                            break;

                        case "I":

                            bgwIAC_manuell.RunWorkerAsync();

                            ReadPosActivForm();

                            DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                    Environment.NewLine + Environment.NewLine +
                                                   "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                   "Bedienerhinweis",
                                  MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.Cancel == DRM)
                                bitAbbruch = true;

                            break;

                        default:

                            MboxText = Environment.NewLine +
                                       "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                                       Environment.NewLine +
                                       "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                            MboxTitle = "Fehler Auswahl Prüfung";

                            DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                                    posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.OK == DRM)
                            {
                                bitAbbruch = true;
                                AbortTest();
                            }
                            break;
                    }
                    break;

                case "DC":

                    switch (m_Rp)
                    {
                        case "U":

                            bgwUDC_manuell.RunWorkerAsync();

                            ReadPosActivForm();

                            DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                    Environment.NewLine + Environment.NewLine +
                                                   "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                   "Bedienerhinweis",
                                                    MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.Cancel == DRM)
                                bitAbbruch = true;
                            break;//case DC; case U


                        case "I":

                            bgwUDC_manuell.RunWorkerAsync();

                            ReadPosActivForm();

                            DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                    Environment.NewLine + Environment.NewLine +
                                                   "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                   "Bedienerhinweis",
                                                    MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.Cancel == DRM)
                                bitAbbruch = true;

                            break;

                        default:

                            MboxText = Environment.NewLine +
                                       "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                                       Environment.NewLine +
                                       "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                            MboxTitle = "Fehler Auswahl Prüfung";

                            DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                                    posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.OK == DRM)
                            {
                                bitAbbruch = true;
                                AbortTest();
                            }
                            break;

                    }
                    break;

                default:

                    MboxText = Environment.NewLine +
                               "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                               Environment.NewLine +
                               "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                    MboxTitle = "Fehler Auswahl Prüfung";

                    DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                            posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.OK == DRM)
                    {
                        bitAbbruch = true;
                        AbortTest();
                    }
                    break;
            }
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (m_Rp == "U" || m_Rp == "I")
            {
                OldValue = tbarUIMinMax.Value;
                tbxSWtBarOld.Text = Convert.ToString(OldValue);
                NewValue = OldValue - tbarUIMinMax.TickFrequency;                             
                tbarUIMinMax.Value = NewValue;
                tbxSWtBarNew.Text = Convert.ToString(NewValue);

                m_SWeite = (OldValue - NewValue) * m_ResSV;               

                UpDown_manuell(m_SWeite);

                OldValue = NewValue;
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                           Environment.NewLine +
                           "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                MboxTitle = "Fehler Auswahl Prüfung";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DialogResult.OK == DRM)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if (m_Rp == "U" || m_Rp == "I")
            {
                OldValue = tbarUIMinMax.Value;
                tbxSWtBarOld.Text = Convert.ToString(OldValue);
                NewValue = OldValue + tbarUIMinMax.TickFrequency;
                tbarUIMinMax.Value = NewValue;
                tbxSWtBarNew.Text = Convert.ToString(NewValue);

                m_SWeite = (OldValue - NewValue) * m_ResSV;

                UpDown_manuell(m_SWeite);

                OldValue = NewValue;
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                           Environment.NewLine +
                           "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                MboxTitle = "Fehler Auswahl Prüfung";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DialogResult.OK == DRM)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
        }

        private void btnLaser_Click(object sender, EventArgs e)
        {
            SetLaser();
        }

        private void btnRepeat_Click(object sender, EventArgs e)
        {            
            ReadPosActivForm();

            F_Repeat F_Repeat = new F_Repeat(zPrüfschritt, U1, I1min, I1max, f, VP, Lichtstärke, Farbort ,posActFormx, posActFormy, ActFormWidth, ActFormHeight);
            DRR = F_Repeat.ShowDialog();
           
            RepeatPrüfschritt(DRR);
        }

        private void btnStartWeiterSpeichern_Click(object sender, EventArgs e)
        {
            StartWeiterSpeichernPrüfablauf();
        }

        private void btnAbbruch_Click(object sender, EventArgs e)
        {
            ReadPosActivForm();

            DRM = F_MessageBox.Show("Möchten Sie die Prüfung abbrechen?.", "Bedienerhinweis Abbruch", MessageBoxButtons.YesNo, new F_MessageBox(),
                                    posActFormx, posActFormy, ActFormWidth, ActFormHeight);

            if (DRM == DialogResult.Yes)
            {
                bitAbbruch = true;
                AbortTest();
            }
        }

        private void cboxiO_CheckedChanged(object sender, EventArgs e)
        {
            AppraisalVisualTest();
        }

        private void cboxniO_CheckedChanged(object sender, EventArgs e)
        {
            AppraisalVisualTest();
        }

        private void lblEinheitminStrom_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            ActFormWidth = F_Abschaltung_Signalgeber.ActiveForm.Bounds.Width;
            ActFormHeight = F_Abschaltung_Signalgeber.ActiveForm.Bounds.Height;
            posActFormx = this.Location.X;
            posActFormy = this.Location.Y;

            if (Control.ModifierKeys == Keys.Control && (ActFormWidth != minsizewidth || ActFormHeight != minsizeheight))
            {
                this.Size = new Size(minsizewidth, minsizeheight);
                this.Location = new Point(((ActFormWidth - maxsizewidth) / 2) + posActFormx, ((ActFormHeight - maxsizeheight) / 2) + posActFormy);

            }
        }

        private void lblEinheitmaxStrom_MouseDoubleClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            ActFormWidth = F_Abschaltung_Signalgeber.ActiveForm.Bounds.Width;
            ActFormHeight = F_Abschaltung_Signalgeber.ActiveForm.Bounds.Height;
            posActFormx = this.Location.X;
            posActFormy = this.Location.Y;

            if (Control.ModifierKeys == Keys.Control && (ActFormWidth != maxsizewidth || ActFormHeight != maxsizeheight))
            {
                this.Size = new Size(maxsizewidth, maxsizeheight);
                this.Location = new Point(((ActFormWidth - minsizewidth) / 2) + posActFormx, ((ActFormHeight - minsizeheight) / 2) + posActFormy);
            }
        }

        private void TasteSperren_Tick(object sender, EventArgs e)
        {
            bittimerTaste = false;
            timerTaste.Stop();
        }

        private void tbarUIMinMax_KeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {

            if (e.KeyCode.Equals(Keys.Left) ||
                e.KeyCode.Equals(Keys.Right) ||
                e.KeyCode.Equals(Keys.Up) ||
                e.KeyCode.Equals(Keys.Down) ||
                e.KeyCode.Equals(Keys.PageUp) ||
                e.KeyCode.Equals(Keys.PageDown))
            {
                left = Keyboard.IsKeyDown(Key.Left);
                right = Keyboard.IsKeyDown(Key.Right);
                up = Keyboard.IsKeyDown(Key.Up);
                down = Keyboard.IsKeyDown(Key.Down);
                pageup = Keyboard.IsKeyDown(Key.PageUp);
                pagedown = Keyboard.IsKeyDown(Key.PageDown);

                if (left && (right || up || down || pagedown || pageup || mousebutton) ||
                    right && (left || up || down || pagedown || pageup || mousebutton) ||
                    down && (left || right || up || pagedown || pageup || mousebutton))
                {
                    MessageBox.Show("Bitte nur eine Taste drücken", "Fehler Tastendruck Keyboard");
                    tbarUIMinMax.Value = OldValue;
                    keydown = false;
                    mousebutton = false;
                }
                else
                {

                        keydown = true;
                        OldValue = tbarUIMinMax.Value;
                        tbxSWtBarOld.Text = Convert.ToString(OldValue);
                }
            }   
        }

        private void tbarUIMinMax_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (e.KeyCode.Equals(Keys.Left) ||
                e.KeyCode.Equals(Keys.Right) ||
                e.KeyCode.Equals(Keys.Up) ||
                e.KeyCode.Equals(Keys.Down)||
                e.KeyCode.Equals(Keys.PageDown)||
                e.KeyCode.Equals(Keys.PageUp))
            {
                left = Keyboard.IsKeyUp(Key.Left);
                right = Keyboard.IsKeyUp(Key.Right);
                up = Keyboard.IsKeyUp(Key.Up);
                down = Keyboard.IsKeyUp(Key.Down);
                pagedown = Keyboard.IsKeyUp(Key.PageDown);
                pageup = Keyboard.IsKeyUp(Key.PageUp);

                if (left && right && up && down && pagedown && pageup)
                {
                    keydown = false;    
                    NewValue = tbarUIMinMax.Value;
                    m_SWeite = (OldValue - NewValue) * m_ResSV;

                    tbxSWtBarNew.Text = Convert.ToString(NewValue);
                    UpDown_manuell(m_SWeite); 
                    OldValue = NewValue;
                }
            }            
        }

        private void tbarUIMinMax_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (keydown)
            {
                MessageBox.Show("Bitte nur eine Taste drücken", "Fehler Tastendruck Maus");
                mousebutton = false;
                keydown = false;
                tbarUIMinMax.Value = OldValue;
            }
            else
            {
                mousebutton = true;                
                tbxSWtBarOld.Text = Convert.ToString(OldValue);
            }
        }

        private void tbarUIMinMax_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            mousebutton = false;
            NewValue = tbarUIMinMax.Value;
            tbxSWtBarNew.Text = Convert.ToString(NewValue);

            m_SWeite = (OldValue - NewValue) * m_ResSV;
         
            UpDown_manuell(m_SWeite);

            OldValue = NewValue;
        }

        // ******************************************Abfrage Status der Anmeldung und Abfrage Prüfer_ID*******************************************************

        public static void GetStatusAnmeldung(bool anmeldung)
        {
            getanmeldung = anmeldung;
        }

        public static void GetPrüferIDAnmeldung(string login_ID_Prüfer)
        {
            ID_Prüfer = login_ID_Prüfer;
        }

        // ************************************************* Setzen der Objekte Form F_LED_Signalgeber *******************************************************

        private void SetObjAfterLogin(bool getanmeldung)
        {
            if (getanmeldung == true)
            {
                ttptbxSachnummer.SetToolTip(coboxSachnummer, "Sachnummer des Prüflings auswählen");
                ttptbxSeriennummer.SetToolTip(tbxSeriennummer, "Seriennummer des Prüflings eingeben oder Einscannen");
                ttptbxSollspannung.SetToolTip(tbxUmin, "Sollspannung aus Datenbank");
                ttptbxSollstrom.SetToolTip(tbxImin, "Sollstrom aus Datenbank");
                ttptbxMeasspannung.SetToolTip(tbxUmess, "Gemessene Spannung am Prüfling");
                ttptbxMeasstrom.SetToolTip(tbxImess, "Gemessener Strom am Prüfling");
                ttpbtnUp.SetToolTip(btnUp, "Korrektur des Sollwertes Spannung oder Strom in pos. Richtung");
                ttpbtnDown.SetToolTip(btnDown, "Korrektur des Sollwertes Spannung oder Strom in neg. Richtung");
                ttptbxVisuellePrüfung.SetToolTip(tbxVisuellePrüfung, "Beschreibung Visuelle Prüfung");
                ttpcboxiO.SetToolTip(cboxiO, "Bewertung Prüfschritt iO");
                ttpcboxniO.SetToolTip(cboxniO, "Bewertung Prüfschritt niO");
                ttpbtnAbbruch.SetToolTip(btnCancel, "Abbruch der Prüfung");
                ttpbtnStartWeiterSpeichern.SetToolTip(btnStartWeiterSpeichern, "Start des Prüfablaufs");

                tsmiDateiAnmelden.Enabled = false;
                tsmiDateiAbmelden.Enabled = true;
                optionenToolStripMenuItem.Enabled = true;
                tsmiTools.Enabled = true;
                tsmiHilfe.Enabled = true;
                coboxSachnummer.Enabled = true;
                tbxUmess.BackColor = SystemColors.Window;
                tbxImess.BackColor = SystemColors.Window;
                tbxVisuellePrüfung.BackColor = SystemColors.Window;
                tbxLichtstärke.BackColor = SystemColors.Window;
                tbxFarbort.BackColor = SystemColors.Window;
                coboxSachnummer.ResetText();
                tbxSeriennummer.Clear();
                tbxUmin.Clear();
                tbxImin.Clear();
                tbxImax.Clear();
                tbxVisuellePrüfung.Clear();
                cboxiO.Checked = false;
                cboxniO.Checked = false;
                btnLaser.Enabled = true;
          
                LoadSachnummerDB();
            }
            else
            {
                MessageBox.Show("Anmeldung fehlgeschlagen", "SetOjectsnachAnmelden", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetObjAfterLogout()
        {
            tsmiDateiAnmelden.Enabled = true;
            tsmiDateiAbmelden.Enabled = false;
            optionenToolStripMenuItem.Enabled = true;
            tsmiTools.Enabled = false;
            tsmiHilfe.Enabled = false;
            coboxSachnummer.Enabled = false;
            tbxSeriennummer.Enabled = false;
            btnUp.Enabled = false;
            btnDown.Enabled = false;
            tbarUIMinMax.Enabled = false;
            tbxUmess.BackColor = SystemColors.Control;
            tbxImess.BackColor = SystemColors.Control;
            tbxVisuellePrüfung.BackColor = SystemColors.Control;
            tbxLichtstärke.BackColor = SystemColors.Control;
            tbxFarbort.BackColor = SystemColors.Control;
            cboxiO.Enabled = false;
            cboxniO.Enabled = false;
            cboxiO.Checked = false;
            cboxniO.Checked = false;
            btnLaser.Enabled = false;
            btnRepeat.Enabled = false;
            btnStartWeiterSpeichern.Enabled = false;
            btnCancel.Enabled = false;
            coboxSachnummer.ResetText();
            tbxSeriennummer.Clear();
            tbxUmin.Clear();
            tbxImin.Clear();
            tbxImax.Clear();
            tbxVisuellePrüfung.Clear();
            tbxLichtstärke.Clear();
            tbxFarbort.Clear();
        }
        
        private void SetObjDuringMeasurement()
        {
            bitStatusPrüfung = true;

            btnStartWeiterSpeichern.Invoke(new Action(() => btnStartWeiterSpeichern.Enabled = btnStartWeiterSpeichern.Enabled = false));
            coboxSachnummer.Invoke(new Action(() => coboxSachnummer.Enabled = coboxSachnummer.Enabled = false));
            btnCancel.Invoke(new Action(() => btnCancel.Enabled = btnCancel.Enabled = true));
            btnUp.Invoke(new Action(() => btnUp.Enabled = btnUp.Enabled = false));
            btnDown.Invoke(new Action(() => btnDown.Enabled = btnDown.Enabled = false));
            btnRepeat.Invoke(new Action(() => btnRepeat.Enabled = btnRepeat.Enabled = false));
            tbarUIMinMax.Invoke(new Action(() => tbarUIMinMax.Enabled = tbarUIMinMax.Enabled = false));

            if (cboxiO.Checked == true)
            {
                cboxiO.Invoke(new Action(() => cboxiO.Checked = cboxiO.Checked = false));
                cboxiO.Invoke(new Action(() => cboxiO.Enabled = cboxiO.Enabled = false));
                cboxniO.Invoke(new Action(() => cboxniO.Enabled = cboxniO.Enabled = false));
            }
            else
            {
                cboxniO.Invoke(new Action(() => cboxniO.Checked = cboxniO.Checked = false));
                cboxniO.Invoke(new Action(() => cboxniO.Enabled = cboxniO.Enabled = false));
                cboxiO.Invoke(new Action(() => cboxiO.Enabled = cboxiO.Enabled = false));
            }
        }

        private void SetObjAfterMeasurement()
        {
            do
            { }
            while (bgwMeasureVoltage.IsBusy || bgwMeasureCurrent.IsBusy);

            if (!(bitRegelung))
            {
                switch (btnStartWeiterSpeichern.Text)
                {
                    case "&Weiter":
                        this.Invoke(new Action(() => ttpbtnStartWeiterSpeichern.SetToolTip(btnStartWeiterSpeichern, "Weiter zum nächsten Prüfschritt")));
                        break;

                    case "&Speichern":
                        this.Invoke(new Action(() => ttpbtnStartWeiterSpeichern.SetToolTip(btnStartWeiterSpeichern, "Prüfergebnisse in DB speichern")));
                        break;

                    default:
                        break;
                }

                this.Invoke(new Action(() => SendKeys.Send("%()")));                //ALT-Taste drücken, um die Keycodes anzuzeigen
                this.Invoke(new Action(() => SendKeys.Send("%()")));                //ALT-Taste drücken, um Focus von Menüleiste zu entfernen  

                btnStartWeiterSpeichern.Invoke(new Action(() => btnStartWeiterSpeichern.Select()));

                if (bitRegelung == false)
                {
                    if (m_PrfgVi)
                    {
                        cboxiO.Invoke(new Action(() => cboxiO.Enabled = cboxiO.Enabled = true));
                        cboxniO.Invoke(new Action(() => cboxniO.Enabled = cboxniO.Enabled = true));
                    }
                    else
                    {
                        btnStartWeiterSpeichern.Invoke(new Action(() => btnStartWeiterSpeichern.Enabled = btnStartWeiterSpeichern.Enabled = true));
                    }

                    if (((Umess < m_Umax + 0.01 && m_Rp == "U") || (Imess < m_Imax && m_Rp == "I") && (tbarUIMinMax.Value < tbarUIMinMax.Maximum)))
                    {
                        btnUp.Invoke(new Action(() => btnUp.Enabled = btnUp.Enabled = true));
                    }

                    if (((Umess > m_Umin - 0.01 && m_Rp == "U") || (Imess > m_Imin && m_Rp == "I") && (tbarUIMinMax.Value > tbarUIMinMax.Minimum)))
                    {
                        btnDown.Invoke(new Action(() => btnDown.Enabled = btnDown.Enabled = true));
                    }

                    btnRepeat.Invoke(new Action(() => btnRepeat.Enabled = btnRepeat.Enabled = true));
                    tbarUIMinMax.Invoke(new Action(() => tbarUIMinMax.Enabled = tbarUIMinMax.Enabled = true));
                }
                else if (aPrüfschritte >= zPrüfschritt && bitRegelung == false) // && PrfgAbschaltung[zPrüfschritt] == true)
                {
                    if ((m_Umin <= Umess && Umess <= m_Umax && m_Rp == "U" && ((m_Imin <= Imess && Imess <= m_Imax) || (m_Imin == 0 && m_Imax == 0)) ||
                      ((m_Umin <= Umess && Umess <= m_Umax) && (m_Imin <= Imess && Imess <= m_Imax) && m_Rp == "I")) && bitAbschaltung == true)
                    {
                        cboxiO.Invoke(new Action(() => cboxiO.Enabled = cboxiO.Enabled = true));
                        cboxniO.Invoke(new Action(() => cboxniO.Enabled = cboxniO.Enabled = false));
                    }
                    else
                    {
                        cboxiO.Invoke(new Action(() => cboxiO.Enabled = cboxiO.Enabled = false));
                        cboxniO.Invoke(new Action(() => cboxniO.Enabled = cboxniO.Enabled = true));
                    }
                }                                
            }
        }

        private void SetObjForNextPschritt()
        {
            btnStartWeiterSpeichern.Enabled = false; tbxImess.Clear();

            tbxUmess.Clear();
            tbxImess.Clear();
            tbxUmess.BackColor = SystemColors.Window;
            tbxImess.BackColor = SystemColors.Window;
        }

        private void ResetObj()
        {
            cboxiO.Checked = false;
            cboxniO.Checked = false;
            cboxiO.Enabled = false;
            cboxniO.Enabled = false;
            cboxiO.BackColor = SystemColors.Control;
            cboxniO.BackColor = SystemColors.Control;

            btnUp.Enabled = false;
            btnDown.Enabled = false;
            btnStartWeiterSpeichern.Enabled = false;
            btnRepeat.Enabled = false;
            btnStartWeiterSpeichern.Text = "&Start";
            ttpbtnStartWeiterSpeichern.SetToolTip(btnStartWeiterSpeichern, "Start des Prüfablaufs");
            btnCancel.Enabled = false;

            tbxUmess.BackColor = SystemColors.Window;
            tbxImess.BackColor = SystemColors.Window;

            tbxUmin.Clear(); tbxUmax.Clear(); tbxImin.Clear(); tbxImax.Clear(); tbxSeriennummer.Clear();
            tbxUmess.Clear(); tbxImess.Clear(); tbxVisuellePrüfung.Clear(); tbxLichtstärke.Clear(); tbxFarbort.Clear();
            tbxsetU.Clear(); tbxsetI.Clear(); tbxsetf.Clear(); tbxWave.Clear(); tbxOut.Clear();
            tbxRangeU.Clear(); tbxResU.Clear(); tbxRangeI.Clear(); tbxResI.Clear();
            tbxUdiff.Clear(); tbxSWeite.Clear();
            tbxaPhasen.Clear(); tbxaPrüfschritte.Clear(); tbxlowerTemp.Clear(); tbxupperTemp.Clear();
            tbxzPhasen.Clear(); tbxzPrüfschritt.Clear(); tbxU.Clear();
            tbxAuftrag.Clear();

            btnLaser.BackColor = SystemColors.Control;

            lblTempData.Text = "Temperatur";
            lblTempData.BackColor = SystemColors.Control;

            coboxSachnummer.Enabled = true;

            if (coboxSachnummer.Text != "")
            {
                tbxSeriennummer.Enabled = true;
                tbxSeriennummer.Focus();
            }
        }

        private void SetVarForNextPschritt()
        {
            bitRegelung = false; bitupdown_manuell = false; bitAbschaltung = false;
            ErgPrfgU = ""; ErgPrfgI = ""; ErgPrfgAb = ""; ErgPrfgCo = ""; ErgPrfgVi = ""; 
            bitPrfgAb = false;
            zTaste = 0; zPhasen = 1;
        }

        private void ResetVar()
        {
            // Variable DigIO rücksetzen

            OutDigIO = 0;
            SumOutDigIO = 0;

            // Variablen rücksetzen

            bitStatusPrüfung = false; bitRegelung = false; bitupdown_manuell = false; bitSetIndicatorLamp = false;
            bitSetLaser = false; bitSetConfigurationDMM_U = false; bitSetConfigurationDMM_I = false; bitAbbruch = false;
            ErgPrfgU = ""; ErgPrfgI = ""; ErgPrfgAb = ""; ErgPrfgCo = ""; ErgPrfgVi = ""; 
            bitAbschaltung = false; bitPrfgAb = false;
            m_Umax = 0; m_Umin = 0;
            zTaste = 0;

            // Variablen Rücklesen Messergebnisse und Multimeter setzen

            ReadCurr = "";
            Imess = 0;
            ReadVolt = "";
            Umess = 0;

            RangeDMM_U = "0";                         // Messbereich Spannungsmesser
            SetRangeDMM_U = "";                       // gesetzter Messbereich Spannungsmesser
            RangeDMM_I = "0";                         // Messbereich Strommesser
            SetRangeDMM_I = "";                      // Messbereich Spannungsmesser       
            Res = "";                                 // Resolution Spannungs- oder Strommesser

            // Variablen SV setzen

            setU = "";
            setI = "";
            m_I = 0;
            setf = "";
            setW = "";

            // Variablen Parametersatz aus DB

            // Freigabe = false;
            //aPrüfschritte = 0;      // Anzahl Prüfschritte
            zPrüfschritt = 1;       // Zähler Prüfschritt
            zPhasen = 1;
            setUart = "";           // String für Spannungsart, AC oder DC

            // Variablen für Regelung

            Udiff = 0;
            Idiff = 0;
            m_SWeite = 0;


            Serialnummer = "";
            Auftragsnummer = "";
            LastSachnummer = Sachnummer;

            //LastLimesTestID = 0;
            LimesTestID = 0;

            existSerialnummer = false;
        }

        // *********************************************************** RS232 Digital I/O *********************************************************************

        // Konfiguration RS232 aus TXT-Datei lesen

        private void ReadConfigtxt()
        {
            int i = 0;
            Config = new string[8];

            int ASubstring = 0;
            int ESubstring = 0;
            int LSubstring = 0;


            FileStream fs = new FileStream("D:\\Programme\\Prüfung LED Signalgeber\\ConfigLED.txt", FileMode.Open);
            StreamReader sr = new StreamReader(fs);

            while (sr.Peek() != -1)
            {
                // configuration = sr.ReadLine(); 
                Config[i] = sr.ReadLine();
                i++;
            }
            sr.Close();

            // String aus Array abrufen und PortName extrahieren

            PortName = Config[3];
            ASubstring = PortName.IndexOf(": ");
            ASubstring = ASubstring + 2;
            ESubstring = PortName.IndexOf(";");
            LSubstring = ESubstring - ASubstring;
            PortName = PortName.Substring(ASubstring, LSubstring);

            // String aus Array abrufen und BaudRate extrahieren

            strCOM = Config[4];
            ASubstring = strCOM.IndexOf(": ");
            ASubstring = ASubstring + 2;
            ESubstring = strCOM.IndexOf(";");
            LSubstring = ESubstring - ASubstring;
            strCOM = strCOM.Substring(ASubstring, LSubstring);
            BaudRate = Convert.ToInt16(strCOM);

            // String aus Array abrufen und DataBits extrahieren

            strCOM = Config[5];
            ASubstring = strCOM.IndexOf(": ");
            ASubstring = ASubstring + 2;
            ESubstring = strCOM.IndexOf(";");
            LSubstring = ESubstring - ASubstring;
            strCOM = strCOM.Substring(ASubstring, LSubstring);
            DataBits = Convert.ToInt32(strCOM);

            // String aus Array abrufen und StopBits extrahieren

            strCOM = Config[6];
            ASubstring = strCOM.IndexOf(": ");
            ASubstring = ASubstring + 2;
            ESubstring = strCOM.IndexOf(";");
            LSubstring = ESubstring - ASubstring;
            strCOM = strCOM.Substring(ASubstring, LSubstring);
            StopBits = (System.IO.Ports.StopBits)Enum.Parse(typeof(System.IO.Ports.StopBits), strCOM);

            // String aus Array abrufen und Parity extrahieren

            strCOM = Config[7];
            ASubstring = strCOM.IndexOf(": ");
            ASubstring = ASubstring + 2;
            ESubstring = strCOM.IndexOf(";");
            LSubstring = ESubstring - ASubstring;
            strCOM = strCOM.Substring(ASubstring, LSubstring);
            Parity = (System.IO.Ports.Parity)Enum.Parse(typeof(System.IO.Ports.Parity), strCOM);
        }

        // Parametrierung der RS232 Übernahme der Daten aus Funktion der Form F_Einstellen

        public static void GetConfigSerPort(string PortName, string DefaultBaudRate, string DefaultDataBits, string DefaultStopbits, string DefaultParity)
        {
            GetPortName = PortName;
            GetBaudRate = Convert.ToInt16(DefaultBaudRate);
            GetDataBits = Convert.ToInt32(DefaultDataBits);
            GetStopBits = (System.IO.Ports.StopBits)Enum.Parse(typeof(System.IO.Ports.StopBits), DefaultStopbits);
            GetParity = (System.IO.Ports.Parity)Enum.Parse(typeof(System.IO.Ports.Parity), DefaultParity);
        }

        // Übergabe der Parameter an serielle Schnittstelle, lesen der Rückantwort, Fehlerabfrage

        public void SetConfigSerPort()
        {
            F_Einstellungen F_Einstellungen = new F_Einstellungen();
            F_Einstellungen.SetSerPortDefaultParameter();

            serPortDigIO.PortName = GetPortName;
            serPortDigIO.BaudRate = GetBaudRate;
            serPortDigIO.DataBits = GetDataBits;
            serPortDigIO.StopBits = GetStopBits;
            serPortDigIO.Parity = GetParity;
        }

        // Setzen der Prüflampe und des Lasers

        private void SetIndicatorLamp()
        {
            if (bitStatusPrüfung && (!(bitSetIndicatorLamp)))
            {
                bitSetIndicatorLamp = true;
                OutDigIO = 1;
                SetOutDigIO(OutDigIO);
                serPortDigIO_AusgabeRückanwort();
            }
            else
            {
                bitSetIndicatorLamp = false;
                OutDigIO = -1;
                SetOutDigIO(OutDigIO);
                serPortDigIO_AusgabeRückanwort();
            }
        }

        private void SetLaser()
        {

            if (btnLaser.BackColor != Color.Red && (!(bitSetLaser)))
            {
                bitSetLaser = true;

                F_WarnLaser WarnLaser = new F_WarnLaser();
                WarnLaser.ShowDialog();

                btnLaser.BackColor = Color.Red;
                OutDigIO = 2;
                SetOutDigIO(OutDigIO);
                serPortDigIO_AusgabeRückanwort();
            }
            else if (btnLaser.BackColor == Color.Red && bitSetLaser)
            {
                bitSetLaser = false;
                btnLaser.BackColor = SystemColors.Control;
                OutDigIO = -2;
                SetOutDigIO(OutDigIO);
                serPortDigIO_AusgabeRückanwort();
            }
        }

        // Setzen der Ausgänge der DigIO

        private void SetOutDigIO(int OutDigIO)
        {
            if (OutDigIO != 0)
            {
                SumOutDigIO = SumOutDigIO + this.OutDigIO;
                PrüfOutDigIO = SumOutDigIO ^ mask;
                set = Convert.ToByte(SumOutDigIO);
                prüf = Convert.ToByte(PrüfOutDigIO);

                byte[] send = new byte[5] { 0x63, 0x31, 0x39, set, prüf };

                try
                {
                    if (!(serPortDigIO.IsOpen))
                    {
                        try
                        {
                            serPortDigIO.Open();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Com- Port der Digital-I/O kann nicht geöffnet werden. \n" + ex.Message, "Fehler DigIO Open Port (SetOutDigIO)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            // SetOutDigIO(OutDigIO);
                        }
                    }
                    serPortDigIO.Write(send, 0, 5);
                    // serPortDigIO.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Fehler beim Senden an die Digital-I/O. \n" + ex.Message, "Fehler DigIO (SetOutDigIO)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //  SetOutDigIO(OutDigIO);
                }

                tbxSumOutDigIO.Text = Convert.ToString(SumOutDigIO);
                tbxOutDigIO.Text = Convert.ToString(OutDigIO);
                tbxset.Text = Convert.ToString(set);
                tbxprüf.Text = Convert.ToString(prüf);
            }
        }

        // Reset DigIO

        private void ResetOutDigIO()
        {
            OutDigIO = SumOutDigIO * (-1);
            SetOutDigIO(OutDigIO);
        }

        // Lesen der Rückantort

        private void serPortDigIO_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            readserportDigIO = new byte[6];

            try
            {
                serPortDigIO.Read(readserportDigIO, 0, 6);
                serPortDigIO.Close();

                if (readserportDigIO[0] != 13 || readserportDigIO[1] != 10 || readserportDigIO[2] != 111 ||
                    readserportDigIO[3] != 107 || readserportDigIO[4] != 13 || readserportDigIO[5] != 10)
                {
                    serPortDigIO.DiscardInBuffer();
                   // SetOutDigIO(OutDigIO);
                }

            }
            catch (ArgumentNullException ex)
            {
                MessageBox.Show("Die Kommunikation mit der Dig I/O ist nicht möglich. \n \n" + ex,
                                "serPortDigIO_DataReceived (ArgumentNullException) ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show("Die Kommunikation mit der Dig I/O ist nicht möglich. \n \n" + ex,
                                "serPortDigIO_DataReceived (InvalidOperationException) ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                MessageBox.Show("Die Kommunikation mit der Dig I/O ist nicht möglich. \n \n" + ex,
                                "serPortDigIO_DataReceived (ArgumentOutOfRangeException) ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (ArgumentException ex)
            {
                MessageBox.Show("Die Kommunikation mit der Dig I/O ist nicht möglich. \n \n" + ex,
                                "serPortDigIO_DataReceived (ArgumentException) ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (TimeoutException ex)
            {
                MessageBox.Show("Die Kommunikation mit der Dig I/O ist nicht möglich. \n \n" + ex,
                                "serPortDigIO_DataReceived (TimeoutException) ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void serPortDigIO_AusgabeRückanwort()
        {
            tbxReadByte1.Clear(); tbxReadByte2.Clear(); tbxReadByte3.Clear();
            tbxReadByte4.Clear(); tbxReadByte5.Clear(); tbxReadByte6.Clear();

            try
            {
                tbxReadByte1.Text = readserportDigIO[0].ToString();
                tbxReadByte2.Text = readserportDigIO[1].ToString();
                tbxReadByte3.Text = readserportDigIO[2].ToString();
                tbxReadByte4.Text = readserportDigIO[3].ToString();
                tbxReadByte5.Text = readserportDigIO[4].ToString();
                tbxReadByte6.Text = readserportDigIO[5].ToString();
            }
            catch { }
        }

        // Fehlerbehandlung Serialport

        private void serPortDigIO_ErrorReceived(object sender, SerialErrorReceivedEventArgs e)
        {
            if (serPortDigIO.IsOpen)
            {
                SetOutDigIO(OutDigIO);
                serPortDigIO_AusgabeRückanwort();
            }
            else
            {
                try
                {
                    serPortDigIO.Open();
                    SetOutDigIO(OutDigIO);
                    serPortDigIO_AusgabeRückanwort();
                }
                catch (UnauthorizedAccessException ex)
                {
                    MessageBox.Show("" + ex,
                                    "serPortDigIO_ErrorReceived (UnauthorizedAccessException)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (ArgumentOutOfRangeException ex)
                {
                    MessageBox.Show("" + ex,
                                   "serPortDigIO_ErrorReceived (ArgumentOutOfRangeException)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
                catch (ArgumentException ex)
                {
                    MessageBox.Show("" + ex,
                                   "serPortDigIO_ErrorReceived (ArgumentException)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
                catch (IOException ex)
                {
                    MessageBox.Show("" + ex,
                                   "serPortDigIO_ErrorReceived (IOException)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
                catch (InvalidOperationException ex)
                {
                    MessageBox.Show("" + ex,
                                   "serPortDigIO_ErrorReceived (InvalidOperationException)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
            }
        }

        // *************************************************************** Prüfschritte **********************************************************************
        
        private void WriteParameterinVariables(int zPhasen)
        {
            if (zPhasen == 1)
            {
                m_U = U1[zPrüfschritt];
                m_lowtol = lowtolU1[zPrüfschritt];
                m_uptol = uptolU1[zPrüfschritt];
                m_Imin = I1min[zPrüfschritt];
                m_Imax = I1max[zPrüfschritt];
                m_SWeite = Sweite1[zPrüfschritt];
                m_Rp = Rp[zPrüfschritt];
                m_Uart = Uart[zPrüfschritt];
                m_Freq = f[zPrüfschritt];
                m_PrfgU = PrfgU[zPrüfschritt];
                m_PrfgI = PrfgI[zPrüfschritt];
                m_PrfgVi = PrfgVi[zPrüfschritt];
                m_PrfgAb = PrfgAb[zPrüfschritt];
                m_PrfgCo = PrfgCo[zPrüfschritt];
            }
            else if (zPhasen == 2)
            {
                m_U = U2[zPrüfschritt];
                m_lowtol = lowtol2[zPrüfschritt];
                m_uptol = uptol2[zPrüfschritt];
                m_Imin = I2min[zPrüfschritt];
                m_Imax = I2max[zPrüfschritt];
                m_SWeite = Sweite2[zPrüfschritt];
            }

            if (m_Rp == "U")
                m_ResSV = 0.01;
            else if (m_Rp == "I")
                m_ResSV = 0.001;
            else
            {
                MboxText = Environment.NewLine +
                           "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                           Environment.NewLine +
                           "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                MboxTitle = "Fehler Auswahl Prüfung";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DialogResult.OK == DRM)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
        }

        private void LoadParameterInObjects()
        {
            tbxUmin.Text = Convert.ToString(m_Umin);
            tbxUmax.Text = Convert.ToString(m_Umax);
            tbxImin.Text = Convert.ToString(m_Imin);
            tbxImax.Text = Convert.ToString(m_Imax);
            tbxzPhasen.Text = Convert.ToString(zPhasen);
            tbxVisuellePrüfung.Text = Convert.ToString(VP[zPrüfschritt]);
            tbxSWeite.Text = Convert.ToString(m_SWeite);
            tbxU.Text = Convert.ToString(m_U);
            tbxLichtstärke.Text = Convert.ToString(Lichtstärke[zPrüfschritt]);
            tbxFarbort.Text = Convert.ToString(Farbort[zPrüfschritt]);

            switch (m_Rp)
            {
                case "U":
                    lbltBarmin.Text = Convert.ToString(Math.Round(m_U - m_lowtol, 2)) + " V";
                    lbltBarmax.Text = Convert.ToString(Math.Round(m_U + m_uptol, 2)) + " V";
                    tbarUIMinMax.Maximum = Convert.ToInt32(((m_uptol + m_lowtol) / m_ResSV) + 6);
                    tbarUIMinMax.Value = tbarUIMinMax.Maximum / 2;
                    OldValue = tbarUIMinMax.Value;
                    tbxSWtBarOld.Text = Convert.ToString(OldValue);
                    break;

                case "I":
                    lbltBarmin.Text = Convert.ToString(m_Imin) + " mA";
                    lbltBarmax.Text = Convert.ToString(m_Imax) + " mA";
                    tbarUIMinMax.Maximum = Convert.ToInt32(((m_Imax - m_Imin) / m_ResSV) + 4);
                    tbarUIMinMax.Value = tbarUIMinMax.Maximum / 2;
                    OldValue = tbarUIMinMax.Value;
                    tbxSWtBarOld.Text = Convert.ToString(OldValue);
                    break;

                default:
                    break;
            }
        }

        private int CheckNumbersofPhases()
        {
            if (U2[zPrüfschritt] != 0 && m_Rp == "U")
            {
                aPhasen = 2;
            }
            else if (U2[zPrüfschritt] != 0 && m_Rp == "I")
            {
                F_MessageBox.Show(Environment.NewLine +
                                  "Für den Regelparameter " + m_Rp + " steht die ausgewählte Prüfung nicht zur Verfügung.\n" +
                                  "Bitte berichtigen Sie den entsprechenden Datensatz in der Datenbank.",
                                  "Fehler Auswahl Prüfung",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                aPhasen = 1;
            }

            tbxaPhasen.Text = Convert.ToString(aPhasen);

            return aPhasen;
        }

        private void CalcTol(int Phasen)
        {
            if (zPhasen == 1)
            {
                m_Umin = U1[zPrüfschritt] - lowtolU1[zPrüfschritt];
                m_Umin = Math.Round(m_Umin, 3);
                m_Umax = U1[zPrüfschritt] + uptolU1[zPrüfschritt];
                m_Umax = Math.Round(m_Umax, 3);
                m_I = m_Imin + ((m_Imax - m_Imin) / 2);
            }
            else if (zPhasen == 2)
            {
                m_Umin = U2[zPrüfschritt] - lowtol2[zPrüfschritt];
                m_Umin = Math.Round(m_Umin, 3);
                m_Umax = U2[zPrüfschritt] + uptol2[zPrüfschritt];
                m_Umax = Math.Round(m_Umax, 3);

                //m_U2min = U2[zPrüfschritt] - lowtol2[zPrüfschritt];
                //m_U2min = Math.Round(m_U2min, 3);
                //m_U2max = U2[zPrüfschritt] + uptol2[zPrüfschritt];
                //m_U2max = Math.Round(m_U2max, 3);
            }
        }

        private void StartWeiterSpeichernPrüfablauf()
        {
            switch (btnStartWeiterSpeichern.Text)
            {
                case "&Start":

                    SetObjDuringMeasurement();

                    btnStartWeiterSpeichern.Text = "&Weiter";

                    tbxaPhasen.Text = Convert.ToString(aPhasen);                    // Ausgabe der Variable "aPhasen" in Textbox
                    tbxzPhasen.Text = Convert.ToString(zPhasen);                    // Ausgabe der Variable "zPhasen" in Textbox
                    tbxaPrüfschritte.Text = Convert.ToString(aPrüfschritte);
                    tbxzPrüfschritt.Text = Convert.ToString(zPrüfschritt);

                    MeasureTemperature();

                    if (!bitAbbruch)
                    {
                        SetIndicatorLamp();
                        CheckNumbersofPhases();
                        CalcTol(1);
                        bgwSetConfigurationDMM.RunWorkerAsync();
                        Thread.Sleep(1000);
                        SelectionTestCycle(Typenbezeichnung);
                    }
                    else
                        AbortTest();

                    break;

                case "&Weiter":

                    SetOutputSV_ACDC("1");

                    AppraisalVisualTest();
                    AppraisalAutoTests();
                    SaveResults();
                    SetObjForNextPschritt();
                    SetVarForNextPschritt();
                    SetObjDuringMeasurement();

                    zPrüfschritt++;
                    tbxzPrüfschritt.Text = Convert.ToString(zPrüfschritt);

                    if (zPrüfschritt == aPrüfschritte)
                    {
                        btnStartWeiterSpeichern.Text = "&Speichern";
                    }

                    WriteParameterinVariables(1);
                    CheckNumbersofPhases();
                    CalcTol(1);
                    bgwSetConfigurationDMM.RunWorkerAsync();
                    Thread.Sleep(1000);
                    LoadParameterInObjects();
                    SelectionTestCycle(Typenbezeichnung);
                    break;

                case "&Speichern":

                    if (Typenbezeichnung == "LED 136 Matrix-BG")
                    {

                        // Messwerte in DB speichern

                        SetDateTime();
                        CheckSerialnummerExist();

                        if (existSerialnummer)
                            UpdateResultsinDB();
                        else
                            SaveResultsinDB();

                        // Konfiguration für nächsten Prüfling

                        SetOutputSV_ACDC("1");
                        SetIndicatorLamp();
                        SetLaser();
                        ResetVar();
                        ResetObj();
                        LoadParameterInObjects();
                    }
                    else
                    {

                        LoadLimesTestIDDB();
                        CheckSerialnummerExist();

                        if (LimesTestID != 0)
                        {

                            // Messwerte in DB speichern

                            AppraisalVisualTest();
                            AppraisalAutoTests();
                            SaveResults();
                            SetDateTime();

                            if (existSerialnummer)
                                UpdateResultsinDB();
                            else
                                SaveResultsinDB();

                            // Konfiguration für nächsten Prüfling

                            SetOutputSV_ACDC("1"); ;
                            //SetIndicatorLamp();
                            //SetLaser();
                            ResetOutDigIO();
                            ResetVar();
                            ResetObj();
                            LoadParameterInObjects();
                        }
                        else
                        {
                            MessageBox.Show("Bitte speichern Sie die Messdaten der optischen Prüfung ab.\n" +
                                            "Anschließend können Sie die Messdaten der elektrischen Prüfung speichern.",
                                            "Fehler Messdaten speichern",
                                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    break;

                default:
                    bitAbbruch = true;
                    AbortTest();
                    break;
            }
        }

        public static void GetNumberPrüfschrittToRepeat(int zPrüfschrittF_Repeat)
        {
            zPrüfschritt = zPrüfschrittF_Repeat;
            DRR = DialogResult.OK;
        }

        private void RepeatPrüfschritt(DialogResult DRR)
        {
            if (DRR == DialogResult.OK)
            {
                SetOutputSV_ACDC("1");
                SetObjForNextPschritt();
                SetVarForNextPschritt();
                SetObjDuringMeasurement();

                tbxzPrüfschritt.Text = Convert.ToString(zPrüfschritt);

                if (zPrüfschritt < aPrüfschritte)
                {
                    btnStartWeiterSpeichern.Text = "&Weiter";
                }
                else
                {
                    btnStartWeiterSpeichern.Text = "&Speichern";
                }

                WriteParameterinVariables(zPhasen);
                CheckNumbersofPhases();
                CalcTol(zPhasen);
                bgwSetConfigurationDMM.RunWorkerAsync();
                Thread.Sleep(1000);
                LoadParameterInObjects();
                SelectionTestCycle(Typenbezeichnung);
            }

        }

        private void SelectionTestCycle(string Typenbezeichnung)
        {
            switch (Typenbezeichnung)
            {
                case "LED 70":
                    TestCycle_LED70();
                    break;

                case "HLED 70":
                    TestCycle_HLED70();
                    break;

                case "SIO D 70 Basic":
                    TestCycle_SIO_D_70_Basic();
                    break;

                case "LED 136":
                    TestCycle_LED136();
                    break;

                case "SIO D 136 Basic":
                    TestCycle_SIO_D_136_Basic();
                    break;

//                case "LED 136 Matrix-BG":
//                    TestCycle_LED136_Matrix_BG();
//                    break;

                case "LED 210":
                    TestCycle_LED210();
                    break;

                case "LED 210-1":
                    TestCycle_LED210_1();
                    break;

                case "LED Anzeigemodul":
                    TestCycle_LED_Anzeigemodul();
                    break;

                case "MA 480":
                    TestCycle_MA480();
                    break;

                default:

                    if (Typenbezeichnung != "")
                    {
                        ReadPosActivForm();

                        F_MessageBox.Show(Environment.NewLine +
                                          "Für den in der Datenbank angegebenen Prüflingstyp" + Typenbezeichnung + " existiert keine Prüfroutine.\n" +
                                          "Bitte berichtigen Sie den entsprechenden Datensatz in der Datenbank.",
                                          "Fehler Auswahl Prüfungling",
                                          MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);
                    }
                    else
                    {
                        ReadPosActivForm();

                        F_MessageBox.Show(Environment.NewLine +
                                          "Es ist keine Typenbezeichnung für den Signalgeber in der Datenbank angegeben..\n" +
                                          Environment.NewLine +
                                          "Bitte berichtigen Sie den entsprechenden Datensatz in der Datenbank.",
                                          "Fehler Auswahl Prüfungling",
                                          MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);
                    }

                    bitAbbruch = true;
                    AbortTest();
                    break;
            }
        }

        private void TestCycle_LED70()
        {
            if ((m_PrfgU || m_PrfgI) && (!(m_PrfgAb && m_PrfgCo)))
            {
                TestU_I();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgCo) && m_PrfgAb))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den LED70- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                            Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgAb LED70";

                F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                  posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb) && m_PrfgCo))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den LED70- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                            Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgCo LED70";

                F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                  posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Für den LED70- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung LED70- Signalgeber";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DialogResult.OK == DRM)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
        }

        private void TestCycle_HLED70()
        {
            if ((m_PrfgU || m_PrfgI) && !m_PrfgAb && !m_PrfgCo)
            {
                TestU_I();
            }
            else if ((!m_PrfgU && !m_PrfgCo && m_PrfgI && m_PrfgAb))
            {
                TestAb();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb) && m_PrfgCo))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den HLED70- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                            Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgCo HLED70";

                F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                  posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den HLED70- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                            Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung HLED70";

                F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                  posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
        }

        private void TestCycle_SIO_D_70_Basic()
        {
            if ((m_PrfgU || m_PrfgI) && (!(m_PrfgAb && m_PrfgCo)))
            {
                TestU_I();
            }
            else if (!(m_PrfgU || m_PrfgI && m_PrfgAb && m_PrfgCo) && m_PrfgVi)
            {
                TestU_I();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgCo) && m_PrfgAb))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den SIO D 70 Basic- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                            Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgAb SIO D 70 Basic";

                F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                  posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb) && m_PrfgCo))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine + Environment.NewLine +
                          "Für den SIO D 70 Basic-  Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgCo SIO D 70 Basic";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Für den SIO D 70 Basic- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung SIO D 70 Basic";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
        }
        
        private void TestCycle_LED136()
        {
            if ((m_PrfgU || m_PrfgI) && (!(m_PrfgAb && m_PrfgCo)))
            {
                TestU_I();
            }
            else if (!(m_PrfgU || m_PrfgI && m_PrfgAb && m_PrfgCo) && m_PrfgVi)
            {
                TestU_I();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgCo) && m_PrfgAb))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den LED136- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                            Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgAb LED136";

                F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                  posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb) && m_PrfgCo))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine + Environment.NewLine +
                          "Für den LED136- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgCo LED136";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Für den LED136- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung LED136";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
        }

        private void TestCycle_SIO_D_136_Basic()
        {
            if ((m_PrfgU || m_PrfgI) && (!(m_PrfgAb && m_PrfgCo)))
            {
                TestU_I();
            }
            else if (!(m_PrfgU || m_PrfgI && m_PrfgAb && m_PrfgCo) && m_PrfgVi)
            {
                TestU_I();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgCo) && m_PrfgAb))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den SIO D 136 Basic- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                            Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgAb SIO D 136 Basic";

                F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                  posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb) && m_PrfgCo))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine + Environment.NewLine +
                          "Für den SIO D 136 Basic-  Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgCo SIO D 136 Basic";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Für den SIO D 136 Basic- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung SIO D 136 Basic";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
        }
        
        private void TestCycle_LED136_Matrix_BG()
        {
            if (!(m_PrfgU && m_PrfgI && m_PrfgAb && m_PrfgCo) && m_PrfgVi)
            {
                TestU_I();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb)) && m_PrfgCo)
            {
                TestCo();
            }
            else
            {
                ReadPosActivForm();

                F_MessageBox.Show(Environment.NewLine +
                                  "Für die LED136 Matrix-BG steht die ausgewählte Prüfung nicht zur Verfügung." +
                                  Environment.NewLine +
                                  "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.",
                                  "Fehler Auswahl Prüfung PrfgCo LED136 Matrix-BG",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
        }

        private void TestCycle_LED210()
        {
            if ((m_PrfgU || m_PrfgI) && (!(m_PrfgAb && m_PrfgCo)))
            {
                TestU_I();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgCo) && m_PrfgAb))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den LED210- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgAb LED210";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb) && m_PrfgCo))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den LED210- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgCo LED210";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Für den LED210- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung LED210- Signalgeber";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
        }

        private void TestCycle_LED210_1()
        {
            if ((m_PrfgU || m_PrfgI) && (!(m_PrfgAb && m_PrfgCo)))
            {
                TestU_I();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgCo) && m_PrfgAb))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den LED210_1- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgAb LED210_1";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb) && m_PrfgCo))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für den LED210_1- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgCo LED210_1";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Für den LED210_1- Signalgeber steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung LED210_1- Signalgeber";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
        }

        private void TestCycle_LED_Anzeigemodul()
        {
            if ((m_PrfgU || m_PrfgI) && (!(m_PrfgAb && m_PrfgCo)))
            {
                TestU_I();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgCo) && m_PrfgAb))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine + 
                           "Für das LED Anzeigemodul steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgAb LED Anzeigemodul";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb) && m_PrfgCo))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für das LED Anzeigemodul steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgCo LED Anzeigemodul";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Für das LED Anzeigemodul steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung LED Anzeigemodul";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
        }

        private void TestCycle_MA480()
        {
            if ((m_PrfgU || m_PrfgI) && (!(m_PrfgAb && m_PrfgCo)))
            {
                TestU_I();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgCo) && m_PrfgAb))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für das LED Anzeigemodul MA 480 steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgAb LED Anzeigemodul MA480";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else if ((!(m_PrfgU && m_PrfgI && m_PrfgAb) && m_PrfgCo))
            {
                ReadPosActivForm();

                MboxText = Environment.NewLine +
                           "Für das LED Anzeigemodul MA 480 steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung PrfgCo LED Anzeigemodul MA 480";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
            else
            {
                MboxText = Environment.NewLine +
                           "Für das LED Anzeigemodul MA 480 steht die ausgewählte Prüfung nicht zur Verfügung." +
                           Environment.NewLine +
                           "Bitte berichtigen Sie den Prüfschritt " + ID_Pschritt + " in der Datenbank.";

                MboxTitle = "Fehler Auswahl Prüfung LED Anzeigemodul MA 480";

                DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }
        }

        private void TestU_I()
        {
            switch (m_Uart)
            {
                case "AC":

                    switch (m_Rp)
                    {
                        case "U":

                            ReadPosActivForm();

                            DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                    Environment.NewLine + Environment.NewLine +
                                                    "Spannung wird auf " + m_U + "V eingeregelt.",
                                                    "Bedienerhinweis " + Typenbezeichnung + " PrfgU_I",
                                                    MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.Cancel == DRM)                        
                                bitAbbruch = true;

                                bgwUAC_regeln.RunWorkerAsync();

                                ReadPosActivForm();

                                DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                        Environment.NewLine + Environment.NewLine +
                                                       "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                       "Bedienerhinweis",
                                                        MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                                if (DialogResult.Cancel == DRM)
                                    bitAbbruch = true;

                                if (aPhasen == 2 && !bitAbbruch)
                                {
                                    zPhasen = 2;

                                    WriteParameterinVariables(zPhasen);
                                    CalcTol(zPhasen);
                                    LoadParameterInObjects();

                                    ReadPosActivForm();

                                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                            Environment.NewLine +
                                                            "In Phase " + zPhasen + " wird Spannung jetzt auf " + m_U + "V eingeregelt.",
                                                            "Bedienerhinweis " + Typenbezeichnung + " PrfgU_I",
                                                            MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                                    if (DialogResult.Cancel == DRM)
                                        bitAbbruch = true;

                                    bgwUAC_regeln.RunWorkerAsync();

                                    ReadPosActivForm();

                                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                            Environment.NewLine + Environment.NewLine +
                                                           "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                           "Bedienerhinweis",
                                                            MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                                    if (DialogResult.Cancel == DRM)
                                        bitAbbruch = true;


                                }
                            break;

                        case "I":
                            
                            ReadPosActivForm();

                            DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                    Environment.NewLine + Environment.NewLine +
                                                    "Strom wird auf " + m_I + "mA eingeregelt.",
                                                    "Bedienerhinweis " + Typenbezeichnung + " PrfgU_I",
                                                    MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.Cancel == DRM)
                                bitAbbruch = true;                               

                                bgwIAC_regeln.RunWorkerAsync();

                                ReadPosActivForm();

                                switch (Manufacturer[2])
                                {
                                    case "HBS-ELECTRONIC":
                                        DRMCC = F_Message_CC_Mode.Show("Wait for CC-Mode", 15, new F_Message_CC_Mode(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);
                                        break;

                                    case "SCHULTZ ELECTRONIC":
                                        DRMCC = F_Message_CC_Mode.Show("Wait for CC-Mode", 5, new F_Message_CC_Mode(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);
                                        break;
                                }

                                if (DialogResult.Cancel == DRMCC)                       
                                    bitAbbruch = true;
                                    
                                else if (DialogResult.OK == DRMCC)
                                {
                                    bitSV_CC_Mode = true;

                                    bgwIAC_regeln.RunWorkerAsync();

                                    ReadPosActivForm();

                                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                            Environment.NewLine + Environment.NewLine +
                                                           "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                           "Bedienerhinweis",
                                                            MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                                    if (DialogResult.Cancel == DRM)
                                        bitAbbruch = true;
                           
                            }
                            break;

                        default:

                            MboxText = Environment.NewLine +
                                       "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                                       Environment.NewLine +
                                       "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                            MboxTitle = "Fehler Auswahl Prüfung PrfgU_I";

                            DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                                    posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.OK == DRM)
                            {
                                bitAbbruch = true;
                                AbortTest();
                            }
                            break;
                    }
                    break;


                case "DC":

                    switch (m_Rp)
                    {
                        case "U":

                            ReadPosActivForm();

                            DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                    Environment.NewLine + Environment.NewLine +
                                                    "Spannung wird auf " + m_U + "V eingeregelt.",
                                                    "Bedienerhinweis " + Typenbezeichnung + " PrfgU_I",
                                                    MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.Cancel == DRM)
                                                            bitAbbruch = true;
                              bgwUDC_regeln.RunWorkerAsync();
                            
                                ReadPosActivForm();

                                DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                        Environment.NewLine + Environment.NewLine +
                                                       "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                       "Bedienerhinweis",
                                                        MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                                if (DialogResult.Cancel == DRM)
                                    bitAbbruch = true;

                                if (aPhasen == 2 && !bitAbbruch)
                                {
                                    zPhasen = 2;

                                    WriteParameterinVariables(zPhasen);
                                    CalcTol(zPhasen);
                                    LoadParameterInObjects();

                                    ReadPosActivForm();

                                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                            Environment.NewLine +
                                                            "In Phase " + zPhasen + " wird Spannung jetzt auf " + m_U + "V eingeregelt.",
                                                            "Bedienerhinweis " + Typenbezeichnung + " PrfgU_I",
                                                            MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                                    if (DialogResult.Cancel == DRM)
                                        bitAbbruch = true;

                                    bgwUDC_regeln.RunWorkerAsync();

                                    ReadPosActivForm();

                                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                            Environment.NewLine + Environment.NewLine +
                                                           "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                           "Bedienerhinweis",
                                                            MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                                    if (DialogResult.Cancel == DRM)
                                        bitAbbruch = true;

                                }
                            break;//case DC; case U

                        case "I":

                            ReadPosActivForm();

                            DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                    Environment.NewLine + Environment.NewLine +
                                                    "Strom wird auf " + m_I + "mA eingeregelt.",
                                                    "Bedienerhinweis " + Typenbezeichnung + " PrfgU_I",
                                                    MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.Cancel == DRM)
                                bitAbbruch = true;
                               

                                bgwIDC_regeln.RunWorkerAsync();


                                ReadPosActivForm();

                                DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                                        Environment.NewLine + Environment.NewLine +
                                                       "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                                       "Bedienerhinweis",
                                                        MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                                if (DialogResult.Cancel == DRM)
                                    bitAbbruch = true;                    
                            break;

                        default:

                            MboxText = Environment.NewLine +
                                       "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                                       Environment.NewLine +
                                       "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                            MboxTitle = "Fehler Auswahl Prüfung PrfgU_I";

                            DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                                    posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                            if (DialogResult.OK == DRM)
                            {
                                bitAbbruch = true;
                                AbortTest();
                            }
                            break;//case DC; case I
                    }
                    break;//case DC

                default:

                    MboxText = Environment.NewLine +
                               "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                               Environment.NewLine +
                               "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                    MboxTitle = "Fehler Auswahl Prüfung PrfgU_I";

                    DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                            posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.OK == DRM)
                    {
                        bitAbbruch = true;
                        AbortTest();
                    }
                    break;
            }
        }
       
        private void TestCo()
        {
            switch (m_Rp)
            {
                case "U":

                    ReadPosActivForm();

                    DRM = F_MessageBox.Show(Environment.NewLine + Environment.NewLine +
                                            "Für die " + Typenbezeichnung + " steht die ausgewählte Prüfung nicht zur Verfügung.",
                                            "Fehler Auswahl Prüfung PrfgCo " + Typenbezeichnung,
                                            MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.OK == DRM)
                    {
                        bitAbbruch = true;
                        AbortTest();
                    }
                    break;

                case "I":

                    ReadPosActivForm();

                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                            Environment.NewLine + Environment.NewLine +
                                            "Strom wird auf " + m_I + "mA eingeregelt.",
                                            "Bedienerhinweis" + Typenbezeichnung + "PrfgCo",
                                            MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.Cancel == DRM)
                        bitAbbruch = true;

                    bgwIDC_regeln.RunWorkerAsync();

                    ReadPosActivForm();

                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                            Environment.NewLine + Environment.NewLine +
                                           "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                           "Bedienerhinweis",
                                            MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.Cancel == DRM)
                        bitAbbruch = true;

                    if(!(bitAbbruch))
                    {

                        ReadPosActivForm();

                        DRMC = F_MessageBox_Corona_LED_136_Matrix.Show(MessageBoxButtons.OKCancel, new F_MessageBox_Corona_LED_136_Matrix(), 
                                                                       posActFormx, posActFormy, ActFormWidth, ActFormHeight);
                                                          
                        if (DialogResult.Cancel == DRMC)
                            bitAbbruch = true;
                    }

                    if (!(bitAbbruch))
                    {
                        ReadPosActivForm();

                        DRMPP = F_LED_136_Matrix_Print_Prot.Show(Sachnummer, Auftragsnummer, Serialnummer, ID_Prüfer,
                                                                 MessageBoxButtons.OKCancel, new F_LED_136_Matrix_Print_Prot(),
                                                                 posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                        if (DialogResult.Cancel == DRMPP)
                            bitAbbruch = true;
                       
                   
                    }

                    if (bitAbbruch)
                    {
                        AbortTest();
                    }
                    break;

                default:

                    MboxText = Environment.NewLine +
                               "Die Angabe zum Regelparameter fehlt in der Datenbank." +
                               Environment.NewLine +
                               "Bitte beheben Sie dieses Problem und starten die Prüfung erneut.";

                    MboxTitle = "Fehler Auswahl Prüfung Prfg_Co";

                    DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                            posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.OK == DRM)
                    {
                        bitAbbruch = true;
                        AbortTest();
                    }
                    break;
            }
        }

        private void TestAb()
        {
            switch (m_Rp)
            {
                case "U":

                    //Start Regelvorgang auf Umin Abschaltbereich

                    ReadPosActivForm();

                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                            Environment.NewLine + Environment.NewLine +
                                            "Spannung wird auf " + m_U + "V eingeregelt.", 
                                            "Bedienerhinweis " + Typenbezeichnung + " TestAb",
                                            MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.Cancel == DRM)
                        bitAbbruch = true;

                    bgwUAC_regeln.RunWorkerAsync();

                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                            Environment.NewLine + Environment.NewLine +
                                           "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                           "Bedienerhinweis",
                                            MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.Cancel == DRM)
                        bitAbbruch = true;

                    //Start Regelvorgang auf Umax Abschaltbereich, nach jeder Messung wird auf Abschalten des Signalgebers geprüft

                    ReadPosActivForm();

                    DRM = F_MessageBox.Show("Die erforderliche Spannung zum Start der Abschaltprüfung, von " + m_U + " V ist eingeregelt." +
                                            Environment.NewLine +
                                            "Bestätigen Sie mit OK, um die Abschaltprüfung zu starten, Cancel für Abbruch.",
                                            "Prüfung Abschaltung " + Typenbezeichnung,
                                            MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.Cancel == DRM)
                    {
                        bitAbbruch = true;
                        AbortTest();
                    }

                    zPhasen += 1;

                    WriteParameterinVariables(zPhasen);
                    LoadParameterInObjects();
                    CalcTol(zPhasen);

                    bgwUAC_regeln.RunWorkerAsync();

                    ReadPosActivForm();

                    DRM = F_MessageBox.Show(Environment.NewLine + "Prüfschritt: " + zPrüfschritt + " von " + aPrüfschritte +
                                            Environment.NewLine + Environment.NewLine +
                                           "Bitte warten, bis Regelvorgang abgeschlossen ist.",
                                           "Bedienerhinweis",
                                            MessageBoxButtons.RetryCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    if (DialogResult.Cancel == DRM)
                        bitAbbruch = true;

                    if (bitAbschaltung)
                    {
                        ReadPosActivForm();

                        MboxText = "Es wurde die Abschaltung des Signalgebers " + Typenbezeichnung +" detektiert." +
                                   "Bitte kontrollieren Sie, ob die Abschaltung des Signalgebers erfolgt ist." +
                                   Environment.NewLine +
                                   "Hat der Signalgebers abgeschaltet?";

                        MboxTitle = "Kontrolle Abschaltung Signalgeber";

                        DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.YesNo, new F_MessageBox(),
                                                posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                        if (DRM == DialogResult.Yes)
                        {
                            ErgPrfgAb = "JA";
                            cboxiO.Checked = true;
                        }
                        else
                        {
                            ErgPrfgAb = "NEIN";
                            cboxniO.Checked = false;
                        }
                    }
                    else
                    {
                        ReadPosActivForm();

                        MboxText = Environment.NewLine +
                                   "Es wurde keine Abschaltung des Signalgebers " + Typenbezeichnung + " detektiert." +
                                   Environment.NewLine +
                                   "Der Prüfschritt Abschaltung ist somit nicht bestanden.";

                        MboxTitle = "Kontrolle Abschaltung Signalgeber";

                        DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.OK, new F_MessageBox(),
                                                posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                        ErgPrfgAb = "NEIN";
                    }

                    break;

                case "I":

                    ReadPosActivForm();

                    F_MessageBox.Show(Environment.NewLine +
                                      "Für den Regelparameter " + m_Rp + " steht die ausgewählte Prüfung nicht zur Verfügung.\n" + 
                                      "Bitte berichtigen Sie den entsprechenden Datensatz in der Datenbank.",
                                      "Fehler Auswahl Prüfung TestAb",
                                      MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                    bitAbbruch = true;
                    AbortTest();
                    break;

                    

                default:

                    ReadPosActivForm();

                    F_MessageBox.Show(Environment.NewLine +
                                      "Für den Regelparameter " + m_Rp + " steht die ausgewählte Prüfung nicht zur Verfügung.\n" + 
                                      "Bitte berichtigen Sie den entsprechenden Datensatz in der Datenbank.",
                                      "Fehler Auswahl Prüfung TestAb",
                                      MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);
                    break;
            }
        }
        
        private void TestVi()
        {
            MboxText = VP[zPrüfschritt];
            MboxTitle = "Visuelle Prüfung";

            DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.YesNo, new F_MessageBox(), 
                                    posActFormx, posActFormy, ActFormWidth, ActFormHeight);

            if (DRM == DialogResult.Yes)
                ErgPrfgVi = "JA";
            else
                ErgPrfgVi = "NEIN";
        }

        private void bgwUAC_regeln_DoWork(object sender, DoWorkEventArgs e)
        {
            textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(bitRegelung)));
            textBox3.Invoke(new Action(() => textBox3.Text = Convert.ToString(bitupdown_manuell)));
            textBox4.Invoke(new Action(() => textBox4.Text = Convert.ToString(bitSV_CC_Mode)));

            if (bitAbbruch == true)
                e.Cancel = true;

            switch (zPhasen)
            {
                case 1:

                    if (!bitRegelung && !e.Cancel && !bitPrfgAb)
                    {
                        setU = Convert.ToString(m_U * 1.03);                 // 103% von U einstellen, wegen Spannungsfall über Speiseleitung

                        if (m_Imax == 0) m_Imax = 1500;

                        m_I = m_Imax * 1e-3 * 1.1;                          // 110% von I als Strombegrenzung einstellen
                        setI = Convert.ToString(m_I);

                        SetVoltageSV_ACDC(setU);
                        SetCurrentSV_ACDC(setI);
                        SetCurrentMaxSV_ACDC();
                        SetFrequencySV_ACDC(m_Freq);
                        SetWaveSV_ACDC(setW);
                        SetOutputSV_ACDC("0");
                       // Thread.Sleep(500);
                        Thread.Sleep(50);
                        bgwMeasureVoltage.RunWorkerAsync();
                        CheckRegelparamU();
                    }

                    if (bitRegelung && !e.Cancel && !bitPrfgAb)
                    {
                        while (bitRegelung && !e.Cancel)
                        {
                            if (bitAbbruch)
                                e.Cancel = true;

                            CalcSWeiteU(Umess, m_U);

                            if ((m_U - 0.01) < Umess)                    
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) - m_SWeite);
                            }
                            else if ((m_U + 0.01) > Umess)             
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);
                            }

                            SetObjDuringMeasurement();
                            SetVoltageSV_ACDC(setU);
                          //  Thread.Sleep(500);
                            Thread.Sleep(50);
                            bgwMeasureVoltage.RunWorkerAsync();
                            CheckRegelparamU();
                        }
                    }
                    break;      //case 1

                case 2:

                    if (!e.Cancel && !(m_PrfgAb || m_PrfgCo))
                    {
                        bitRegelung = true;

                        while (bitRegelung && !e.Cancel)
                        {
                            if (bitAbbruch)
                                e.Cancel = true;

                            CalcSWeiteU(Umess, m_U);

                            if ((m_U - 0.01) < Umess)                     //&& zPhasen == 1)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) - m_SWeite);
                            }
                            else if ((m_U + 0.01) > Umess)               //&& zPhasen == 1)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);
                            }

                            SetObjDuringMeasurement();
                            SetVoltageSV_ACDC(setU);
                          //  Thread.Sleep(500);
                            Thread.Sleep(50);
                            bgwMeasureVoltage.RunWorkerAsync();
                            CheckRegelparamU();
                        }
                    }
                    else if (!e.Cancel && !bitPrfgAb)
                    {
                        bitRegelung = true;

                        while (bitRegelung && !e.Cancel)
                        {
                            if (bitAbbruch)
                                e.Cancel = true;

                            CalcSWeiteU(Umess, m_U);

                            if ((m_U - 0.01) < Umess)                     //&& zPhasen == 1)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) - m_SWeite);
                            }
                            else if ((m_U + 0.01) > Umess)               //&& zPhasen == 1)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);
                            }

                            SetObjDuringMeasurement();
                            SetVoltageSV_ACDC(setU);
                           // Thread.Sleep(500);
                            Thread.Sleep(50);
                            bgwMeasureCurrent.RunWorkerAsync();
                            bgwMeasureVoltage.RunWorkerAsync();
                            CheckRegelparamU();
                        }
                    }
                    break;      //case 2
            }
        }

        private void bgwUAC_regeln_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                ReadPosActivForm();

                DRM = F_MessageBox.Show("Abbruch durch Nutzer.", "Bedienerhinweis Abbruch", MessageBoxButtons.OKCancel, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DRM == DialogResult.OK)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
            else if (!(e.Error == null))
            {
                ReadPosActivForm();

                F_MessageBox.Show(Environment.NewLine + Environment.NewLine +
                                  "Error: " + e.Error.Message, "Fehler Backgroundworker (bgwTestCycle_LED70)",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }

            OpenForm = Application.OpenForms["F_MessageBox"];

            if (OpenForm != null)
            {
                OpenForm.Close();                            
            }
        }

        private void bgwIAC_regeln_DoWork(object sender, DoWorkEventArgs e)
        {
            textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(bitRegelung)));
            textBox3.Invoke(new Action(() => textBox3.Text = Convert.ToString(bitupdown_manuell)));
            textBox4.Invoke(new Action(() => textBox4.Text = Convert.ToString(bitSV_CC_Mode)));

            if (bitAbbruch == true)
                e.Cancel = true;

            if (bitRegelung == false && bitupdown_manuell == false)
            {
                if (bitSV_CC_Mode == false && e.Cancel == false)
                {
                    textBox5.Invoke(new Action(() => textBox5.Text = "Not CC_Mode"));
                    //  m_I = m_Imin + ((m_Imax - m_Imin) / 2);

                    if (m_Umax == 0)
                        m_Umax = 10;
                    setU = Convert.ToString(m_Umax * 1.1);         // 10% mehr als Umax einstellen, wegen Spannungsfall über Speiseleitung
                    setI = Convert.ToString(m_I * 1e-3 * 0.5);     // 50% von Imax einstellen, da SV schneller in Begrenzung geht

                    SetVoltageSV_ACDC(setU);
                    SetCurrentSV_ACDC(setI);
                    SetCurrentMaxSV_ACDC();
                    SetFrequencySV_ACDC(m_Freq);
                    SetWaveSV_ACDC(setW);
                    SetOutputSV_ACDC("0");
                }
                else if (bitSV_CC_Mode == true && e.Cancel == false)
                {
                    textBox5.Invoke(new Action(() => textBox5.Text = "CC_Mode"));
                    setI = Convert.ToString(m_I * 1e-3 * 0.8);     // 80% Imin einstellen, SV schwingt sonst über
                    SetCurrentSV_ACDC(setI);
                   // Thread.Sleep(500);                            // Wartezeit, sonst Busfehler Strommessung
                    Thread.Sleep(50);
                    bgwMeasureCurrent.RunWorkerAsync();
                    CheckRegelparamI(Typenbezeichnung);
                }
            }
            else if (bitRegelung == false && bitupdown_manuell == true)
            {
                setI = setI.Replace(".", ",");
                setI = Convert.ToString(Convert.ToDouble(setI) + m_SWeite);
                tbxSWeite.Invoke(new Action(() => tbxSWeite.Text = Convert.ToString(m_SWeite)));
                SetCurrentSV_ACDC(setI);
               // Thread.Sleep(500);
                Thread.Sleep(50);
                bgwMeasureCurrent.RunWorkerAsync();
                bgwMeasureVoltage.RunWorkerAsync();
                SetObjAfterMeasurement();
            }

            while (bitRegelung && bitSV_CC_Mode && e.Cancel == false)
            {
                if (bitAbbruch == true)
                    e.Cancel = true;

                CalcSWeiteIAC(Imess, m_I);

                if ((m_I - 1) < Imess)                    //&& zPhasen == 1)
                {
                    setI = setI.Replace(".", ",");
                    setI = Convert.ToString(((Convert.ToDouble(setI) / 1e-3) - m_SWeite) * 1e-3);
                }
                else if ((m_I + 1) > Imess)                        // && zPhasen == 1)
                {
                    setI = setI.Replace(".", ",");
                    setI = Convert.ToString(((Convert.ToDouble(setI) / 1e-3) + m_SWeite) * 1e-3);
                }

                SetObjDuringMeasurement();
                SetCurrentSV_ACDC(setI);
               // Thread.Sleep(500);
                Thread.Sleep(50);
                bgwMeasureCurrent.RunWorkerAsync();
                CheckRegelparamI(Typenbezeichnung);
            }
        }

        private void bgwIAC_regeln_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                ReadPosActivForm();

                DRM = F_MessageBox.Show("Abbruch durch Nutzer.", "Bedienerhinweis Abbruch", MessageBoxButtons.OKCancel, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DRM == DialogResult.OK)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
            else if (!(e.Error == null))
            {
                ReadPosActivForm();

                F_MessageBox.Show(Environment.NewLine + Environment.NewLine +
                                  "Error: " + e.Error.Message, "Fehler Backgroundworker (bgwTestCycle_LED70)",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }

            OpenForm = Application.OpenForms["F_MessageBox"];

            if (OpenForm != null)
            {
                OpenForm.Close();
            }   
        }

        private void bgwUAC_manuell_DoWork(object sender, DoWorkEventArgs e)
        {
            if (bitAbbruch == true)
                e.Cancel = true;

            if (bitRegelung == false)
            {
                setU = setU.Replace(".", ",");
                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);            
                tbxSWeite.Invoke(new Action(() => tbxSWeite.Text = Convert.ToString(m_SWeite)));
                tbxsetU.Invoke(new Action(() => tbxsetU.Text = setU));
                SetVoltageSV_ACDC(setU);
              //  Thread.Sleep(500);
                Thread.Sleep(50);
                bgwMeasureCurrent.RunWorkerAsync();
                bgwMeasureVoltage.RunWorkerAsync();
                SetObjAfterMeasurement();
            }
        }

        private void bgwUAC_manuell_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                ReadPosActivForm();

                DRM = F_MessageBox.Show("Abbruch durch Nutzer.", "Bedienerhinweis Abbruch", MessageBoxButtons.OKCancel, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DRM == DialogResult.OK)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
            else if (!(e.Error == null))
            {
                ReadPosActivForm();

                F_MessageBox.Show(Environment.NewLine + Environment.NewLine +
                                  "Error: " + e.Error.Message, "Fehler Backgroundworker (bgwU_manuell)",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }

            OpenForm = Application.OpenForms["F_MessageBox"];

            if (OpenForm != null)
            {
                OpenForm.Close();
            }
        }

        private void bgwIAC_manuell_DoWork(object sender, DoWorkEventArgs e)
        {
            if (bitAbbruch == true)
                e.Cancel = true;

            if (bitRegelung == false)
            {
                setI = setI.Replace(".", ",");
                setI = Convert.ToString(Convert.ToDouble(setI) + m_SWeite);
                tbxSWeite.Invoke(new Action(() => tbxSWeite.Text = Convert.ToString(m_SWeite)));
                tbxsetI.Invoke(new Action(() => tbxsetI.Text = setI));
                SetCurrentSV_ACDC(setI);
               // Thread.Sleep(500);
                Thread.Sleep(50);
                bgwMeasureCurrent.RunWorkerAsync();
                bgwMeasureVoltage.RunWorkerAsync();
                SetObjAfterMeasurement();
            }
        }

        private void bgwIAC_manuell_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                ReadPosActivForm();

                DRM = F_MessageBox.Show("Abbruch durch Nutzer.", "Bedienerhinweis Abbruch", MessageBoxButtons.OKCancel, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DRM == DialogResult.OK)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
            else if (!(e.Error == null))
            {
                ReadPosActivForm();

                F_MessageBox.Show(Environment.NewLine + Environment.NewLine +
                                  "Error: " + e.Error.Message, "Fehler Backgroundworker (bgwI_manuell)",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }

            OpenForm = Application.OpenForms["F_MessageBox"];

            if (OpenForm != null)
            {
                OpenForm.Close();
            }
        }

        private void bgwUDC_regeln_DoWork(object sender, DoWorkEventArgs e)
        {
            textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(bitRegelung)));
            textBox3.Invoke(new Action(() => textBox3.Text = Convert.ToString(bitupdown_manuell)));
            textBox4.Invoke(new Action(() => textBox4.Text = Convert.ToString(bitSV_CC_Mode)));

            if (bitAbbruch == true)
                e.Cancel = true;

            switch (zPhasen)
            {
                case 1:

                    if (!bitRegelung && !e.Cancel && !bitPrfgAb)
                    {
                        setU = Convert.ToString(m_U * 1.03);                 // 103% von U einstellen, wegen Spannungsfall über Speiseleitung

                        if (m_Imax == 0) m_Imax = 1500;

                        m_I = m_Imax * 1e-3 * 1.1;                          // 110% von I als Strombegrenzung einstellen
                        setI = Convert.ToString(m_I);

                        SetVoltageSV_ACDC(setU);
                        SetCurrentSV_ACDC(setI);
                        SetCurrentMaxSV_ACDC();
                        SetFrequencySV_ACDC(m_Freq);
                        SetWaveSV_ACDC(setW);
                        SetOutputSV_ACDC("0");
                       // Thread.Sleep(500);
                        Thread.Sleep(50);
                        bgwMeasureVoltage.RunWorkerAsync();
                        CheckRegelparamU();
                    }

                    if (bitRegelung && !e.Cancel && !bitPrfgAb)
                    {
                        while (bitRegelung && !e.Cancel)
                        {
                            if (bitAbbruch)
                                e.Cancel = true;

                            CalcSWeiteU(Umess, m_U);

                            if ((m_U - 0.01) < Umess)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) - m_SWeite);
                            }
                            else if ((m_U + 0.01) > Umess)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);
                            }

                            SetObjDuringMeasurement();
                            SetVoltageSV_ACDC(setU);
                           // Thread.Sleep(500);
                            Thread.Sleep(50);
                            bgwMeasureVoltage.RunWorkerAsync();
                            CheckRegelparamU();
                        }
                    }
                    break;      //case 1

                case 2:

                    if (!e.Cancel && !(m_PrfgAb || m_PrfgCo))
                    {
                        while (bitRegelung && !e.Cancel)
                        {
                            if (bitAbbruch)
                                e.Cancel = true;

                            CalcSWeiteU(Umess, m_U);

                            if ((m_U - 0.01) < Umess)                     //&& zPhasen == 1)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) - m_SWeite);
                            }
                            else if ((m_U + 0.01) > Umess)               //&& zPhasen == 1)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);
                            }

                            SetObjDuringMeasurement();
                            SetVoltageSV_ACDC(setU);
                           // Thread.Sleep(500);
                            Thread.Sleep(50);
                            bgwMeasureVoltage.RunWorkerAsync();
                            CheckRegelparamU();
                        }
                    }
                    else if (!e.Cancel && !bitPrfgAb)
                    {
                        bitRegelung = true;

                        while (bitRegelung && !e.Cancel)
                        {
                            if (bitAbbruch)
                                e.Cancel = true;

                            CalcSWeiteU(Umess, m_U);

                            if ((m_U - 0.01) < Umess)                     //&& zPhasen == 1)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) - m_SWeite);
                            }
                            else if ((m_U + 0.01) > Umess)               //&& zPhasen == 1)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);
                            }

                            SetObjDuringMeasurement();
                            SetVoltageSV_ACDC(setU);
                           // Thread.Sleep(500);
                            Thread.Sleep(50);
                            bgwMeasureCurrent.RunWorkerAsync();
                            bgwMeasureVoltage.RunWorkerAsync();
                            CheckRegelparamU();
                        }
                    }
                    break;      //case 2
            }
        }

        private void bgwUDC_regeln_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                ReadPosActivForm();

                DRM = F_MessageBox.Show("Abbruch durch Nutzer.", "Bedienerhinweis Abbruch", MessageBoxButtons.OKCancel, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DRM == DialogResult.OK)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
            else if (!(e.Error == null))
            {
                ReadPosActivForm();

                F_MessageBox.Show(Environment.NewLine + Environment.NewLine +
                                  "Error: " + e.Error.Message, "Fehler Backgroundworker (bgwUDC_regeln)",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }

            OpenForm = Application.OpenForms["F_MessageBox"];

            if (OpenForm != null)
            {
                OpenForm.Close();
            }
        }

        private void bgwIDC_regeln_DoWork(object sender, DoWorkEventArgs e)
        {
            textBox2.Invoke(new Action(() => textBox2.Text = Convert.ToString(bitRegelung)));
            textBox3.Invoke(new Action(() => textBox3.Text = Convert.ToString(bitupdown_manuell)));
            textBox4.Invoke(new Action(() => textBox4.Text = Convert.ToString(bitSV_CC_Mode)));

            if (bitAbbruch == true)
                e.Cancel = true;

            switch (zPhasen)
            {
                case 1:

                    if (!bitRegelung && !e.Cancel && !bitPrfgAb)
                    {
                        setU = Convert.ToString(m_U);                                        // U einstellen
  
                        m_I = (((m_Imax - m_Imin) / 2) + m_Imin) * 1e-3;
                        setI = Convert.ToString(m_Imax * 1e-3 * 2);                       // Strombegrenzung auf 150%  von Imax einstellen

                        SetVoltageSV_ACDC(setU);
                        SetCurrentSV_ACDC(setI);
                        SetWaveSV_ACDC(setW);
                        SetOutputSV_ACDC("0");
                      //  Thread.Sleep(1000);
                        Thread.Sleep(500);
                        MeasCurrSV_ACDC();
                        CheckRegelparamIDC();
                    }

                    if (bitRegelung && !e.Cancel && !bitPrfgAb)
                    {
                        while (bitRegelung && !e.Cancel)
                        {
                            if (bitAbbruch)
                                e.Cancel = true;

                            CalcSWeiteIDC(Imess, m_I);

                            //if ((m_I - 0.005) < Imess)
                            if (m_I < Imess)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) - m_SWeite);
                            }
                            //else if ((m_I + 0.005) > Imess)
                            else if (m_I > Imess)
                            {
                                setU = setU.Replace(".", ",");
                                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);
                            }

                            SetObjDuringMeasurement();
                            SetVoltageSV_ACDC(setU);
                            Thread.Sleep(500);
                            MeasCurrSV_ACDC();
                            CheckRegelparamIDC();
                        }
                    }
                    break;      //case 1

                case 2:

                    break;
            }
        }

        private void bgwIDC_regeln_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                ReadPosActivForm();

                DRM = F_MessageBox.Show("Abbruch durch Nutzer.", "Bedienerhinweis Abbruch", MessageBoxButtons.OKCancel, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DRM == DialogResult.OK)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
            else if (!(e.Error == null))
            {
                ReadPosActivForm();

                F_MessageBox.Show(Environment.NewLine + Environment.NewLine +
                                  "Error: " + e.Error.Message, "Fehler Backgroundworker (bgwIDC_regeln)",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }

            OpenForm = Application.OpenForms["F_MessageBox"];

            if (OpenForm != null)
            {
                OpenForm.Close();
            }
        }

        private void bgwUDC_manuell_DoWork(object sender, DoWorkEventArgs e)
        {
            if (bitAbbruch == true)
                e.Cancel = true;

            if (bitRegelung == false)
            {
                setU = setU.Replace(".", ",");
                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);
                tbxSWeite.Invoke(new Action(() => tbxSWeite.Text = Convert.ToString(m_SWeite)));
                SetVoltageSV_ACDC(setU);
              //  Thread.Sleep(500);
                Thread.Sleep(50);
                bgwMeasureCurrent.RunWorkerAsync();
                bgwMeasureVoltage.RunWorkerAsync();
                SetObjAfterMeasurement();
            }
        }

        private void bgwUDC_manuell_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                ReadPosActivForm();

                DRM = F_MessageBox.Show("Abbruch durch Nutzer.", "Bedienerhinweis Abbruch", MessageBoxButtons.OKCancel, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DRM == DialogResult.OK)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
            else if (!(e.Error == null))
            {
                ReadPosActivForm();

                F_MessageBox.Show(Environment.NewLine + Environment.NewLine +
                                  "Error: " + e.Error.Message, "Fehler Backgroundworker (bgwUDC_manuell)",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }

            OpenForm = Application.OpenForms["F_MessageBox"];

            if (OpenForm != null)
            {
                OpenForm.Close();
            }
        }

        private void bgwIDC_manuell_DoWork(object sender, DoWorkEventArgs e)
        {
            if (bitAbbruch == true)
                e.Cancel = true;

            if (bitRegelung == false && !(e.Cancel))
            {
                setU = setU.Replace(".", ",");
                setU = Convert.ToString(Convert.ToDouble(setU) + m_SWeite);
                tbxSWeite.Invoke(new Action(() => tbxSWeite.Text = Convert.ToString(m_SWeite)));
                SetVoltageSV_ACDC(setU);
               // Thread.Sleep(500);
                Thread.Sleep(100);
                bgwMeasureCurrent.RunWorkerAsync();
                bgwMeasureVoltage.RunWorkerAsync();
                SetObjAfterMeasurement();
            }
        }

        private void bgwIDC_manuell_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                ReadPosActivForm();

                DRM = F_MessageBox.Show("Abbruch durch Nutzer.", "Bedienerhinweis Abbruch", MessageBoxButtons.OKCancel, new F_MessageBox(),
                                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                if (DRM == DialogResult.OK)
                {
                    bitAbbruch = true;
                    AbortTest();
                }
            }
            else if (!(e.Error == null))
            {
                ReadPosActivForm();

                F_MessageBox.Show(Environment.NewLine + Environment.NewLine +
                                  "Error: " + e.Error.Message, "Fehler Backgroundworker (bgwIDC_manuell)",
                                  MessageBoxButtons.OK, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);

                bitAbbruch = true;
                AbortTest();
            }

            OpenForm = Application.OpenForms["F_MessageBox"];

            if (OpenForm != null)
            {
                OpenForm.Close();
            }
        }

        private double CalcSWeiteU(double Umess, double m_U)
        {
            Udiff = Math.Round(Umess - m_U, 3);
            tbxUdiff.Invoke(new Action(() => tbxUdiff.Text = Convert.ToString(Udiff)));          // Variable Udiff in Textbox anzeigen

            switch (zPhasen)
            {
                case 1:

                    if ((Sweite1[zPrüfschritt] == 0 ) ||
                        (Sweite1[zPrüfschritt] != 0 && Sweite1[zPrüfschritt] > Math.Abs(Udiff)) ||
                        (m_SWeite > Math.Abs(Udiff)))
                    {
                        switch (Manufacturer[2])                                            // Manufacturer SV
                        {
                            case "HBS-ELECTRONIC":

                                if (Math.Abs(Udiff) > 0.05)                             // Auswahl der Schrittweite für Regelung in Abhängigkeit von der Differenz
                                {                                                       // zwischen Vorgabewert und Messwert, Messwert kleiner als Vorgabewert.
                                    if (Math.Abs(Udiff) > 0.1)
                                    {
                                        if (Math.Abs(Udiff) > 0.2)
                                        {
                                            if (Math.Abs(Udiff) > 0.5)
                                            {
                                                if (Math.Abs(Udiff) > 0.7)
                                                {
                                                    m_SWeite = 0.5;
                                                }
                                                else
                                                {
                                                    m_SWeite = 0.3;
                                                }
                                            }
                                            else
                                            {
                                                m_SWeite = 0.1;
                                            }
                                        }
                                        else
                                        {
                                            m_SWeite = 0.05;
                                        }
                                    }
                                    else
                                    {
                                        m_SWeite = 0.01;
                                    }
                                }
                                else
                                {
                                    m_SWeite = 0.005;
                                }
                                break;

                            case "SCHULTZ ELECTRONIC":                                      // Manufacturer SV
                                m_SWeite = Math.Abs(Udiff);
                                break;

                            default:
                                m_SWeite = 0.01;
                                break;
                        }
                    }
                    else if (Sweite1[zPrüfschritt] != 0 && (Sweite1[zPrüfschritt] < Math.Abs(Udiff)))
                    {
                        m_SWeite = Sweite1[zPrüfschritt];
                    }
                    break;

                case 2:

                    if (m_PrfgAb && Sweite2[zPrüfschritt] == 0)
                    // if (bitPrfgAb && Sweite2[zPrüfschritt] == 0)
                    {
                        if (m_SWeite < Math.Abs(Udiff))
                        {
                            m_SWeite = 0.05;
                        }
                        else
                        {
                            m_SWeite = 0.01;
                        }
                    }
                    else if (m_PrfgAb && Sweite2[zPrüfschritt] != 0)
                    //else if (bitPrfgAb && Sweite2[zPrüfschritt] != 0)
                    {
                        if (m_SWeite > Math.Abs(Udiff))
                        {
                            if (Math.Abs(Udiff) > 0.05)
                            {
                                if (Math.Abs(Udiff) > 0.5)
                                {
                                    m_SWeite = 0.2;
                                }
                                else
                                {
                                    m_SWeite = 0.1;
                                }
                            }
                            else
                            {
                                m_SWeite = 0.01;
                            }

                        }
                    }
                    else if (!(m_PrfgAb) && Sweite2[zPrüfschritt] == 0)
                    {
                        m_SWeite = Math.Abs(Udiff);
                    }
                    break;
                default:
                    m_SWeite = 0.01;
                    break;
            }
            
            tbxSWeite.Invoke(new Action(() => tbxSWeite.Text = Convert.ToString(m_SWeite)));

            return m_SWeite;
        }

        private double CalcSWeiteIAC(double Imess, double m_I)
        {
            Idiff = Math.Round(Imess - m_I, 3);
            tbxIdiff.Invoke(new Action(() => tbxIdiff.Text = Convert.ToString(Idiff)));          // Variable Idiff in Textbox anzeigen

            switch (zPhasen)
            {
                case 1:

                    if ((Sweite1[zPrüfschritt] == 0) ||
                        (Sweite1[zPrüfschritt] != 0 && Sweite1[zPrüfschritt] > Math.Abs(Idiff)) ||
                        (m_SWeite > Math.Abs(Idiff)))
                    {
                        switch (Manufacturer[2])                                            // Manufacturer SV
                        {
                            case "HBS-ELECTRONIC":

                                if (Math.Abs(Idiff) > 0.005)                             // Auswahl der Schrittweite für Regelung in Abhängigkeit von der Differenz
                                {                                                       // zwischen Vorgabewert und Messwert, Messwert kleiner als Vorgabewert.
                                    if (Math.Abs(Idiff) > 0.01)
                                    {
                                        if (Math.Abs(Idiff) > 0.02)
                                        {
                                            if (Math.Abs(Idiff) > 0.05)
                                            {
                                                if (Math.Abs(Idiff) > 0.07)
                                                {
                                                    m_SWeite = 0.02;
                                                }
                                                else
                                                {
                                                    m_SWeite = 0.005;
                                                }
                                            }
                                            else
                                            {
                                                m_SWeite = 0.004;
                                            }
                                        }
                                        else
                                        {
                                            m_SWeite = 0.003;
                                        }
                                    }
                                    else
                                    {
                                        m_SWeite = 0.002;
                                    }
                                }
                                else
                                {
                                    m_SWeite = 0.001;
                                }
                                break;

                            case "SCHULTZ ELECTRONIC":                                      // Manufacturer SV
                                m_SWeite = Math.Abs(Idiff);
                                break;

                            default:
                                m_SWeite = 0.001;
                                break;
                        }
                    }
                    break;

                case 2:

                    break;

                default:

                    break;
            }                 

            tbxSWeite.Invoke(new Action(() => tbxSWeite.Text = Convert.ToString(m_SWeite)));

            return m_SWeite;
        }
        
        private double CalcSWeiteIDC(double Imess, double m_I)
        {
            Idiff = Math.Round(Imess - m_I, 3);
            tbxIdiff.Invoke(new Action(() => tbxIdiff.Text = Convert.ToString(Idiff)));          // Variable Idiff in Textbox anzeigen

            switch (zPhasen)
            {
                case 1:

                    if ((Sweite1[zPrüfschritt] == 0) ||
                        (Sweite1[zPrüfschritt] != 0 && Sweite1[zPrüfschritt] > Math.Abs(Idiff)) ||
                        (m_SWeite > Math.Abs(Idiff)))
                    {
                        switch (Manufacturer[2])                                            // Manufacturer SV
                        {
                            case "HBS-ELECTRONIC":

                                if (Math.Abs(Idiff) > 0.001)                           // Auswahl der Schrittweite für Regelung in Abhängigkeit von der Differenz
                                {                                                       // zwischen Vorgabewert und Messwert, Messwert kleiner als Vorgabewert.
                                    if (Math.Abs(Idiff) > 0.01)
                                    {
                                        if (Math.Abs(Idiff) > 0.02)
                                        {
                                            if (Math.Abs(Idiff) > 0.05)
                                            {
                                                if (Math.Abs(Idiff) > 0.07)
                                                {
                                                    m_SWeite = 0.5;
                                                }
                                                else
                                                {
                                                    m_SWeite = 0.3;
                                                }
                                            }
                                            else
                                            {
                                                m_SWeite = 0.1;
                                            }
                                        }
                                        else
                                        {
                                            m_SWeite = 0.05;
                                        }
                                    }
                                    else
                                    {
                                        m_SWeite = 0.01;
                                    }
                                }
                                else
                                {
                                    m_SWeite = 0.001;
                                }
                                break;

                            case "SCHULTZ ELECTRONIC":                                      // Manufacturer SV
                                 if (Math.Abs(Idiff) > 0.005)                           // Auswahl der Schrittweite für Regelung in Abhängigkeit von der Differenz
                                {                                                       // zwischen Vorgabewert und Messwert, Messwert kleiner als Vorgabewert.
                                    if (Math.Abs(Idiff) > 0.01)
                                    {
                                        if (Math.Abs(Idiff) > 0.02)
                                        {
                                            if (Math.Abs(Idiff) > 0.05)
                                            {
                                                if (Math.Abs(Idiff) > 0.08)
                                                {
                                                    m_SWeite = 0.3;
                                                }
                                                else
                                                {
                                                    m_SWeite = 0.15;
                                                }
                                            }
                                            else
                                            {
                                                m_SWeite = 0.1;
                                            }
                                        }
                                        else
                                        {
                                            m_SWeite = 0.05;
                                        }
                                    }
                                    else
                                    {
                                        m_SWeite = 0.02;
                                    }
                                }
                                else
                                {
                                    m_SWeite = 0.01;
                                }
                                break;

                            default:
                                m_SWeite = 0.01;
                                break;
                        }
                    }
                    else if (Sweite1[zPrüfschritt] != 0 && (Sweite1[zPrüfschritt] < Math.Abs(Udiff)))
                    {
                        m_SWeite = Sweite1[zPrüfschritt];
                    }
                    break;

                case 2:
                    
                    break;
            }

            tbxSWeite.Invoke(new Action(() => tbxSWeite.Text = Convert.ToString(m_SWeite)));

            return m_SWeite;
        }

        public static void AppraisalSwitchOFF_manuell(bool bitAbschaltung_visuell)
        {
            getbitAbschaltung_visuell = bitAbschaltung_visuell;
        }

        private void CheckRegelparamU()
        {
            do
            { }
            while (bgwMeasureVoltage.IsBusy || bgwMeasureCurrent.IsBusy);

            switch (zPhasen)
            {
                case 1:

                    switch (Manufacturer[2])
                    {
                        case "HBS-ELECTRONIC":                                          // Manufacturer SV

                            if ((m_U - 0.07) > Umess || Umess > (m_U + 0.07))
                            {
                                bitRegelung = true;
                                tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                            }
                            else
                            {
                                bitRegelung = false;
                                tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                bgwMeasureCurrent.RunWorkerAsync();
                                SetObjAfterMeasurement();
                            }
                            break;

                        case "SCHULTZ ELECTRONIC":                                      // Manufacturer SV

                            if ((m_U - 0.01) > Umess || Umess > (m_U + 0.01))
                            {
                                bitRegelung = true;
                                tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                            }
                            else
                            {
                                bitRegelung = false;
                                tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                bgwMeasureCurrent.RunWorkerAsync();
                                SetObjAfterMeasurement();
                            }
                            break;
                    }
                    break;          //case1

                case 2:

                    if (m_PrfgAb)
                    {
                        switch (Manufacturer[2])
                        {
                            case "HBS-ELECTRONIC":                                          // Manufacturer SV

                                if (Imess > m_Imax)
                                {
                                    if ((m_U - 0.01) > Umess || Umess < (m_U + 0.05))
                                    {
                                        bitRegelung = true;
                                        tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                    }
                                    else
                                    {
                                        bitRegelung = false;
                                        tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                        bitAbschaltung = false;
                                        SetObjAfterMeasurement();
                                    }
                                }
                                else
                                {
                                    bitRegelung = false;
                                    tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                    bitAbschaltung = true;
                                    SetObjAfterMeasurement();
                                }
                                break;

                            case "SCHULTZ ELECTRONIC":                                      // Manufacturer SV

                                if (Imess > m_Imax)
                                {
                                    if ((m_U - 0.01) > Umess || Umess < (m_U + 0.01))
                                    {
                                        bitRegelung = true;
                                        tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                    }
                                    else
                                    {
                                        bitRegelung = false;
                                        tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                        bitAbschaltung = false;
                                        SetObjAfterMeasurement();
                                    }
                                }
                                else
                                {
                                    bitRegelung = false;
                                    tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                    bitAbschaltung = true;
                                    SetObjAfterMeasurement();
                                }
                                break;
                        }
                    }
                    else if (m_PrfgU || m_PrfgI)
                    {
                        switch (Manufacturer[2])
                        {
                            case "HBS-ELECTRONIC":                                          // Manufacturer SV

                                if ((m_U - 0.07) > Umess || Umess > (m_U + 0.07))
                                {
                                    bitRegelung = true;
                                    tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                }
                                else
                                {
                                    bitRegelung = false;
                                    tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                    bgwMeasureCurrent.RunWorkerAsync();
                                    SetObjAfterMeasurement();
                                }
                                break;

                            case "SCHULTZ ELECTRONIC":                                      // Manufacturer SV

                                if ((m_U - 0.005) > Umess || Umess > (m_U + 0.005))
                                {
                                    bitRegelung = true;
                                    tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                }
                                else
                                {
                                    bitRegelung = false;
                                    tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                                    bgwMeasureCurrent.RunWorkerAsync();
                                    SetObjAfterMeasurement();
                                }
                                break;
                        }

                    }
                    break;      //case2                    
            }
        }

        private void CheckRegelparamI(string Typenbezeichnung)
        {
            do
            { }
            while (bgwMeasureVoltage.IsBusy || bgwMeasureCurrent.IsBusy);

            if (zPhasen == 1)
            {
                switch (Manufacturer[2])
                {
                    case "HBS-ELECTRONIC":                                          // Manufacturer SV

                        if ((m_Imin - 1) > Imess || Imess > (m_Imax + 1))                 // 1mA über Imin und unter Imax einregeln
                        {
                            bitRegelung = true;
                            tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                        }
                        else
                        {
                            bitRegelung = false;
                            tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                            bgwMeasureVoltage.RunWorkerAsync();
                            SetObjAfterMeasurement();
                        }
                        break;

                    case "SCHULTZ ELECTRONIC":                                      // Manufacturer SV

                        if ((m_I - 1) > Imess || Imess > (m_I + 1))
                        {
                            bitRegelung = true;
                            tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                        }
                        else
                        {
                            bitRegelung = false;
                            tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                            bgwMeasureVoltage.RunWorkerAsync();
                            SetObjAfterMeasurement();
                        }
                        break;

                    default:
                        break;
                }
            }  
        }

        private void CheckRegelparamIDC()
        {
            do
            { }
            while (bgwMeasCurrSV_ACDC.IsBusy);

            if (zPhasen == 1)
            {
                switch (Manufacturer[2])
                {
                    case "HBS-ELECTRONIC":                                          // Manufacturer SV

                        if ((m_I - 0.001) > Imess || Imess > (m_I + 0.001))                 // 1mA über Imin und unter Imax einregeln
                        {
                            bitRegelung = true;
                            tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                        }
                        else
                        {
                            bitRegelung = false;
                            tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                            bgwMeasureCurrent.RunWorkerAsync();
                            bgwMeasureVoltage.RunWorkerAsync();
                            SetObjAfterMeasurement();
                        }
                        break;

                    case "SCHULTZ ELECTRONIC":                                      // Manufacturer SV

                        if ((m_I - 0.002) > Imess || Imess > (m_I + 0.002))
                        {
                            bitRegelung = true;
                            tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                        }
                        else
                        {
                            bitRegelung = false;
                            tbxRegelung.Invoke(new Action(() => tbxRegelung.Text = Convert.ToString(bitRegelung)));
                            bgwMeasureCurrent.RunWorkerAsync();
                            bgwMeasureVoltage.RunWorkerAsync();
                            SetObjAfterMeasurement();
                        }
                        break;

                    default:
                        break;
                }
            }
        }

        private void AppraisalVisualTest()
        {
            if (m_PrfgVi)
            {
                if (cboxiO.Enabled || cboxniO.Enabled)
                {
                    if (cboxiO.Checked)
                    {
                        tbxVisuellePrüfung.BackColor = Color.LightGreen;
                        cboxniO.Enabled = false;
                        btnStartWeiterSpeichern.Enabled = true;
                        ErgPrfgVi = "JA";
                    }
                    else if (cboxniO.Checked)
                    {
                        tbxVisuellePrüfung.BackColor = Color.LightPink;
                        cboxiO.Enabled = false;
                        btnStartWeiterSpeichern.Enabled = true;
                        ErgPrfgVi = "NEIN";
                    }
                    else
                    {
                        tbxVisuellePrüfung.BackColor = SystemColors.Window;
                        cboxiO.Enabled = true;
                        cboxniO.Enabled = true;
                        btnStartWeiterSpeichern.Enabled = false;
                    }
                }
            }
            else
            {
                ErgPrfgVi = "n.b.";
            }
        }

        private void AppraisalAutoTests()
        {
            if (m_PrfgU)
            {
                if (m_Umin <= Umess && Umess <= m_Umax)
                    ErgPrfgU = "JA";
                else
                    ErgPrfgU = "NEIN";
            }
            else
            {
                ErgPrfgU = "n.b.";
            }

            if (m_PrfgI)
            {
                if (m_Imin <= Imess && Imess <= m_Imax ||
                   (m_Imin == 0 && Imess <= m_Imax) ||
                   (m_Imax == 0 && Imess >= m_Imin))
                    ErgPrfgI = "JA";
                else
                    ErgPrfgI = "NEIN";
            }
            else
            {
                ErgPrfgI = "n.b.";
            }

            if (m_PrfgAb)
            {
                if (Imess <= m_Imax)
                    ErgPrfgI = "JA";
                else
                    ErgPrfgI = "NEIN";
            }
            else
            {
                ErgPrfgAb = "n.b.";
            }

            if (m_PrfgCo)
            {
                if ((m_Umin <= Umess && Umess <= m_Umax && m_Imin <= Imess && Imess <= m_Imax) ||
                    (m_Umin <= Umess && Umess <= m_Umax && m_Imin == 0 && Imess <= m_Imax) ||
                    (m_Umin <= Umess && Umess <= m_Umax && m_Imax == 0 && Imess >= m_Imin) ||
                    (m_Umin <= Umess && Umess <= m_Umax && m_Imin == 0 && m_Imax == 0))
                {
                    ErgPrfgCo = "JA";
                }
                else
                {
                    ErgPrfgCo = "NEIN";
                }
            }
            else
            {
                ErgPrfgCo = "n.b.";
            }

            if (!m_PrfgVi)
                ErgPrfgVi = "n.b.";
        }

        private void SetDateTime()
        {
            Datum_Uhrzeit = DateTime.Now;
        }

        public static void GetErgProtMatrix(double[,] Erg)
        {
            ErgLED136Mat = new double[4, 3];
            ErgAbweichxy = new double[2];

            ErgLED136Mat = Erg;
        }
        
        private void SaveResults()
        {
            if (zPrüfschritt == 1)                                         // Initialisierung Arrays Messergebnisse
            {
                ErgUI = new double[13, 3];
                ErgPrfg = new string[13, 6];                               // Ergebnisse PrfgU, PrfgI, PrfgCo, PrfgAb, PrfgVi
            }

            ErgUI[zPrüfschritt, 1] = Convert.ToDouble(Umess);
            ErgUI[zPrüfschritt, 2] = Convert.ToDouble(Imess);
            ErgPrfg[zPrüfschritt, 1] = ErgPrfgU;
            ErgPrfg[zPrüfschritt, 2] = ErgPrfgI;
            ErgPrfg[zPrüfschritt, 3] = ErgPrfgCo;
            ErgPrfg[zPrüfschritt, 4] = ErgPrfgAb;
            ErgPrfg[zPrüfschritt, 5] = ErgPrfgVi;
        }

        private void UpdateResultsinDB()
        {
            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb;" +
                                   "Jet OLEDB:Database Password=Goniometer";
            cmd.Connection = con;

            if (Typenbezeichnung != "LED 136 Matrix-BG")
            {
                try
                {
                    con.Open();

                    cmd.CommandText = "UPDATE Messergebnisse SET " +

                                      "Limes_TestID = '" + LimesTestID + "', " +
                                      "Sachnummer = '" + Sachnummer + "', " +
                                      "Serialnummer = '" + Serialnummer + "', " +
                                      "Datum_Uhrzeit = '" + Datum_Uhrzeit + "', " +
                                      "PruefTemp = '" + PruefTemp + "', " +
                                      "ID_Parametersatz = '" + ID_Parametersatz + "', " +
                                      "ID_Messequipment = '" + ID_Messequipment + "', " +
                                      "ID_Prüfer = '" + ID_Prüfer + "', " +

                                      "U1 = '" + ErgUI[1, 1] + "', I1 = '" + ErgUI[1, 2] + "', " +
                                      "PrfgU1 = '" + ErgPrfg[1, 1] + "', PrfgI1 = '" + ErgPrfg[1, 2] + "', PrfgCo1 = '" + ErgPrfg[1, 3] + "', " +
                                      "PrfgAb1 = '" + ErgPrfg[1, 4] + "', PrfgVi1 = '" + ErgPrfg[1, 5] + "', " +

                                      "U2 = '" + ErgUI[2, 1] + "', I2 = '" + ErgUI[2, 2] + "', " +
                                      "PrfgU2 = '" + ErgPrfg[2, 1] + "', PrfgI2 = '" + ErgPrfg[2, 2] + "', PrfgCo2 = '" + ErgPrfg[2, 3] + "', " +
                                      "PrfgAb2 = '" + ErgPrfg[2, 4] + "', PrfgVi2 = '" + ErgPrfg[2, 5] + "', " +

                                      "U3 = '" + ErgUI[3, 1] + "', I3 = '" + ErgUI[3, 2] + "', " +
                                      "PrfgU3 = '" + ErgPrfg[3, 1] + "', PrfgI3 = '" + ErgPrfg[3, 2] + "', PrfgCo3 = '" + ErgPrfg[3, 3] + "', " +
                                      "PrfgAb3 = '" + ErgPrfg[3, 4] + "', PrfgVi3 = '" + ErgPrfg[3, 5] + "', " +

                                      "U4 = '" + ErgUI[4, 1] + "', I4 = '" + ErgUI[4, 2] + "', " +
                                      "PrfgU4 = '" + ErgPrfg[4, 1] + "', PrfgI4 = '" + ErgPrfg[4, 2] + "', PrfgCo4 = '" + ErgPrfg[4, 3] + "', " +
                                      "PrfgAb4 = '" + ErgPrfg[4, 4] + "', PrfgVi4 = '" + ErgPrfg[4, 5] + "', " +

                                      "U5 = '" + ErgUI[5, 1] + "', I5 = '" + ErgUI[5, 2] + "', " +
                                      "PrfgU5 = '" + ErgPrfg[5, 1] + "', PrfgI5 = '" + ErgPrfg[5, 2] + "', PrfgCo5 = '" + ErgPrfg[5, 3] + "', " +
                                      "PrfgAb5 = '" + ErgPrfg[5, 4] + "', PrfgVi5 = '" + ErgPrfg[5, 5] + "', " +

                                      "U6 = '" + ErgUI[6, 1] + "', I6 = '" + ErgUI[6, 2] + "', " +
                                      "PrfgU6 = '" + ErgPrfg[6, 1] + "', PrfgI6 = '" + ErgPrfg[6, 2] + "', PrfgCo6 = '" + ErgPrfg[6, 3] + "', " +
                                      "PrfgAb6 = '" + ErgPrfg[6, 4] + "', PrfgVi6 = '" + ErgPrfg[6, 5] + "', " +

                                      "U7 = '" + ErgUI[7, 1] + "', I7 = '" + ErgUI[7, 2] + "', " +
                                      "PrfgU7 = '" + ErgPrfg[7, 1] + "', PrfgI7 = '" + ErgPrfg[7, 2] + "', PrfgCo7 = '" + ErgPrfg[7, 3] + "', " +
                                      "PrfgAb7 = '" + ErgPrfg[7, 4] + "', PrfgVi7 = '" + ErgPrfg[7, 5] + "', " +

                                      "U8 = '" + ErgUI[8, 1] + "', I8 = '" + ErgUI[8, 2] + "', " +
                                      "PrfgU8 = '" + ErgPrfg[8, 1] + "', PrfgI8 = '" + ErgPrfg[8, 2] + "', PrfgCo8 = '" + ErgPrfg[8, 3] + "', " +
                                      "PrfgAb8 = '" + ErgPrfg[8, 4] + "', PrfgVi8 = '" + ErgPrfg[8, 5] + "', " +

                                      "U9 = '" + ErgUI[9, 1] + "', I9 = '" + ErgUI[9, 2] + "', " +
                                      "PrfgU9 = '" + ErgPrfg[9, 1] + "', PrfgI9 = '" + ErgPrfg[9, 2] + "', PrfgCo9 = '" + ErgPrfg[9, 3] + "', " +
                                      "PrfgAb9 = '" + ErgPrfg[9, 4] + "', PrfgVi9 = '" + ErgPrfg[9, 5] + "', " +

                                      "U10 = '" + ErgUI[10, 1] + "', I10 = '" + ErgUI[10, 2] + "', " +
                                      "PrfgU10 = '" + ErgPrfg[10, 1] + "', PrfgI10 = '" + ErgPrfg[10, 2] + "', PrfgCo10 = '" + ErgPrfg[10, 3] + "', " +
                                      "PrfgAb10 = '" + ErgPrfg[10, 4] + "', PrfgVi10 = '" + ErgPrfg[10, 5] + "', " +

                                      "U11 = '" + ErgUI[11, 1] + "', I11 = '" + ErgUI[11, 2] + "', " +
                                      "PrfgU11 = '" + ErgPrfg[11, 1] + "', PrfgI11 = '" + ErgPrfg[11, 2] + "', PrfgCo11 = '" + ErgPrfg[11, 3] + "', " +
                                      "PrfgAb11 = '" + ErgPrfg[11, 4] + "', PrfgVi11 = '" + ErgPrfg[11, 5] + "', " +

                                      "U12 = '" + ErgUI[12, 1] + "', I12 = '" + ErgUI[12, 2] + "', " +
                                      "PrfgU12 = '" + ErgPrfg[12, 1] + "', PrfgI12 = '" + ErgPrfg[12, 2] + "', PrfgCo12 = '" + ErgPrfg[12, 3] + "', " +
                                      "PrfgAb12 = '" + ErgPrfg[12, 4] + "', PrfgVi12 = '" + ErgPrfg[12, 5] + "' " +

                                      " WHERE Serialnummer = '" + Serialnummer + "'";

                    cmd.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Abspeichern der Messergebnisse in DB ist fehlgeschlagen.\n\n" + ex.Message, "Fehler Messergebnisse speichen (UpdateResultsinDB)");
                }

                con.Close();
            }

            if (Typenbezeichnung == "LED 136 Matrix-BG")
            {
                try
                {
                    con.Open();

                    cmd.CommandText = "UPDATE ErgLED136MatrixBG SET" +

                                      "Sachnummer =  '" + Sachnummer + "', " +
                                      "Serialnummer = '" + Serialnummer + "', " +
                                      "Datum_Uhrzeit = '" + Datum_Uhrzeit + "', " +
                                      "ID_Prüfer = '" + ID_Prüfer + "', " +
                                      "difhPosL1 = '" + ErgLED136Mat[0, 0] + "', n08PosL1 = '" + ErgLED136Mat[0, 1] + "', n05PosL1 = '" + ErgLED136Mat[0, 2] + "', " +
                                      "difhPosL2 = '" + ErgLED136Mat[1, 0] + "', n08PosL2 = '" + ErgLED136Mat[1, 1] + "', n05PosL2 = '" + ErgLED136Mat[1, 2] + "', " +
                                      "difhPosL3 = '" + ErgLED136Mat[2, 0] + "', n08PosL3 = '" + ErgLED136Mat[2, 1] + "', n05PosL3 = '" + ErgLED136Mat[2, 2] + "', " +
                                      "AbweichX = '" + ErgLED136Mat[3, 0] + "', AbweichY = '" + ErgLED136Mat[3, 1] + "' " +

                                      " WHERE Serialnummer = '" + Serialnummer + "'";

                    cmd.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Abspeichern der Messergebnisse in DB ist fehlgeschlagen.\n\n" + ex.Message, "Fehler Messergebnisse speichen (SaveResultsinDB)");
                }

                con.Close();
            }
        }

        private void SaveResultsinDB()
        {
            OleDbConnection con = new OleDbConnection();

            OleDbCommand cmd = new OleDbCommand();
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   "Data Source=D:\\Programme\\LMT\\Data\\Testergebnisse\\Testergebnisse_Signalgeber_elektrisch.mdb;" +
                                   "Jet OLEDB:Database Password=Goniometer";
            cmd.Connection = con;

            if (Typenbezeichnung != "LED 136 Matrix-BG")
            {

                try
                {
                    con.Open();

                    cmd.CommandText = "insert into Messergebnisse" +
                                      "(Limes_TestID, Sachnummer, Serialnummer, Datum_Uhrzeit, PruefTemp, ID_Parametersatz, ID_Messequipment, ID_Prüfer," +

                                         "U1, I1, PrfgU1, PrfgI1, PrfgCo1, PrfgAb1, PrfgVi1, U2, I2, PrfgU2, PrfgI2, PrfgCo2, PrfgAb2, PrfgVi2, U3, I3, PrfgU3, PrfgI3, PrfgCo3, PrfgAb3, PrfgVi3," +
                                         "U4, I4, PrfgU4, PrfgI4, PrfgCo4, PrfgAb4, PrfgVi4, U5, I5, PrfgU5, PrfgI5, PrfgCo5, PrfgAb5, PrfgVi5, U6, I6, PrfgU6, PrfgI6, PrfgCo6, PrfgAb6, PrfgVi6," +
                                         "U7, I7, PrfgU7, PrfgI7, PrfgCo7, PrfgAb7, PrfgVi7, U8, I8, PrfgU8, PrfgI8, PrfgCo8, PrfgAb8, PrfgVi8, U9, I9, PrfgU9, PrfgI9, PrfgCo9, PrfgAb9, PrfgVi9," +
                                         "U10, I10, PrfgU10, PrfgI10, PrfgCo10, PrfgAb10, PrfgVi10, U11, I11, PrfgU11, PrfgI11, PrfgCo11, PrfgAb11, PrfgVi11, U12, I12, PrfgU12, PrfgI12, PrfgCo12, PrfgAb12, PrfgVi12) values( '" +

                                         LimesTestID + "', '" + Sachnummer + "', '" + Serialnummer + "', '" + Datum_Uhrzeit + "', '" + PruefTemp + "', '" + ID_Parametersatz + "', '" + ID_Messequipment + "', '" + ID_Prüfer + "', '" +

                                         ErgUI[1, 1] + "', '" + ErgUI[1, 2] + "', '" + ErgPrfg[1, 1] + "', '" + ErgPrfg[1, 2] + "', '" + ErgPrfg[1, 3] + "', '" + ErgPrfg[1, 4] + "', '" + ErgPrfg[1, 5] + "', '" +
                                         ErgUI[2, 1] + "', '" + ErgUI[2, 2] + "', '" + ErgPrfg[2, 1] + "', '" + ErgPrfg[2, 2] + "', '" + ErgPrfg[2, 3] + "', '" + ErgPrfg[2, 4] + "', '" + ErgPrfg[2, 5] + "', '" +
                                         ErgUI[3, 1] + "', '" + ErgUI[3, 2] + "', '" + ErgPrfg[3, 1] + "', '" + ErgPrfg[3, 2] + "', '" + ErgPrfg[3, 3] + "', '" + ErgPrfg[3, 4] + "', '" + ErgPrfg[3, 5] + "', '" +
                                         ErgUI[4, 1] + "', '" + ErgUI[4, 2] + "', '" + ErgPrfg[4, 1] + "', '" + ErgPrfg[4, 2] + "', '" + ErgPrfg[4, 3] + "', '" + ErgPrfg[4, 4] + "', '" + ErgPrfg[4, 5] + "', '" +
                                         ErgUI[5, 1] + "', '" + ErgUI[5, 2] + "', '" + ErgPrfg[5, 1] + "', '" + ErgPrfg[5, 2] + "', '" + ErgPrfg[5, 3] + "', '" + ErgPrfg[5, 4] + "', '" + ErgPrfg[5, 5] + "', '" +
                                         ErgUI[6, 1] + "', '" + ErgUI[6, 2] + "', '" + ErgPrfg[6, 1] + "', '" + ErgPrfg[6, 2] + "', '" + ErgPrfg[6, 3] + "', '" + ErgPrfg[6, 4] + "', '" + ErgPrfg[6, 5] + "', '" +
                                         ErgUI[7, 1] + "', '" + ErgUI[7, 2] + "', '" + ErgPrfg[7, 1] + "', '" + ErgPrfg[7, 2] + "', '" + ErgPrfg[7, 3] + "', '" + ErgPrfg[7, 4] + "', '" + ErgPrfg[7, 5] + "', '" +
                                         ErgUI[8, 1] + "', '" + ErgUI[8, 2] + "', '" + ErgPrfg[8, 1] + "', '" + ErgPrfg[8, 2] + "', '" + ErgPrfg[8, 3] + "', '" + ErgPrfg[8, 4] + "', '" + ErgPrfg[8, 5] + "', '" +
                                         ErgUI[9, 1] + "', '" + ErgUI[9, 2] + "', '" + ErgPrfg[9, 1] + "', '" + ErgPrfg[9, 2] + "', '" + ErgPrfg[9, 3] + "', '" + ErgPrfg[9, 4] + "', '" + ErgPrfg[9, 5] + "', '" +
                                         ErgUI[10, 1] + "', '" + ErgUI[10, 2] + "', '" + ErgPrfg[10, 1] + "', '" + ErgPrfg[10, 2] + "', '" + ErgPrfg[10, 3] + "', '" + ErgPrfg[10, 4] + "', '" + ErgPrfg[10, 5] + "', '" +
                                         ErgUI[11, 1] + "', '" + ErgUI[11, 2] + "', '" + ErgPrfg[11, 1] + "', '" + ErgPrfg[11, 2] + "', '" + ErgPrfg[11, 3] + "', '" + ErgPrfg[11, 4] + "', '" + ErgPrfg[11, 5] + "', '" +
                                         ErgUI[12, 1] + "', '" + ErgUI[12, 2] + "', '" + ErgPrfg[12, 1] + "', '" + ErgPrfg[12, 2] + "', '" + ErgPrfg[12, 3] + "', '" + ErgPrfg[12, 4] + "', '" + ErgPrfg[12, 5] + "')";

                    cmd.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Abspeichern der Messergebnisse in DB ist fehlgeschlagen.\n\n" + ex.Message, "Fehler Messergebnisse speichen (SaveResultsinDB)");
                }

                con.Close();
            }

            if (Typenbezeichnung == "LED 136 Matrix-BG")
            {
                try
                {
                    con.Open();

                    cmd.CommandText = "insert into ErgLED136MatrixBG" +
                                      "(Sachnummer, Serialnummer, Datum_Uhrzeit, ID_Prüfer," +
                                      "difhPosL1, n08PosL1, n05PosL1," +
                                      "difhPosL2, n08PosL2, n05PosL2," +
                                      "difhPosL3, n08PosL3, n05PosL3," +
                                      "AbweichX, AbweichY) values( '" +

                                      Sachnummer + "', '" + Serialnummer + "', '" + Datum_Uhrzeit + "', '" + ID_Prüfer + "', '" +
                                      ErgLED136Mat[0, 0] + "', '" + ErgLED136Mat[0, 1] + "', '" + ErgLED136Mat[0, 2] + "', '" +
                                      ErgLED136Mat[1, 0] + "', '" + ErgLED136Mat[1, 1] + "', '" + ErgLED136Mat[1, 2] + "', '" +
                                      ErgLED136Mat[2, 0] + "', '" + ErgLED136Mat[2, 1] + "', '" + ErgLED136Mat[2, 2] + "', '" +
                                      ErgLED136Mat[3, 0] + "', '" + ErgLED136Mat[3, 1] + "')";

                    cmd.ExecuteNonQuery();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Abspeichern der Messergebnisse in DB ist fehlgeschlagen.\n\n" + ex.Message, "Fehler Messergebnisse speichen (SaveResultsinDB)");
                }

                con.Close();
            }
        }

        private void AbortTest()
        {
            if (bitAbbruch)
            {
                bitStatusPrüfung = false;

                SetOutputSV_ACDC("1");
                ResetOutDigIO();
                ResetVar();
                ResetObj();                
                LoadParameterInObjects();
            }
        }

        public void btnTest_Click(object sender, EventArgs e)
        {
            //F_Abschaltung_Signalgeber F_Abschaltung_Signalgeber = new F_Abschaltung_Signalgeber();
            //F_Abschaltung_Signalgeber.ShowDialog();

            // F_Test test = new F_Test();
            // test.ShowDialog();

            //Form1Width = F_Abschaltung_Signalgeber.ActiveForm.Bounds.Width;
            //Form1Height = F_Abschaltung_Signalgeber.ActiveForm.Bounds.Height;
            //posForm1x = this.Location.X;
            //posForm1y = this.Location.Y;

            //F_LED_Signalgeber_ReadPos();

            //F_LED_136_Matrix F_LED_136_Matrix = new F_LED_136_Matrix();
            //F_LED_136_Matrix.Serialnummer = Serialnummer;
            //F_LED_136_Matrix.Auftragsnummer = Auftragsnummer;
            //F_LED_136_Matrix.ID_Prüfer = ID_Prüfer;
            //F_LED_136_Matrix.Sachnummer = Sachnummer;
            //F_LED_136_Matrix.Form1Width = Form1Width;
            //F_LED_136_Matrix.Form1Height = Form1Height;
            //F_LED_136_Matrix.posForm1x = posForm1x;
            //F_LED_136_Matrix.posForm1y = posForm1y;
            //F_LED_136_Matrix.Show();



            //aTimer = new System.Timers.Timer(1000);
            //aTimer.Elapsed += new System.Timers.ElapsedEventHandler(WaitForCC_Mode);
            //aTimer.Interval = 1000;
            //aTimer.Enabled = false;

            //aTimer.Start();
            //backgroundWorker1.RunWorkerAsync();

            //F_LED_Signalgeber_ReadPos();

            //DRMC = F_MessageBox.Show("Prüfschritt " + zPrüfschritt + " von " + aPrüfschritte + Environment.NewLine +
            //                                    "Dies ist ein Test!!!!!!!" + Environment.NewLine +
            //                                    "Prüfschritt " + zPrüfschritt + " von " + aPrüfschritte + Environment.NewLine +
            //                                    "Dies ist ein Test!!!!!!!", "Bedienerhinweis",
            //                                    MessageBoxButtons.OKCancel, new F_MessageBox(), posForm1x, posForm1y, Form1Width, Form1Height);


            //MessageBox.Show("DR: " + DRM);

            //DialogResult DR = F_Message_CC_Mode.Show("Wait for CC-Mode", 15, new F_Message_CC_Mode(), posForm1x, posForm1y, Form1Width, Form1Height);

            //MessageBox.Show("DR: " + DR);


            //DRM = F_Message_CC_Mode.Show("Wait for CC-Mode", 12, new F_Message_CC_Mode(), posForm1x, posForm1y, Form1Width, Form1Height);

            //DRM = F_MessageBox_Corona_LED_136_Matrix.Show(new F_MessageBox_Corona_LED_136_Matrix(), posForm1x, posForm1y, Form1Width, Form1Height);

            //MessageBox.Show("DR: " + DRM);


            //ReadPosActivForm();

            ////DRM = F_LED_136_Matrix_Print_Prot.Show(Sachnummer, Auftragsnummer, Serialnummer, ID_Prüfer, MessageBoxButtons.OKCancel, new F_LED_136_Matrix_Print_Prot(), posForm1x, posForm1y, Form1Width, Form1Height);


            ////MessageBox.Show("DR: " + DRM);

            //DRM = F_MessageBox.Show("Die erforderliche Spannung zum Start der Abschaltprüfung, von " + m_U + " V ist eingeregelt." +
            //                             Environment.NewLine +
            //                             "Bestätigen Sie mit OK, um die Abschaltprüfung zu starten, Cancel für Abbruch.",
            //                             "Prüfung Abschaltung " + Typenbezeichnung,
            //                             MessageBoxButtons.OKCancel, new F_MessageBox(), posActFormx, posActFormy, ActFormWidth, ActFormHeight);


            ReadPosActivForm();

            //MboxText = "Es wurde die Abschaltung des Signalgebers " + Typenbezeichnung + " detektiert." +
            //                       "Bitte kontrollieren Sie, ob die Abschaltung des Signalgebers erfolgt ist." +
            //                       Environment.NewLine +
            //                       "Hat der Signalgebers abgeschaltet?";

            //MboxTitle = "Kontrolle Abschaltung Signalgeber";

            //DRM = F_MessageBox.Show(MboxText, MboxTitle, MessageBoxButtons.YesNo, new F_MessageBox(),
            //                        posActFormx, posActFormy, ActFormWidth, ActFormHeight);


            //Sachnummer = "999999999999999";
            //Auftragsnummer = "11111111111111111";
            //Serialnummer = "2222222222";
            //ID_Prüfer = "D999";

            //DRMPP = F_LED_136_Matrix_Print_Prot.Show(Sachnummer, Auftragsnummer, Serialnummer, ID_Prüfer,
            //                                         MessageBoxButtons.OKCancel, new F_LED_136_Matrix_Print_Prot(),
            //                                         posActFormx, posActFormy, ActFormWidth, ActFormHeight);


            /*      aPrüfschritte = 5;

                  U1 = new double[aPrüfschritte];

                  U1[0] = 1.5;
                  U1[1] = 1.9;
                  U1[2] = 2.7;
                  U1[3] = 5.5;
                  U1[4] = 9.5;
      */
       //     F_Repeat F_Repeat = new F_Repeat(zPrüfschritt, U1, I1min, I1max, f, VP, Lichtstärke, Farbort, posActFormx, posActFormy, ActFormWidth, ActFormHeight);
       //     F_Repeat.ShowDialog();
        

            

    
        }
               
    }        
 }
