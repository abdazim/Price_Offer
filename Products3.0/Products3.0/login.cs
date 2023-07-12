using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.NetworkInformation;    //GetMACAddress
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;

namespace Products3._0
{
    public partial class login : Form
    {
        public login()
        {
            InitializeComponent();
            //clientMAC = "7824AFC07705"; //cmd -> getmac -  put the correct value here when you know it
            clientMAC = "2C4D54671EB9"; //my lab pc
            clientMAC_new = "40A3CC7F7895";
            //clientMAC = "025041000001";
            //mypc: 025041000001
            GetMACAddress(); // run function when run the program 
            Clipboard.SetText(thisComputerMAC); //copy text to clipboard 
            textBox2.Text = thisComputerMAC;  //copy thisComputerMAC to textbox2 to validate the mac 
        }

     


        /// <summary>
        /// Log file create function
        /// </summary>
        /// <param name="sEventName"></param>
        /// <param name="sControlName"></param>
        /// <param name="sFormName"></param>
        public void LogFile(string sEventName, string sControlName, string sFormName)
        {
            StreamWriter log;
            if (!File.Exists("logfile.txt"))
            {
                log = new StreamWriter("logfile.txt");
            }
            else
            {
                log = File.AppendText("logfile.txt");
            }
            // Write to the file:
            log.WriteLine("===============================================Srart============================================");
            log.WriteLine("Data Time:" + DateTime.Now);
            log.WriteLine("--------------");
            //log.WriteLine("Exception Name:" + sExceptionName);
            log.WriteLine("Event Name:" + sEventName);
            log.WriteLine("---------------");
            log.WriteLine("Control Name:" + sControlName);
            log.WriteLine("---------------");
            log.WriteLine("Form Name:" + sFormName);
            log.WriteLine("===============================================End==============================================");
            // Close the stream:
            log.Close();
        }



        /// <summary>
        /// mac adress function
        /// </summary>
        /// <returns></returns>
        /// 
        string thisComputerMAC;
        string clientMAC;
        string clientMAC_new;
        public string GetMACAddress()
        {
                NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
                String sMacAddress = string.Empty;
            try
            {
                foreach (NetworkInterface adapter in nics)
                {
                    if (sMacAddress == String.Empty)// only return MAC Address from first card  
                    {
                        IPInterfaceProperties properties = adapter.GetIPProperties();
                        sMacAddress = adapter.GetPhysicalAddress().ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                LogFile("mac adress function-Eror Message: " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
                thisComputerMAC = sMacAddress.ToString();
                return sMacAddress; // this function shoud retrun value              
        }
     
        
  
        /// <summary>
        /// get first network ethernet card name function
        /// </summary>
        public void getNetwork()
        {        
            NetworkInterface[] interfaces = NetworkInterface.GetAllNetworkInterfaces();
            try
            {
                foreach (NetworkInterface adapter in interfaces)
                {
                    MessageBox.Show(adapter.Description);
                    return;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                LogFile("getNetwork function-Eror Message: " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }

        //MenuStrip////////////////////////////////////////////////////////////////////


        /// <summary>
        /// activate button menueitem menuestrip 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void activateAndCopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(thisComputerMAC); //copy text to clipboard 
            textBox2.Text = thisComputerMAC;  //copy thisComputerMAC to textbox2 to validate the mac 
            //MessageBox.Show(thisComputerMAC);  

        }

        /// <summary>
        /// Ethernet adapter button menueitem  menuestrip
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void networkCardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            getNetwork();
        }


        /// <summary>
        /// about
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            var About = new About();
            About.Show();
        }


        /// <summary>
        /// exit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }


        /// <summary>
        /// enter button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        /// 
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!String.IsNullOrEmpty(textBox2.Text))
                {
                    if (textBox2.Text == clientMAC && textBox2.Text == thisComputerMAC || textBox2.Text == clientMAC_new && textBox2.Text == thisComputerMAC) //check textbox2(thiscomputermac) if equal to clientmac(my code entered)
                    {
                        this.Hide(); // hide login form after correct login
                        var form1 = new Form1();
                        //ShowDialog shows the form modally, meaning you cannot go to the parent form
                        //Show function shows the form in a non modal form. This means that you can click on the parent form.                
                        form1.ShowDialog(); // open new form1 and disable move (login.cs) 
                        this.Close(); //close all (login+form1)
                    }
                    else
                    {
                        MessageBox.Show("Your Account not activated");
                        textBox2.Text = "";
                    }
                }
                else { MessageBox.Show("Please Activate Your Account "); }
            }
            catch (Exception ex)
            {
                LogFile("login - Enter button-Exception " + ex.Message, ex.StackTrace, this.FindForm().Name);
            }
        }


    }
}
