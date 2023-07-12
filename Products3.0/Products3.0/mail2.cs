using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.Net.Mime;

namespace Products3._0
{
    public partial class mail2 : Form
    {
        //ArrayList alAttachments;
        MailMessage mailMessage;
        public mail2()
        {
            InitializeComponent();
            textBox2.PasswordChar = '*';
            textBox7.Enabled = false;
        }



        /// <summary>
        /// send mail
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            label7.Visible = true;
            try
            {
                if (textBox1.Text != null && textBox2.Text != null)
                {
                    mailMessage = new MailMessage(textBox1.Text, textBox4.Text);
                    //////////////////// logo picture in the body message
                    string htmlBody = "<html><body> <br><img src=\"cid:logo\" width='240' height='110'> </body></html>";
                    AlternateView avHtml = AlternateView.CreateAlternateViewFromString
                    (htmlBody, null, MediaTypeNames.Text.Html);
                    // Create a LinkedResource object for each embedded image
                    LinkedResource pic1 = new LinkedResource("logo.jpg", MediaTypeNames.Image.Jpeg);
                    pic1.ContentId = "logo";
                    avHtml.LinkedResources.Add(pic1);
                    mailMessage.AlternateViews.Add(avHtml);
                    ///////////////////////////////////////////////////////

                    mailMessage.Subject = textBox5.Text;
                    mailMessage.Body = textBox6.Text;
                    mailMessage.IsBodyHtml = true;

                    /* Set the SMTP server and send the email with attachment */

                    SmtpClient smtpClient = new SmtpClient();

                    // smtpClient.Host = emailServerInfo.MailServerIP;
                    //this will be the host in case of gamil and it varies from the service provider

                    smtpClient.Host = "smtp.gmail.com";
                    //smtpClient.Port = Convert.ToInt32(emailServerInfo.MailServerPortNumber);
                    //this will be the port in case of gamil for dotnet and it varies from the service provider

                    smtpClient.Port = 587;
                    smtpClient.UseDefaultCredentials = true;

                    //smtpClient.Credentials = new System.Net.NetworkCredential(emailServerInfo.MailServerUserName, emailServerInfo.MailServerPassword);
                    smtpClient.Credentials = new NetworkCredential(textBox1.Text, textBox2.Text);

                    //Attachment
                    Attachment attachment = new Attachment(textBox7.Text);
                    if (attachment != null)
                    {
                        mailMessage.Attachments.Add(attachment);
                    }


                    //this will be the true in case of gamil and it varies from the service provider
                    smtpClient.EnableSsl = true;
                    smtpClient.Timeout = 8000;  //delay to 8sec, because the tool with no internet will stuck in line above.12
                    smtpClient.Send(mailMessage);
                }
                label7.Visible = false;
                string msg = "Message Send Successfully ";
                msg += " To : " + textBox4.Text;

                MessageBox.Show(msg.ToString());

                /* clear the controls */
                //textBox1.Text = string.Empty;
                //textBox2.Text = string.Empty;
                //textBox4.Text = string.Empty;
                //textBox5.Text = string.Empty;
                //textBox6.Text = string.Empty;
                //textBox7.Text = string.Empty;
            }
            catch (Exception ex)
            {
                label7.Visible = false;
                MessageBox.Show(ex.Message.ToString());
            }
        }


        /// <summary>
        /// add excel file 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofdlg = new OpenFileDialog();
            if (ofdlg.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    textBox7.Text = ofdlg.FileName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                }
            }
        }


        /// <summary>
        /// exit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
