using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace PAra2305
{
    /// <summary>
    /// Логика взаимодействия для OtpravkaPocgtExcel.xaml
    /// </summary>
    public partial class OtpravkaPocgtExcel : Window
    {
        public OtpravkaPocgtExcel()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)//смотри правильно кнопки это excel  
        {
            



                TextRange range = new TextRange(MessegeRtb.Document.ContentStart, MessegeRtb.Document.ContentEnd);

                MailMessage messagev = new MailMessage(From.Text, To.Text, Subject.Text, range.Text);
                messagev.IsBodyHtml = true;
                messagev.Attachments.Add(new Attachment("питон.excel"));// я так и не понял какой путь файла 

                string server = "smtp.mail.ru";
                string servergm = "smtp.mail.ru";
                SmtpClient smtpclient;

                if (server == "smtp.yandex.ru")
                {
                    smtpclient = new SmtpClient("smtp.yandex.ru", 587);
                }
                else if (servergm == "smtp.gmail.com")
                {
                    smtpclient = new SmtpClient("smtp.gmail.com", 587);
                }
                else if (server == "smtp.rambler.ru")
                {
                    smtpclient = new SmtpClient("smtp.rambler.ru", 465);
                }
                else
                {
                    smtpclient = new SmtpClient("smtp.mail.ru", 587);
                }

                smtpclient.Credentials = new NetworkCredential(From.Text, Pass.Text);
                smtpclient.EnableSsl = true;
                smtpclient.Send(messagev);



            }
        }
    }

