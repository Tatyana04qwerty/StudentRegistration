using System;
using System.Collections.Generic;
using System.Linq;
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
using static Diplom.AppData;
using Diplom.Models;
using System.IO;
using Path = System.IO.Path;
using System.Net;
using System.Net.Mail;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для GenerateNewPasswordWindow.xaml
    /// </summary>
    public partial class GenerateNewPasswordWindow : Window
    {
        private static Users _current = new Users();

        public GenerateNewPasswordWindow(Users current)
        {
            InitializeComponent();
            _current = current;
            tblInfo.Text = $"{_current.Surname} {_current.Name[0]}.{_current.Patronymic[0]}";
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string chars = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM!@#$%&";
                string result = "";

                Random rnd = new Random();
                for (int i = 0; i < 16; i++)
                    result += chars[rnd.Next(chars.Length)];

                tbNewPass.Text = result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var item = _current;
                if (tbNewPass.Text == "")
                {
                    MessageBox.Show("Сначала сгенерируйте пароль");
                }
                else
                {
                    item.TemporaryPassword = tbNewPass.Text;
                    //if(item.Email != null)
                    //{
                    //    try
                    //    {
                    //        //MailAddress from = new MailAddress("collegekolomnatest@mail.ru", "ИС \"Колледж \"Коломна\"");
                    //        MailAddress from = new MailAddress("tanchikiva142850@gmail.com", "ИС \"Колледж \"Коломна\"");
                    //        MailAddress to = new MailAddress("testcollegekolomna@mail.ru", "TestToName");
                    //        MailMessage myMail = new MailMessage(from, to)
                    //        {
                    //            Subject = "Восстановление пароля",
                    //            SubjectEncoding = Encoding.UTF8,
                    //            Body = "Ваш логин:" + Environment.NewLine + item.Login + Environment.NewLine + "Временный пароль:" + Environment.NewLine + tbNewPass.Text,
                    //            BodyEncoding = Encoding.UTF8
                    //        };

                    //        SmtpClient mySmtpClient = new SmtpClient("smtp.mail.ru")
                    //        {
                    //            Port = 587,
                    //            EnableSsl = true,
                    //            UseDefaultCredentials = true,
                    //            Credentials = new NetworkCredential(from.Address, "trustnobody142840"),
                    //        };

                    //        mySmtpClient.Send(myMail);
                    //    }

                    //    catch (SmtpException ex)
                    //    {
                    //        throw new ApplicationException
                    //          ("Smtp исключение: " + ex.Message);
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        throw ex;
                    //    }
                    //}
                    //else
                    //{
                        try
                        {
                            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                            var filePath = Path.Combine(desktopPath, $"{item.Surname}_временный_пароль.txt");
                            StreamWriter sw = new StreamWriter(filePath);
                            sw.WriteLine("Временный пароль");
                            sw.WriteLine(tbNewPass.Text);
                            sw.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Ошибка записи файла " + ex.Message);
                        }
                    //}

                    GetContext().SaveChanges();
                    MessageBox.Show("Временный пароль сгенерирован");
                    Hide();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }
    }
}
