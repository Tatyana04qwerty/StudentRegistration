using Diplom.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для ProfileWindow.xaml
    /// </summary>
    public partial class ProfileWindow : Window
    {
        private static Users _currentUser = new Users();

        public ProfileWindow(Users current)
        {
            InitializeComponent();
            _currentUser = current;
            UpLoad();
        }

        public void UpLoad()
        {
            DataContext = _currentUser;
            lbFIO.Content = $"{_currentUser.Surname} {_currentUser.Name} {_currentUser.Patronymic}";
        }

        private void SaveProfile_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var item = _currentUser;
                if (tbPhone.Text == "" || tbEmail.Text == "" || tbLogin.Text == "" || tbPassword.Text == "")
                {
                    MessageBox.Show("Все поля должны быть заполнены", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    Regex regex = new Regex(@"\w*\@\w*\.\w*");
                    MatchCollection matches = regex.Matches(tbEmail.Text);
                    if (matches.Count == 0)
                    {
                        MessageBox.Show("Неверный формат адреса почты", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        item.Phone = tbPhone.Text;
                        item.Email = tbEmail.Text;
                        item.Login = tbLogin.Text;
                        item.Password = tbPassword.Text;
                        item.TemporaryPassword = null;

                        GetContext().SaveChanges();
                        MessageBox.Show("Данные сохранены");
                    }
                }
            }
            catch
            {
                MessageBox.Show("Введите корректный формат");
            }
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string chars = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM!@#$%^&*()";
                string result = "";

                Random rnd = new Random();
                for (int i = 0; i < 14; i++)
                    result += chars[rnd.Next(chars.Length)];

                tbPassword.Text = result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }
    }
}
