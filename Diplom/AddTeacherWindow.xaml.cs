using Diplom.Models;
using System;
using System.Collections.Generic;
using System.IO;
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
using Path = System.IO.Path;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для AddTeacherWindow.xaml
    /// </summary>
    public partial class AddTeacherWindow : Window
    {
        private static Roles _role = new Roles();
        private List<Users> _listUsers = new List<Users>();
        private List<Roles> _listRoles = new List<Roles>();
        int toCheck;

        public AddTeacherWindow()
        {
            InitializeComponent();
            UpLoad();
            Events();
        }

        private void UpLoad()
        {
            _listUsers = GetContext().Users.ToList();
            _listRoles = GetContext().Roles.ToList();
            if (cbRole.Items.Count == 0)
            {
                foreach(Roles role in _listRoles)
                {
                    cbRole.Items.Add(role.RoleName);
                }
            }
            tbName.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbSurname.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbPatronymic.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbPhone.PreviewTextInput += new TextCompositionEventHandler(textBoxPhone_PreviewTextInput);
        }

        void textBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[А-Яа-я]").Success) e.Handled = true;
        }

        void textBoxPhone_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[0-9]").Success) e.Handled = true;
        }

        // Слои окна
        #region
        private void Events()
        {
            btnNext.Click += (s, e) =>
            {
                if (tbSurname.Text == "" || tbName.Text == "" || tbPatronymic.Text == "")
                {
                    MessageBox.Show("Заполните обязательные поля", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    if (_listUsers.Where(x => x.Surname == tbSurname.Text && x.Name == tbName.Text && x.Patronymic == tbPatronymic.Text).Count() != 0)
                    {
                        MessageBox.Show($"Такой пользователь уже существует", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        if (tbEmail.Text != "")
                        {
                            Regex regex = new Regex(@"\w*\@\w*\.\w*");
                            MatchCollection matches = regex.Matches(tbEmail.Text);
                            if (matches.Count == 0)
                            {
                                MessageBox.Show("Неверный формат адреса почты", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                            else
                            {
                                lbFIO.Content = tbSurname.Text + " " + tbName.Text + " " + tbPatronymic.Text;
                                SwitchLayers(nameof(logPass));
                                SwitchLayers(nameof(FinishGegistrate));
                            }
                        }
                        else
                        {
                            lbFIO.Content = tbSurname.Text + " " + tbName.Text + " " + tbPatronymic.Text;
                            SwitchLayers(nameof(logPass));
                            SwitchLayers(nameof(FinishGegistrate));
                        }
                    }
                }
            };

            btnBack.Click += (s, e) =>
            {
                SwitchLayers(nameof(StartGegistrate));
            };
        }

        private void SwitchLayers(string LayerName)
        {
            List<Grid> layers = new List<Grid>()
            {
                StartGegistrate,
                FinishGegistrate
            };

            foreach (var layer in layers)
            {
                layer.Visibility = (layer.Name == LayerName) ? Visibility.Visible : Visibility.Hidden;
            }
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            SwitchLayers(nameof(FIO));
        }
        #endregion

        private void Registrate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _role = _listRoles.Where(x => x.RoleName.Equals(cbRole.Text)).FirstOrDefault();
                if (tbLogin.Text != "" && tbPassword.Text != "")
                {
                    if (_listUsers.Where(x => x.Login == tbLogin.Text).Count() == 0)
                    {
                        var item = new Users
                        {
                            Surname = tbSurname.Text,
                            Name = tbName.Text,
                            Patronymic = tbPatronymic.Text,
                            Login = tbLogin.Text,
                            Password = tbPassword.Text,
                            RoleID = _role.ID
                        };
                        if (tbEmail.Text != "")
                        {
                            item.Email = tbEmail.Text;
                        }
                        if (tbPhone.Text != "+7(___)___-__-__")
                        {
                            item.Phone = tbPhone.Text;
                        }

                        GetContext().Users.Add(item);
                        GetContext().SaveChanges();

                        try
                        {
                            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                            var filePath = Path.Combine(desktopPath, $"{item.Surname}_логин_пароль.txt");
                            StreamWriter sw = new StreamWriter(filePath);
                            sw.WriteLine("Логин");
                            sw.WriteLine(tbLogin.Text);
                            sw.WriteLine("Пароль");
                            sw.WriteLine(tbPassword.Text);
                            sw.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Ошибка записи файла " + ex.Message);
                        }

                        MessageBox.Show("Пользователь добавлен. Файл с логином и паролем сохранен на рабочий стол");
                        Hide();
                    }
                    else
                    {
                        MessageBox.Show("Пользователь с таким логином уже существует", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    MessageBox.Show("Заполните поля логина и пароля", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string charsPass = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM!@#$%&";
                string charsLog = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPASDFGHJKLZXCVBNM";
                string resultPass = "";
                string resultLog = "";

                Random rnd = new Random();
                for (int i = 0; i < 14; i++)
                    resultPass += charsPass[rnd.Next(charsPass.Length)];

                Random rnd1 = new Random();
                bool check = true;
                while (check)
                {
                    for (int i = 0; i < 14; i++)
                        resultLog += charsLog[rnd.Next(charsLog.Length)];
                    toCheck = GetContext().Users.Where(x => x.Login.Equals(resultLog)).Count();
                    if(toCheck == 0) check = false;
                }

                tbPassword.Text = resultPass;
                tbLogin.Text = resultLog;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }
    }
}
