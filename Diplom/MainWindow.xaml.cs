using Diplom.Models;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using static Diplom.AppData;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static Users currentUser = new Users();
        private static Users currentUserTemp = new Users();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Login(object sender, RoutedEventArgs e)
        {
            try
            {
                currentUserTemp = GetContext().Users.Where(x => x.Login.Equals(tbLogin.Text)).FirstOrDefault();
                if (currentUserTemp == null)
                {
                    MessageBox.Show("Такого пользователя нет", "Ошибка авторизации", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                else
                {
                    if (currentUserTemp.TemporaryPassword == null)
                        currentUser = GetContext().Users.Where(x => x.Login.Equals(tbLogin.Text) && x.Password.Equals(pbPassword.Password)).FirstOrDefault();
                    else
                        currentUser = GetContext().Users.Where(x => x.Login.Equals(tbLogin.Text) && (x.TemporaryPassword.Equals(pbPassword.Password) || x.Password.Equals(pbPassword.Password))).FirstOrDefault();
                    if (currentUser == null)
                    {
                        MessageBox.Show("Неверный пароль", "Ошибка авторизации", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    else
                    {
                        switch (currentUser.RoleID)
                        {
                            case 1:
                                var wnd = new StudentsWindow(currentUser)
                                {
                                    Title = "Заведующий отделением"
                                };
                                wnd.Show();
                                Hide();
                                break;
                            case 2:
                                var wndw = new MyGroupsWindow(currentUser)
                                {
                                    Title = "Классный руководитель"
                                };
                                wndw.Show();
                                Hide();
                                break;
                            default:
                                MessageBox.Show("Данные не обнаружены", "Уведомление", MessageBoxButton.OK, MessageBoxImage.Information);
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка" + ex.Message.ToString() + "Критическая работа приложения", "Уведомление", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void Image_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            pbPassword.Visibility = Visibility.Collapsed;
            tbPassword.Visibility = Visibility.Visible;
            tbPassword.Text = pbPassword.Password;
            imgVisibility.Source = new BitmapImage(new Uri("/Resources/icons/iconVisibleT.png", UriKind.RelativeOrAbsolute));
        }

        private void Image_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            pbPassword.Visibility = Visibility.Visible;
            tbPassword.Visibility = Visibility.Collapsed;
            imgVisibility.Source = new BitmapImage(new Uri("/Resources/icons/iconVisibleF.png", UriKind.RelativeOrAbsolute));
        }
    }
}
