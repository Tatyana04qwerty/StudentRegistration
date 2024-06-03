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
using System.Windows.Shapes;
using static Diplom.AppData;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для TeachersWindow.xaml
    /// </summary>
    public partial class TeachersWindow : Window
    {
        private static List<Users> _list = new List<Users>();
        private static Users _currentUser = new Users();

        public TeachersWindow(Users current)
        {
            InitializeComponent();
            _currentUser = current;
            UpLoad();
        }

        public void UpLoad()
        {
            _list = GetContext().Users.ToList();
            dgTeachers.ItemsSource = _list;
        }

        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                dgTeachers.ItemsSource = _list.Where(Item => Item.Surname.ToLower().Contains(tbSearch.Text.ToLower())
                || Item.Name.ToLower().Contains(tbSearch.Text.ToLower()) || Item.Patronymic.ToLower().Contains(tbSearch.Text.ToLower()));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void AddNewTeacher_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new AddTeacherWindow();
            wnd.ShowDialog();
            UpLoad();
        }

        private void GenerateNewPassword_Click(object sender, RoutedEventArgs e)
        {
            Users current = (sender as Button)?.DataContext as Users;
            var wnd = new GenerateNewPasswordWindow(current);
            wnd.ShowDialog();
            UpLoad();
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            Users current = (sender as Button)?.DataContext as Users;
            if(current.ID == _currentUser.ID)
            {
                MessageBox.Show("Пользователь не может удалить сам себя", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else if (MessageBox.Show($"Вы точно хотите удалить пользователя из системы?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                try
                {
                    GetContext().Users.Remove(current);
                    GetContext().SaveChanges();
                    MessageBox.Show("Пользователь удален");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
                UpLoad();
        }
    }
}
