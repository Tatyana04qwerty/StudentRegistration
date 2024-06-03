using Diplom.Models;
using System;
using System.Collections.Generic;
using System.Linq;
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
using static Diplom.AppData;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для GroupsWindow.xaml
    /// </summary>
    public partial class GroupsWindow : Window
    {
        private static List<Groups> _list = new List<Groups>();
        private static List<Students> _students = new List<Students>();
        private static List<Specialities> _specialities = new List<Specialities>();
        private static List<string> listFilterSpeciality = new List<string>();

        public GroupsWindow()
        {
            InitializeComponent();
            UpLoad();
        }

        public void UpLoad()
        {
            _specialities = GetContext().Specialities.ToList();
            _list = GetContext().Groups.OrderBy(p => p.GroupNumber).ToList();
            dgGroups.ItemsSource = _list;

            if (CbFilterStatus.Items.Count == 0)
            {
                CbFilterStatus.Items.Insert(0, "Статус");
                CbFilterStatus.Items.Add("Обучается");
                CbFilterStatus.Items.Add("Выпущена");
            }
            if (CbFilterSpeciality.Items.Count == 0)
            {
                listFilterSpeciality = new List<string>();
                listFilterSpeciality.Insert(0, "Специальность");
                foreach (var speciality in _specialities)
                {
                    listFilterSpeciality.Add(speciality.SpecialityName.ToString());
                }
                CbFilterSpeciality.ItemsSource = listFilterSpeciality;
            }

            ComboBoxItem item = new ComboBoxItem
            {
                BorderBrush = Brushes.Black,
                BorderThickness = new Thickness(0, 0, 0, 1)
            };
            if(CbSort.Items.Count == 0)
            {
                CbSort.Items.Insert(0, "Сортировка");
                CbSort.Items.Add("Название группы по возрастанию");
                CbSort.Items.Add("Название группы по убыванию");
                CbSort.Items.Add(item);
                CbSort.Items.Add("Размер группы по возрастанию");
                CbSort.Items.Add("Размер группы по убыванию");
            }
            CbSort.SelectedIndex = 0;
        }

        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                dgGroups.ItemsSource = _list.Where(Item => Item.Users.Surname.ToLower().Contains(tbSearch.Text.ToLower())
                    || Item.Users.Name.ToLower().Contains(tbSearch.Text.ToLower()) || Item.Users.Patronymic.ToLower().Contains(tbSearch.Text.ToLower())
                     || Item.GroupNumber.Contains(tbSearch.Text));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void UpdateGroups()
        {
            _list = GetContext().Groups.OrderBy(p => p.GroupNumber).ToList();

            if (CbFilterSpeciality.SelectedIndex > 0)
            {
                Groups groupSpeciality = CbFilterSpeciality.SelectedItem as Groups;
                if (groupSpeciality != null)
                {
                    string filter = groupSpeciality.Speciality.ToString();
                    _list = _list.Where(p => p.Speciality == filter).ToList();
                }
            }

            switch (CbFilterStatus.SelectedIndex)
            {
                case 1:
                    _list = _list.Where(p => p.Status == "Обучается").ToList();
                    break;
                case 2:
                    _list = _list.Where(p => p.Status == "Выпущена").ToList();
                    break;
            }

            switch (CbSort.SelectedIndex)
            {
                case 1:
                    _list = _list.OrderBy(p => p.GroupNumber).ToList();
                    break;
                case 2:
                    _list = _list.OrderByDescending(p => p.GroupNumber).ToList();
                    break;
                case 4:
                    _list = _list.OrderBy(p => p.Size).ToList();
                    break;
                case 5:
                    _list = _list.OrderByDescending(p => p.Size).ToList();
                    break;
            }

            dgGroups.ItemsSource = _list;
        }

        private void cbSort_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateGroups();
        }

        private void cbFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateGroups();
        }

        private void SetFormMaster_Click(object sender, RoutedEventArgs e)
        {
            Groups current = (sender as Button)?.DataContext as Groups;
            var wnd = new SetFormMasterWindow(current);
            wnd.ShowDialog();
            UpLoad();
        }

        private void AddNewGroup_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new AddNewGroupWindow(null);
            wnd.ShowDialog();
            UpLoad();
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            Groups current = (sender as Button)?.DataContext as Groups;
            var wnd = new AddNewGroupWindow(current);
            wnd.ShowDialog();
            UpLoad();
        }

        private void Graduate_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = dgGroups.SelectedItems;
            if (selectedItems.Count != 0)
            {
                if (MessageBox.Show("Вы уверены, что хотите выпустить выбранные группы?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        foreach(Groups groups in selectedItems)
                        {
                            _students = GetContext().Students.Where(p => p.GroupID == groups.ID).ToList();
                            foreach (Students student in _students)
                            {
                                student.StatusID = 2;
                            }
                        }
                        GetContext().SaveChanges();
                        MessageBox.Show("Группы выпущены");
                        UpLoad();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Студенты не выбраны", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
