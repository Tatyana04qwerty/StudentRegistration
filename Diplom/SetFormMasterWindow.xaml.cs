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
    /// Логика взаимодействия для SetFormMasterWindow.xaml
    /// </summary>
    public partial class SetFormMasterWindow : Window
    {
        private static Groups _currentGroup = new Groups();
        private static List<Users> _list = new List<Users>();

        public SetFormMasterWindow(Groups current)
        {
            InitializeComponent();
            _currentGroup = current;
            UpLoad();
        }

        public void UpLoad()
        {
            tblInfo.Text = $"Назначение классного руководителя для группы {_currentGroup.GroupNumber}";
            _list = GetContext().Users.ToList();
            dgFormMasters.ItemsSource = _list;
            tbSearch.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
        }

        void textBoxText_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[а-яА-Я]").Success) e.Handled = true;
        }

        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                dgFormMasters.ItemsSource = _list.Where(Item => Item.Surname.ToLower().Contains(tbSearch.Text.ToLower())
                    || Item.Name.ToLower().Contains(tbSearch.Text.ToLower()) || Item.Patronymic.ToLower().Contains(tbSearch.Text.ToLower()));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void SetFormMaster_Click(object sender, RoutedEventArgs e)
        {
            Users current = (sender as Button)?.DataContext as Users;
            if(_currentGroup.FormMaster == null)
            {
                _currentGroup.FormMaster = current.ID;
                GetContext().SaveChanges();
                MessageBox.Show($"Классный руководитель для группы {_currentGroup.GroupNumber} назначен");
            }
            else
            {
                if (MessageBox.Show($"У группы {_currentGroup.GroupNumber} уже есть классный руководитель. Изменить?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    _currentGroup.FormMaster = current.ID;
                    GetContext().SaveChanges();
                    MessageBox.Show($"Классный руководитель для группы {_currentGroup.GroupNumber} изменен");
                }
            }
            UpLoad();
        }
    }
}
