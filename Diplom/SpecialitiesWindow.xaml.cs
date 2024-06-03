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
    /// Логика взаимодействия для SpecialitiesWindow.xaml
    /// </summary>
    public partial class SpecialitiesWindow : Window
    {
        public SpecialitiesWindow()
        {
            InitializeComponent();
            UpLoad();
        }

        public void UpLoad()
        {
            dgSpecialities.ItemsSource = GetContext().Specialities.ToList();
        }

        private void AddNewSpeciality_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new AddNewSpecialityWindow(null)
            {
                Title = "Добавление новой специальности"
            };
            wnd.ShowDialog();
            UpLoad();
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            Specialities current = (sender as Button)?.DataContext as Specialities;
            var wnd = new AddNewSpecialityWindow(current)
            {
                Title = "Изменение существующей специальности"
            };
            wnd.ShowDialog();
            UpLoad();
        }
    }
}
