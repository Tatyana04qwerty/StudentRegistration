using Diplom.Models;
using DocumentFormat.OpenXml.ExtendedProperties;
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
    /// Логика взаимодействия для EmploymentWindow.xaml
    /// </summary>
    public partial class EmploymentWindow : Window
    {
        private static Students _current = new Students();
        private static List<Employment> _listEmployment = new List<Employment>();
        private static Employment _lastE = new Employment();
        private static Employment _currentE = new Employment();

        public EmploymentWindow(Students current)
        {
            InitializeComponent();
            _current = current;
            UpLoad();
            Events();
        }

        public void UpLoad()
        {
            _listEmployment = GetContext().Employment.Where(x => x.StudentID == _current.ID).ToList();
            dgEmployment.ItemsSource = _listEmployment.OrderByDescending(x => x.EmploymentDate);
            _lastE = _listEmployment.LastOrDefault();
            if (_listEmployment.Count != 0 && _lastE.DateOfDismissal == null)
                btnSet.Visibility = Visibility.Visible;
            else btnSet.Visibility = Visibility.Collapsed;
            tblFIO.Text = _current.Surname + " " + _current.Name + " " + _current.Patronymic;
        }

        // Слои окна
        #region
        private void Events()
        {
            btnReturn.Click += (s, e) =>
            {
                SwitchLayers(nameof(newEmployment));
            };
        }

        private void SwitchLayers(string LayerName)
        {
            List<Grid> layers = new List<Grid>()
            {
                newEmployment,
                oldEmployment
            };

            foreach (var layer in layers)
            {
                layer.Visibility = (layer.Name == LayerName) ? Visibility.Visible : Visibility.Hidden;
            }
        }

        private void Return_Click(object sender, RoutedEventArgs e)
        {
            SwitchLayers(nameof(addEmployment));
        }
        #endregion

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            Employment current = (sender as Button)?.DataContext as Employment;
            _currentE = current;
            DataContext = _currentE;
            SwitchLayers(nameof(editEmployment));
            SwitchLayers(nameof(oldEmployment));
        }

        private void Dismiss_Click(object sender, RoutedEventArgs e)
        {
            _lastE = _listEmployment.LastOrDefault();
            _lastE.DateOfDismissal = DateTime.Now;
            _current.IsEmployed = false;
            GetContext().SaveChanges();
            MessageBox.Show("Дата увольнения изменена");
            UpLoad();
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if(_lastE != null)
                {
                    if (_lastE.DateOfDismissal != null)
                    {
                        var item = new Employment
                        {
                            StudentID = _current.ID,
                            Company = tbNewCompany.Text,
                            Location = tbNewLocation.Text,
                            Description = tbNewDescription.Text,
                            EmploymentDate = dpNewEmploymentDate.SelectedDate.Value
                        };

                        GetContext().Employment.Add(item);
                        GetContext().SaveChanges();
                        MessageBox.Show("История трудоустройства обновлена");
                        UpLoad();
                    }
                    else
                    {
                        MessageBox.Show("Студент еще не уволен с предыдущего места работы", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    var itemNew = new Employment
                    {
                        StudentID = _current.ID,
                        Company = tbNewCompany.Text,
                        Location = tbNewLocation.Text,
                        Description = tbNewDescription.Text,
                        EmploymentDate = dpNewEmploymentDate.SelectedDate.Value
                    };

                    GetContext().Employment.Add(itemNew);
                    GetContext().SaveChanges();
                    _current.IsEmployed = true;
                    MessageBox.Show("История трудоустройства обновлена");
                    UpLoad();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if(dpDateOfDismissal.SelectedDate != null)
                {
                    if(dpDateOfDismissal.SelectedDate < dpEmploymentDate.SelectedDate)
                    {
                        MessageBox.Show("Дата увольнения не может быть меньше даты трудоустройства", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else if(dpDateOfDismissal.SelectedDate <= DateTime.Now)
                    {
                        var item = _currentE;

                        item.Company = tbCompany.Text;
                        item.Location = tbLocation.Text;
                        item.Description = tbDescription.Text;
                        item.EmploymentDate = dpEmploymentDate.SelectedDate.Value;
                        item.DateOfDismissal = dpDateOfDismissal.SelectedDate.Value;
                        _current.IsEmployed = false;

                        GetContext().SaveChanges();
                        MessageBox.Show("Изменения сохранены");
                        UpLoad();
                    }
                    else
                    {
                        MessageBox.Show("Нельзя поставить дату увольнения до ее наступления", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
    }
}
