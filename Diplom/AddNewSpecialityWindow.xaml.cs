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
    /// Логика взаимодействия для AddNewSpecialityWindow.xaml
    /// </summary>
    public partial class AddNewSpecialityWindow : Window
    {
        private static Specialities _current = new Specialities();
        private static Specialities _specialityToCheck = new Specialities();
        private static List<Students> _students = new List<Students>();

        public AddNewSpecialityWindow(Specialities current)
        {
            InitializeComponent();
            _current = current;
            UpLoad();
        }

        public void UpLoad()
        {
            if(_current != null)
            {
                btnDelete.Visibility = Visibility.Visible;
                DataContext = _current;
                tblInfo.Text = "Специальность " + _current.SpecialityName;
                btnCancel.Content = "Отменить изменения";
            }
            tbCode.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbName.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            tbAbbrevaite.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
        }

        void textBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[0-9]").Success) e.Handled = true;
        }

        void textBoxText_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[а-яА-Я]").Success) e.Handled = true;
        }

        private void tbName_TextChanged(object sender, TextChangedEventArgs e)
        {
            tblInfo.Text = "Специальность " + tbName.Text;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            // Создание новой специальности
            if (_current == null)
            {
                try
                {
                    if (tbCode.Text == "__.__.__")
                    {
                        MessageBox.Show("Заполните поле кода специальности", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else if (tbName.Text == "")
                    {
                        MessageBox.Show("Заполните поле названия специальности", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        _specialityToCheck = GetContext().Specialities.Where(x => x.SpecialityCode == tbCode.Text
                            && x.SpecialityName == tbName.Text).FirstOrDefault();
                        if (_specialityToCheck != null)
                        {
                            MessageBox.Show($"Такая специальность уже существует", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                        else
                        {
                            var item = new Specialities
                            {
                                SpecialityCode = tbCode.Text,
                                SpecialityName = tbName.Text
                            };
                            if (tbAbbrevaite.Text != "")
                                item.AbbreviatedName = tbAbbrevaite.Text;

                            GetContext().Specialities.Add(item);
                            GetContext().SaveChanges();
                            MessageBox.Show("Специальность сохранена");
                            Hide();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }

            // Изменение существующей специальности
            else
            {
                try
                {
                    if (tbCode.Text == "__.__.__")
                    {
                        MessageBox.Show("Заполните поле кода специальности", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else if (tbName.Text == "")
                    {
                        MessageBox.Show("Заполните поле названия специальности", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        var item = _current;

                        item.SpecialityCode = tbCode.Text;
                        item.SpecialityName = tbName.Text;
                        if (tbAbbrevaite.Text != "")
                            item.AbbreviatedName = tbAbbrevaite.Text;

                        GetContext().SaveChanges();
                        MessageBox.Show("Специальность изменена");
                        Hide();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            _students = GetContext().Students.Where(x => x.SpecialityID == _current.ID).ToList();
            if(_students.Count != 0)
            {
                MessageBox.Show("Специальность не может быть удалена, так как еще используется для учета студентов", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                var specialityForRemoving = GetContext().Specialities.Where(p => p.ID == _current.ID).ToList();
                if (MessageBox.Show($"Вы точно хотите удалить специальность {_current.SpecialityName}?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        GetContext().Specialities.RemoveRange(specialityForRemoving);
                        GetContext().SaveChanges();
                        MessageBox.Show("Данные удалены");
                        UpLoad();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                Hide();
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }
    }
}
