using Diplom.Models;
using DocumentFormat.OpenXml.Spreadsheet;
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
using Groups = Diplom.Models.Groups;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для AddNewGroupWindow.xaml
    /// </summary>
    public partial class AddNewGroupWindow : Window
    {
        private static Groups _current = new Groups();
        private static Groups _groupToCheck = new Groups();
        private static List<Students> _students = new List<Students>();

        public AddNewGroupWindow(Groups current)
        {
            InitializeComponent();
            _current = current;
            UpLoad();
        }

        public void UpLoad()
        {
            if (_current != null)
            {
                Title = "Редактирование группы";
                btnDelete.Visibility = Visibility.Visible;
                DataContext = _current;
                tblInfo.Text = "Группа " + _current.GroupNumber;
                btnCancel.Content = "Отменить изменения";
            }
            else
                Title = "Добавление группы";
            tbName.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
        }

        void textBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[0-9]").Success) e.Handled = true;
        }

        private void tbName_TextChanged(object sender, TextChangedEventArgs e)
        {
            tblInfo.Text = "Группа " + tbName.Text;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            // Создание новой группы
            if (_current == null)
            {
                try
                {
                    if (tbName.Text == "")
                    {
                        MessageBox.Show("Заполните поле номера группы");
                    }
                    else
                    {
                        _groupToCheck = GetContext().Groups.Where(x => x.GroupNumber == tbName.Text).FirstOrDefault();
                        if (_groupToCheck != null)
                        {
                            MessageBox.Show($"Такая группа уже существует", "Внимание");
                        }
                        else
                        {
                            var item = new Groups
                            {
                                GroupNumber = tbName.Text
                            };

                            GetContext().Groups.Add(item);
                            GetContext().SaveChanges();
                            MessageBox.Show("Группа сохранена");
                            Hide();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }

            // Изменение существующей группы
            else
            {
                try
                {
                    if (tbName.Text == "")
                    {
                        MessageBox.Show("Заполните поле номера группы");
                    }
                    else
                    {
                        _groupToCheck = GetContext().Groups.Where(x => x.GroupNumber == tbName.Text).FirstOrDefault();
                        if (_groupToCheck != null)
                        {
                            MessageBox.Show($"Такая группа уже существует", "Внимание");
                        }
                        else
                        {
                            var item = _current;
                            item.GroupNumber = tbName.Text;

                            GetContext().SaveChanges();
                            MessageBox.Show("Группа изменена");
                            Hide();
                        }
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
            _students = GetContext().Students.Where(x => x.GroupID == _current.ID).ToList();
            if (_students.Count != 0)
            {
                MessageBox.Show("Группа не может быть удалена, так как еще используется для учета студентов", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                var specialityForRemoving = GetContext().Specialities.Where(p => p.ID == _current.ID).ToList();
                if (MessageBox.Show($"Вы точно хотите удалить группу {_current.GroupNumber}?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        GetContext().Specialities.RemoveRange(specialityForRemoving);
                        GetContext().SaveChanges();
                        MessageBox.Show("Данные удалены");
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
