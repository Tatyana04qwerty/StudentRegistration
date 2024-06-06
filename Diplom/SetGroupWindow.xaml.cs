using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Collections;
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
using System.Xml.Linq;
using Xceed.Wpf.Toolkit.PropertyGrid.Attributes;
using static Diplom.AppData;
using System.Drawing;
using Font = iTextSharp.text.Font;
using Document = iTextSharp.text.Document;
using Paragraph = iTextSharp.text.Paragraph;
using Diplom.Models;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для SetGroupWindow.xaml
    /// </summary>
    public partial class SetGroupWindow : Window
    {
        private static IList _listStudents = new List<Students>();
        private static Groups _currentGroup = new Groups();
        private static Specialities _currentSpeciality = new Specialities();

        public SetGroupWindow(IList selectedItems)
        {
            InitializeComponent();
            _listStudents = selectedItems;
            UdLoad();
        }

        private void UdLoad()
        {
            tblInfo.Text = $"Выбрано студентов для перевода: {_listStudents.Count}";
            tbGroup.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
        }

        void textBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[0-9]").Success) e.Handled = true;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (tbGroup.Text != "" && tbGroup.Text.Length == 5)
                {
                    _currentGroup = GetContext().Groups.Where(x => x.GroupNumber.Equals(tbGroup.Text)).FirstOrDefault();
                    int k = 0;
                    Students first = (Students)_listStudents[0];
                    string spec = first.Specialities.SpecialityName;
                    foreach (Students student in _listStudents)
                    {
                        if (student.Specialities.SpecialityName != spec)
                        {
                            k++;
                        }
                    }
                    if (k > 0)
                    {
                        MessageBox.Show($"Нельзя перевести в группу студентов разной специальности", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        if (_currentGroup != null)
                        {
                            if(_currentGroup.Size != 0 && _currentGroup.Speciality != spec)
                            {
                                if (MessageBox.Show($"Специальность выбранных студентов отличается от специальности группы. Изменить специальность студентов и перевести в выбранную группу?", "Внимание",
                        MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                                {
                                    _currentGroup = GetContext().Groups.Where(x => x.GroupNumber.Equals(tbGroup.Text)).First();
                                    _currentSpeciality = GetContext().Specialities.Where(x => x.SpecialityName.Equals(spec)).First();
                                    foreach (Students student in _listStudents)
                                    {
                                        student.GroupID = _currentGroup.ID;
                                        student.SpecialityID = _currentSpeciality.ID;
                                    }

                                    GetContext().SaveChanges();
                                    MessageBox.Show($"Студенты переведены в группу {_currentGroup.GroupNumber}");
                                }
                            }
                            else
                            {
                                foreach (Students student in _listStudents)
                                {
                                    student.GroupID = _currentGroup.ID;
                                }

                                GetContext().SaveChanges();
                                MessageBox.Show($"Студенты переведены в группу {_currentGroup.GroupNumber}");
                            }
                        }
                        else
                        {
                            if (MessageBox.Show($"Группы {tbGroup.Text} не существует. Создать группу?", "Внимание",
                        MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                            {
                                var item = new Groups
                                {
                                    GroupNumber = tbGroup.Text
                                };
                                GetContext().Groups.Add(item);
                                GetContext().SaveChanges();
                                MessageBox.Show("Новая группа создана");

                                _currentGroup = GetContext().Groups.Where(x => x.GroupNumber.Equals(tbGroup.Text)).FirstOrDefault();
                                foreach (Students student in _listStudents)
                                {
                                    student.GroupID = _currentGroup.ID;
                                }

                                GetContext().SaveChanges();
                                MessageBox.Show($"Студенты переведены в группу {_currentGroup.GroupNumber}");
                            }
                        }
                    }
                }
                else if (tbGroup.Text == "")
                    MessageBox.Show("Заполните поле с номером группы", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                else
                    MessageBox.Show("Номер группы должен содержать 5 цифр.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }
    }
}
