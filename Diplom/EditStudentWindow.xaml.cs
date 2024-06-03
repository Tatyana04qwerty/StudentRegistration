using System;
using System.CodeDom.Compiler;
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
using Diplom.Models;
using static System.Net.Mime.MediaTypeNames;
using static Diplom.AppData;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для EditStudentWindow.xaml
    /// </summary>
    public partial class EditStudentWindow : Window
    {
        private static Students _current = new Students();

        public EditStudentWindow(Students current)
        {
            InitializeComponent();
            _current = current;
            UpLoad();
        }

        public void UpLoad()
        {
            DataContext = _current;
            tblFIO.Text = _current.Surname + " " + _current.Name + " " + _current.Patronymic;

            cbIsOrphan.Items.Insert(0, "Да");
            cbIsOrphan.Items.Add("Нет");

            if (_current.IsOrphan == true)
                cbIsOrphan.SelectedIndex = 0;
            else
                cbIsOrphan.SelectedIndex = 1;

            cbIsInvalid.Items.Insert(0, "Да");
            cbIsInvalid.Items.Add("Нет");

            if (_current.IsInvalid == true)
                cbIsInvalid.SelectedIndex = 0;
            else
                cbIsInvalid.SelectedIndex = 1;

            cbFormOfStudy.Items.Insert(0, "Очная");
            cbFormOfStudy.Items.Add("Заочная");

            if(_current.FormOfStudy == "Очная")
                cbFormOfStudy.SelectedIndex = 0;
            else
                cbFormOfStudy.SelectedIndex = 1;

            cbTypeOfFinancing.Items.Insert(0, "Бюджет");
            cbTypeOfFinancing.Items.Add("Договор");

            if (_current.TypeOfFinancing == "Бюджет")
                cbTypeOfFinancing.SelectedIndex = 0;
            else
                cbTypeOfFinancing.SelectedIndex = 1;

            if (_current.IsOrphan)
                cbIsOrphan.SelectedIndex = 0;
            else
                cbIsOrphan.SelectedIndex = 1;
            if(_current.IsInvalid)
                cbIsInvalid.SelectedIndex = 0;
            else
            {
                cbIsInvalid.SelectedIndex = 1;
                tbNozology.Visibility = Visibility.Hidden;
                lbNozology.Visibility = Visibility.Hidden;
            }

            tbName.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            tbSurname.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            tbPatronymic.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            tbPhone.PreviewTextInput += new TextCompositionEventHandler(textBoxPhone_PreviewTextInput);
            tbPassportSeries.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbPassportNumber.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbIIAN.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbITN.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
        }

        void textBoxPhone_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[0-9]").Success) e.Handled = true;
        }

        void textBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[0-9\-]").Success) e.Handled = true;
        }

        void textBoxText_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[а-яА-Я]").Success) e.Handled = true;
        }

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //var item = _current;
                if (tbSurname.Text == "" || tbName.Text == "" || tbPatronymic.Text == "" || tbPhone.Text == ""
                    || tbPassportSeries.Text == "" || tbPassportNumber.Text == "" || tbIssuedBy.Text == ""
                     || dpDate.Text == "" || tbAddress.Text == "" || tbIIAN.Text == "")
                {
                    MessageBox.Show("Заполните обязательные поля", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    _current = GetContext().Students.FirstOrDefault(s => s.ID == _current.ID);
                    _current.Surname = tbSurname.Text;
                    _current.Name = tbName.Text;
                    _current.Patronymic = tbPatronymic.Text;
                    _current.Phone = tbPhone.Text;
                    _current.Email = tbEmail.Text;
                    _current.PassportSeries = tbPassportSeries.Text;
                    _current.PassportNumber = tbPassportNumber.Text;
                    _current.IssuedBy = tbIssuedBy.Text;
                    _current.DateOfIssue = dpDate.SelectedDate.Value.Date;
                    _current.PermanentRegistrationAddress = tbAddress.Text;
                    _current.IIAN = tbIIAN.Text;
                    _current.ITN = tbITN.Text;
                    _current.TypeOfFinancing = cbTypeOfFinancing.Text;
                    _current.FormOfStudy = cbFormOfStudy.Text;
                    
                    if (cbIsOrphan.SelectedIndex == 0)
                        _current.IsOrphan = true;
                    else _current.IsOrphan = false;
                    if (cbIsInvalid.SelectedIndex == 0)
                        _current.IsInvalid = true;
                    else _current.IsInvalid = false;

                    if (tbNozology.IsEnabled)
                    {
                        if(tbNozology.Text != "")
                        {
                            _current.CauseOfDisability = tbNozology.Text;
                        }
                    }
                    if(tbNote.Text == "")
                        _current.Note = null;
                    else
                        _current.Note = tbNote.Text;

                    GetContext().SaveChanges();
                    MessageBox.Show("Данные сохранены");
                    Hide();
                }
            }
            catch
            {
                MessageBox.Show("Введите корректный формат");
            }
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }

        private void cbIsInvalid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbIsInvalid.SelectedIndex == 0)
            {
                tbNozology.Visibility = Visibility.Visible;
                lbNozology.Visibility = Visibility.Visible;
            }
            else
            {
                tbNozology.Visibility = Visibility.Hidden;
                lbNozology.Visibility = Visibility.Hidden;
            }
        }
    }
}
