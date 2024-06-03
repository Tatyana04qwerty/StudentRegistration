using Diplom.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
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
    /// Логика взаимодействия для AddStudentWindow.xaml
    /// </summary>
    public partial class AddStudentWindow : Window
    {
        private static Students _studentToCheck = new Students();

        public AddStudentWindow()
        {
            InitializeComponent();
            UpLoad();
        }

        public void UpLoad()
        {
            cbSpeciality.ItemsSource = GetContext().Specialities.ToList();

            cbIsOrphan.Items.Insert(0, "Да");
            cbIsOrphan.Items.Add("Нет");

            cbIsInvalid.Items.Insert(0, "Да");
            cbIsInvalid.Items.Add("Нет");

            cbGender.Items.Insert(0, "Муж.");
            cbGender.Items.Add("Жен.");

            cbFormOfStudy.Items.Insert(0, "Очная");
            cbFormOfStudy.Items.Add("Заочная");

            cbTypeOfFinancing.Items.Insert(0, "Бюджет");
            cbTypeOfFinancing.Items.Add("Договор");

            tbNozology.Visibility = Visibility.Hidden;
            lbNozology.Visibility = Visibility.Hidden;

            tbName.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            tbSurname.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            tbPatronymic.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            tbPhone.PreviewTextInput += new TextCompositionEventHandler(textBoxPhone_PreviewTextInput);
            tbPassportSeries.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbPassportNumber.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbIIAN.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbITN.PreviewTextInput += new TextCompositionEventHandler(textBox_PreviewTextInput);
            tbGPA.PreviewTextInput += new TextCompositionEventHandler(textBoxDecimal_PreviewTextInput);
        }

        void textBoxPhone_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[0-9]").Success) e.Handled = true;
        }

        void textBoxDecimal_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[0-9.]").Success) e.Handled = true;
        }

        void textBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[0-9\-]").Success) e.Handled = true;
        }

        void textBoxText_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[а-яА-Я]").Success) e.Handled = true;
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

        private void SaveChanges_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (tbSurname.Text == "" || tbName.Text == "" || tbPatronymic.Text == "" || tbPhone.Text == "+7(___)___-__-__"
                    || tbPassportSeries.Text == "" || tbPassportNumber.Text == "" || tbIssuedBy.Text == ""
                     || dpDate.SelectedDate == null || tbAddress.Text == "" || tbIIAN.Text == "" || dpDateOfBirth.SelectedDate == null)
                {
                    MessageBox.Show("Заполните обязательные поля", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    _studentToCheck = GetContext().Students.Where(x => x.Surname == tbSurname.Text && x.Name == tbName.Text
                        && x.Patronymic == tbPatronymic.Text && x.Phone == tbPhone.Text).FirstOrDefault();
                    if (_studentToCheck != null)
                    {
                        MessageBox.Show($"Такой студент уже существует", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        if (tbEmail.Text != "")
                        {
                            Regex regex = new Regex(@"\w*\@\w*\.\w*");
                            MatchCollection matches = regex.Matches(tbEmail.Text);
                            if (matches.Count == 0)
                            {
                                MessageBox.Show("Неверный формат адреса почты", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                            else
                            {
                                var item = new Students
                                {
                                    Surname = tbSurname.Text,
                                    Name = tbName.Text,
                                    Patronymic = tbPatronymic.Text,
                                    Gender = cbGender.Text,
                                    DateOfBirth = dpDateOfBirth.SelectedDate.Value.Date,
                                    Phone = tbPhone.Text,
                                    Nationality = tbNationality.Text,
                                    IdentityDocument = tbIdentityDocument.Text,
                                    PassportSeries = tbPassportSeries.Text,
                                    PassportNumber = tbPassportNumber.Text,
                                    IssuedBy = tbIssuedBy.Text,
                                    DateOfIssue = dpDate.SelectedDate.Value.Date,
                                    PermanentRegistrationAddress = tbAddress.Text,
                                    IIAN = tbIIAN.Text,
                                    ITN = tbITN.Text,
                                    BasicEducation = tbBaseEducation.Text,
                                    GPA = Convert.ToDouble(tbGPA.Text),
                                    SpecialityID = cbSpeciality.SelectedIndex + 1,
                                    TypeOfFinancing = cbTypeOfFinancing.Text,
                                    FormOfStudy = cbFormOfStudy.Text,
                                    IsEmployed = false,
                                    StatusID = 1
                                };
                                if (cbIsOrphan.SelectedIndex == 0)
                                    item.IsOrphan = true;
                                else item.IsOrphan = false;
                                if (cbIsInvalid.SelectedIndex == 0)
                                    item.IsInvalid = true;
                                else item.IsInvalid = false;

                                if (tbNozology.IsEnabled)
                                {
                                    if (tbNozology.Text != "")
                                    {
                                        item.CauseOfDisability = tbNozology.Text;
                                    }
                                }

                                GetContext().Students.Add(item);
                                GetContext().SaveChanges();
                                MessageBox.Show("Студент добавлен");
                                Hide();
                            }
                        }
                        else
                        {
                            var item = new Students
                            {
                                Surname = tbSurname.Text,
                                Name = tbName.Text,
                                Patronymic = tbPatronymic.Text,
                                Gender = cbGender.Text,
                                DateOfBirth = dpDateOfBirth.SelectedDate.Value.Date,
                                Phone = tbPhone.Text,
                                Nationality = tbNationality.Text,
                                IdentityDocument = tbIdentityDocument.Text,
                                PassportSeries = tbPassportSeries.Text,
                                PassportNumber = tbPassportNumber.Text,
                                IssuedBy = tbIssuedBy.Text,
                                DateOfIssue = dpDate.SelectedDate.Value.Date,
                                PermanentRegistrationAddress = tbAddress.Text,
                                IIAN = tbIIAN.Text,
                                ITN = tbITN.Text,
                                BasicEducation = tbBaseEducation.Text,
                                GPA = Convert.ToDouble(tbGPA.Text),
                                SpecialityID = cbSpeciality.SelectedIndex + 1,
                                TypeOfFinancing = cbTypeOfFinancing.Text,
                                FormOfStudy = cbFormOfStudy.Text,
                                IsEmployed = false,
                                StatusID = 1
                            };
                            if (cbIsOrphan.SelectedIndex == 0)
                                item.IsOrphan = true;
                            else item.IsOrphan = false;
                            if (cbIsInvalid.SelectedIndex == 0)
                                item.IsInvalid = true;
                            else item.IsInvalid = false;

                            GetContext().Students.Add(item);
                            GetContext().SaveChanges();
                            MessageBox.Show("Студент добавлен");
                            Hide();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
