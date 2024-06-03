using Diplom.Models;
using System;
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
using OfficeOpenXml;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using static Diplom.AppData;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для EmploymentHistoryWindow.xaml
    /// </summary>
    public partial class EmploymentHistoryWindow : Window
    {
        private static Users _currentUser = new Users();
        private static List<Employment> _list = new List<Employment>();
        private static List<Employment> _listRangeGroups = new List<Employment>();
        private static List<Employment> _listAllGroups = new List<Employment>();

        public EmploymentHistoryWindow(Users user)
        {
            InitializeComponent();
            _currentUser = user;
            UpLoad();
        }

        public void UpLoad()
        {
            if (_currentUser.RoleID == 1)
            {
                _listAllGroups = GetContext().Employment.OrderByDescending(p => p.EmploymentDate).ToList();
            }
            else
            {
                _listRangeGroups = GetContext().Employment.Where(x => x.Students.Groups.Users.ID == _currentUser.ID).OrderByDescending(p => p.EmploymentDate).ToList();
            }
            tbSearch.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            if (CbSort.Items.Count == 0)
            {
                CbSort.Items.Insert(0, "Сортировка");
                CbSort.Items.Add("Дата трудоустройства по возрастанию");
                CbSort.Items.Add("Дата трудоустройства по убыванию");
            }
            Update();
        }

        public void Update()
        {
            if(_listAllGroups != null)
                _list = _listAllGroups;
            else if(_listRangeGroups != null)
                _list = _listRangeGroups;
            dgEmployment.ItemsSource = _list;
        }

        void textBoxText_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[а-яА-Яё]").Success) e.Handled = true;
        }

        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                dgEmployment.ItemsSource = _list.Where(Item => Item.Students.Surname.ToLower().Contains(tbSearch.Text.ToLower()) || Item.Students.Name.ToLower().Contains(tbSearch.Text.ToLower())
                    || Item.Students.Patronymic.ToLower().Contains(tbSearch.Text.ToLower()) || Item.Company.ToLower().Contains(tbSearch.Text.ToLower()));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void Sort_Click(object sender, RoutedEventArgs e)
        {
            _list = GetContext().Employment.ToList();
            if (CbSort.SelectedIndex > 0)
            {
                switch (CbSort.SelectedIndex)
                {
                    case 1:
                        _list = _list.OrderBy(p => p.EmploymentDate).ToList();
                        break;
                    case 2:
                        _list = _list.OrderByDescending(p => p.EmploymentDate).ToList();
                        break;
                }
            }
            if (dpStart.SelectedDate != null && dpEnd.SelectedDate != null)
            {
                if (dpStart.SelectedDate < dpEnd.SelectedDate)
                {
                    _list = _list.Where(x => x.EmploymentDate >= dpStart.SelectedDate && x.DateOfDismissal <= dpEnd.SelectedDate ||
                    x.EmploymentDate >= dpStart.SelectedDate && (x.EmploymentDate < dpEnd.SelectedDate && x.DateOfDismissal >= dpEnd.SelectedDate || x.DateOfDismissal == null) ||
                    x.EmploymentDate <= dpStart.SelectedDate && (x.DateOfDismissal > dpStart.SelectedDate && x.DateOfDismissal <= dpEnd.SelectedDate || x.DateOfDismissal == null)).ToList();

                    dgEmployment.ItemsSource = _list;
                }
                else
                {
                    MessageBox.Show("Дата начала перода должна быть меньше даты конца");
                }
            }
            else if(dpStart.SelectedDate != null || dpEnd.SelectedDate != null)
            {
                MessageBox.Show("Заполните оба поля даты");
            }
            dgEmployment.ItemsSource = _list;
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            CbSort.SelectedIndex = 0;
            tbSearch.Clear();
            dpEnd.SelectedDate = null;
            dpStart.SelectedDate = null;

            Update();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog
            {
                FileName = "DocumentEmployment",
                DefaultExt = ".xlsx",
                Filter = "Excel Worksheets (.xlsx)|*.xlsx"
            };

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                string filename = dlg.FileName;

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage excelPackage = new ExcelPackage();
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("История трудоустройства студентов");

                //Заполнение
                try
                {
                    worksheet.Cells[1, 1].Value = "Фамилия";
                    worksheet.Cells[1, 2].Value = "Имя";
                    worksheet.Cells[1, 3].Value = "Отчество";
                    worksheet.Cells[1, 4].Value = "Пол";
                    worksheet.Cells[1, 5].Value = "Дата рождения";
                    worksheet.Cells[1, 5].AutoFitColumns();
                    worksheet.Cells[1, 6].Value = "Возраст";
                    worksheet.Cells[1, 7].Value = "Телефон";
                    worksheet.Cells[1, 8].Value = "Форма обучения";
                    worksheet.Cells[1, 8].AutoFitColumns();
                    worksheet.Cells[1, 9].Value = "Вид финансирования";
                    worksheet.Cells[1, 9].AutoFitColumns();
                    worksheet.Cells[1, 10].Value = "Специальность";
                    worksheet.Cells[1, 10].AutoFitColumns();
                    worksheet.Cells[1, 11].Value = "Группа";
                    worksheet.Cells[1, 12].Value = "Классный руководитель";
                    worksheet.Cells[1, 12].AutoFitColumns();
                    worksheet.Cells[1, 13].Value = "Статус";
                    worksheet.Cells[1, 14].Value = "Трудоустройство";
                    worksheet.Cells[1, 14].AutoFitColumns();
                    worksheet.Cells[1, 15].Value = "Место работы";
                    worksheet.Cells[1, 15].AutoFitColumns();
                    worksheet.Cells[1, 16].Value = "Адрес";
                    worksheet.Cells[1, 17].Value = "Дата трудоустройства";
                    worksheet.Cells[1, 17].AutoFitColumns();
                    worksheet.Cells[1, 18].Value = "Дата увольнения";
                    worksheet.Cells[1, 18].AutoFitColumns();
                    worksheet.Cells[1, 19].Value = "Описание";

                    int row = 2;
                    foreach (var item in _list)
                    {
                        worksheet.Cells[row, 1].Value = item.Students.Surname;
                        worksheet.Cells[row, 1].AutoFitColumns();
                        worksheet.Cells[row, 2].Value = item.Students.Name;
                        worksheet.Cells[row, 2].AutoFitColumns();
                        worksheet.Cells[row, 3].Value = item.Students.Patronymic;
                        worksheet.Cells[row, 3].AutoFitColumns();
                        worksheet.Cells[row, 4].Value = item.Students.Gender;
                        worksheet.Cells[row, 5].Value = item.Students.DateOfBirth;
                        worksheet.Cells[row, 5].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet.Cells[row, 6].Value = item.Students.Age;
                        worksheet.Cells[row, 7].Value = item.Students.Phone;
                        worksheet.Cells[row, 7].AutoFitColumns();
                        worksheet.Cells[row, 8].Value = item.Students.FormOfStudy;
                        worksheet.Cells[row, 9].Value = item.Students.TypeOfFinancing;
                        worksheet.Cells[row, 10].Value = item.Students.Specialities.SpecialityName;
                        worksheet.Cells[row, 10].Style.WrapText = true;
                        if (item.Students.GroupID == null)
                        {
                            worksheet.Cells[row, 11].Value = "На распределении";
                        }
                        else
                        {
                            worksheet.Cells[row, 11].Value = item.Students.Groups.GroupNumber;
                            worksheet.Cells[row, 12].Value = item.Students.Groups.Users.Surname + " " + item.Students.Groups.Users.Name + " " + item.Students.Groups.Users.Patronymic;
                        }
                        worksheet.Cells[row, 13].Value = item.Students.Statuses.Status;
                        worksheet.Cells[row, 14].Value = item.Students.IfEmployed;
                        worksheet.Cells[row, 15].Value = item.Company;
                        worksheet.Cells[row, 15].Style.WrapText = true;
                        worksheet.Cells[row, 16].Value = item.Location;
                        worksheet.Cells[row, 17].Style.WrapText = true;
                        worksheet.Cells[row, 17].Value = item.EmploymentDate;
                        worksheet.Cells[row, 17].Style.Numberformat.Format = "dd.MM.yyyy";
                        if (item.DateOfDismissal == null)
                            worksheet.Cells[row, 18].Value = "-";
                        else
                            worksheet.Cells[row, 18].Value = item.DateOfDismissal;
                        worksheet.Cells[row, 18].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet.Cells[row, 19].Value = item.Description;
                        worksheet.Cells[row, 19].AutoFitColumns();
                        row++;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }

                excelPackage.SaveAs(new FileInfo(filename));
                excelPackage.Dispose();
                MessageBox.Show("Excel документ сформирован успешно");
            }
        }
    }
}
