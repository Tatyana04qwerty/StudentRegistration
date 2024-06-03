using System;
using System.Collections.Generic;
using System.IO;
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
using OfficeOpenXml;
using Diplom.Models;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для MyGroupsWindow.xaml
    /// </summary>
    public partial class MyGroupsWindow : Window
    {
        private static Users _currentUser = new Users();
        private static List<Students> _list = new List<Students>();
        private static List<Students> _listSearch = new List<Students>();
        private static List<Students> _listGraduates = new List<Students>();
        private static List<Groups> _listGroups = new List<Groups>();
        private static List<Groups> _listCurrentGroups = new List<Groups>();
        private static List<Groups> _listGraduatedGroups = new List<Groups>();
        private static List<string> listFilterGroup = new List<string>();
        private static List<string> listFilterGraduatedGroup = new List<string>();

        public MyGroupsWindow(Users current)
        {
            InitializeComponent();
            _currentUser = current;
            if (_currentUser.TemporaryPassword != null)
            {
                _currentUser.TemporaryPassword = null;
                GetContext().SaveChanges();
            }
            UpLoad();
            Events();
        }

        public void UpLoad()
        {
            _listGroups = GetContext().Groups.Where(x => x.FormMaster == _currentUser.ID).ToList();
            _listCurrentGroups = _listGroups.Where(x => x.Status == "Обучается").ToList();
            _listGraduatedGroups = _listGroups.Where(x => x.Status == "Выпущена").ToList();

            // Формирование выпадающих списков фильтров
            #region

            if (CbSortDir.Items.Count == 0)
            {
                CbSortDir.Items.Insert(0, "Направление сортировки");
                CbSortDir.Items.Add("По возрастанию");
                CbSortDir.Items.Add("По убыванию");
            }
            if (CbSortField.Items.Count == 0)
            {
                CbSortField.Items.Insert(0, "Поле для сортировки");
                CbSortField.Items.Add("Фамилия");
                CbSortField.Items.Add("Дата рождения");
                CbSortField.Items.Add("Средний балл");
                CbSortField.Items.Add("Группа");
                CbSortField.Items.Add("Курс");
            }
            if (CbFilterGender.Items.Count == 0)
            {
                CbFilterGender.Items.Insert(0, "Пол");
                CbFilterGender.Items.Add("Мужской");
                CbFilterGender.Items.Add("Женский");
            }
            if (CbFilterTown.Items.Count == 0)
            {
                CbFilterTown.Items.Insert(0, "Город");
                CbFilterTown.Items.Add("Коломна, Коломенский р-н");
                CbFilterTown.Items.Add("Московская обл.");
                CbFilterTown.Items.Add("Прочие регионы");
            }
            if (CbFilterGroup.Items.Count == 0)
            {
                foreach (var group in _listCurrentGroups)
                {
                    listFilterGroup = new List<string>();
                    listFilterGroup.Add(group.GroupNumber.ToString());
                }
                CbFilterGroup.ItemsSource = listFilterGroup;
            }
            if (CbFilterAge.Items.Count == 0)
            {
                CbFilterAge.Items.Insert(0, "Возраст");
                CbFilterAge.Items.Add("Несовершеннолетние");
                CbFilterAge.Items.Add("Совершеннолетние");
                CbFilterAge.Items.Add("14 и младше");
                CbFilterAge.Items.Add("15");
                CbFilterAge.Items.Add("16");
                CbFilterAge.Items.Add("17");
                CbFilterAge.Items.Add("18");
                CbFilterAge.Items.Add("19");
                CbFilterAge.Items.Add("20");
                CbFilterAge.Items.Add("21");
                CbFilterAge.Items.Add("22");
                CbFilterAge.Items.Add("23");
                CbFilterAge.Items.Add("24");
                CbFilterAge.Items.Add("25");
                CbFilterAge.Items.Add("26");
                CbFilterAge.Items.Add("27");
                CbFilterAge.Items.Add("28");
                CbFilterAge.Items.Add("28");
                CbFilterAge.Items.Add("29");
                CbFilterAge.Items.Add("30-34");
                CbFilterAge.Items.Add("35-39");
                CbFilterAge.Items.Add("40 и старше");
            }
            if (CbFilterEmployment.Items.Count == 0)
            {
                CbFilterEmployment.Items.Insert(0, "Трудоустройство");
                CbFilterEmployment.Items.Add("Трудоустроен");
                CbFilterEmployment.Items.Add("Не трудоустроен");
            }
            if (CbFilterTypeOfFinancing.Items.Count == 0)
            {
                CbFilterTypeOfFinancing.Items.Insert(0, "Вид финансирования");
                CbFilterTypeOfFinancing.Items.Add("Бюджет");
                CbFilterTypeOfFinancing.Items.Add("Договор");
            }
            #endregion

            // Формирование выпадающих списков фильтров для выпускников
            #region

            if (CbSortDirGraduates.Items.Count == 0)
            {
                CbSortDirGraduates.Items.Insert(0, "Направление сортировки");
                CbSortDirGraduates.Items.Add("По возрастанию");
                CbSortDirGraduates.Items.Add("По убыванию");
            }
            if (CbSortFieldGraduates.Items.Count == 0)
            {
                CbSortFieldGraduates.Items.Insert(0, "Поле для сортировки");
                CbSortFieldGraduates.Items.Add("Фамилия");
                CbSortFieldGraduates.Items.Add("Дата рождения");
                CbSortFieldGraduates.Items.Add("Средний балл");
                CbSortFieldGraduates.Items.Add("Группа");
                CbSortFieldGraduates.Items.Add("Курс");
            }
            if (CbFilterGenderGraduates.Items.Count == 0)
            {
                CbFilterGenderGraduates.Items.Insert(0, "Пол");
                CbFilterGenderGraduates.Items.Add("Мужской");
                CbFilterGenderGraduates.Items.Add("Женский");
            }
            if (CbFilterTownGraduates.Items.Count == 0)
            {
                CbFilterTownGraduates.Items.Insert(0, "Город");
                CbFilterTownGraduates.Items.Add("Коломна, Коломенский р-н");
                CbFilterTownGraduates.Items.Add("Московская обл.");
                CbFilterTownGraduates.Items.Add("Прочие регионы");
            }
            if (CbFilterGroupGraduates.Items.Count == 0)
            {
                listFilterGraduatedGroup = new List<string>();
                foreach (var group in _listGraduatedGroups)
                {
                    listFilterGraduatedGroup.Add(group.GroupNumber.ToString());
                }
                CbFilterGroupGraduates.ItemsSource = listFilterGraduatedGroup;
            }
            if (CbFilterAgeGraduates.Items.Count == 0)
            {
                CbFilterAgeGraduates.Items.Insert(0, "Возраст");
                CbFilterAgeGraduates.Items.Add("Несовершеннолетние");
                CbFilterAgeGraduates.Items.Add("Совершеннолетние");
                CbFilterAgeGraduates.Items.Add("14 и младше");
                CbFilterAgeGraduates.Items.Add("15");
                CbFilterAgeGraduates.Items.Add("16");
                CbFilterAgeGraduates.Items.Add("17");
                CbFilterAgeGraduates.Items.Add("18");
                CbFilterAgeGraduates.Items.Add("19");
                CbFilterAgeGraduates.Items.Add("20");
                CbFilterAgeGraduates.Items.Add("21");
                CbFilterAgeGraduates.Items.Add("22");
                CbFilterAgeGraduates.Items.Add("23");
                CbFilterAgeGraduates.Items.Add("24");
                CbFilterAgeGraduates.Items.Add("25");
                CbFilterAgeGraduates.Items.Add("26");
                CbFilterAgeGraduates.Items.Add("27");
                CbFilterAgeGraduates.Items.Add("28");
                CbFilterAgeGraduates.Items.Add("28");
                CbFilterAgeGraduates.Items.Add("29");
                CbFilterAgeGraduates.Items.Add("30-34");
                CbFilterAgeGraduates.Items.Add("35-39");
                CbFilterAgeGraduates.Items.Add("40 и старше");
            }
            if (CbFilterEmploymentGraduates.Items.Count == 0)
            {
                CbFilterEmploymentGraduates.Items.Insert(0, "Трудоустройство");
                CbFilterEmploymentGraduates.Items.Add("Трудоустроен");
                CbFilterEmploymentGraduates.Items.Add("Не трудоустроен");
            }
            if (CbFilterTypeOfFinancingGraduates.Items.Count == 0)
            {
                CbFilterTypeOfFinancingGraduates.Items.Insert(0, "Вид финансирования");
                CbFilterTypeOfFinancingGraduates.Items.Add("Бюджет");
                CbFilterTypeOfFinancingGraduates.Items.Add("Договор");
            }
            #endregion

            Update();
        }


        // Слои окна
        #region
        private void Events()
        {
            btnAllStudents.Click += (s, e) =>
            {
                SwitchLayers(nameof(studentsGrid));
                Update();
            };

            btnGraduates.Click += (s, e) =>
            {
                SwitchLayers(nameof(graduatesGrid));
                UpdateGraduates();
            };
        }

        private void SwitchLayers(string LayerName)
        {
            List<Grid> layers = new List<Grid>()
            {
                studentsGrid,
                graduatesGrid
            };

            foreach (var layer in layers)
            {
                layer.Visibility = (layer.Name == LayerName) ? Visibility.Visible : Visibility.Hidden;
            }
        }

        private void Students_Click(object sender, RoutedEventArgs e)
        {
            SwitchLayers(nameof(students));
        }

        private void Graduates_Click(object sender, RoutedEventArgs e)
        {
            SwitchLayers(nameof(graduates));
        }
        #endregion

        // Обучающиеся группы
        #region
        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            CbSortDir.SelectedIndex = 0;
            CbSortField.SelectedIndex = 0;

            CbFilterGender.SelectedIndex = 0;
            CbFilterAge.SelectedIndex = 0;
            CbFilterTown.SelectedIndex = 0;
            CbFilterTypeOfFinancing.SelectedIndex = 0;
            CbFilterEmployment.SelectedIndex = 0;

            tbSearch.Clear();

            Update();
        }

        public void Update()
        {
            _list = GetContext().Students.Where(x => x.StatusID == 1).ToList();

            if (CbSortField.SelectedIndex > 0)
            {
                switch (CbSortField.SelectedIndex)
                {
                    case 1:
                        if (CbSortDir.SelectedIndex == 1)
                            _list = _list.OrderBy(p => p.Surname).ToList();
                        if (CbSortDir.SelectedIndex == 2)
                            _list = _list.OrderByDescending(p => p.Surname).ToList();
                        break;
                    case 2:
                        if (CbSortDir.SelectedIndex == 1)
                            _list = _list.OrderBy(p => p.DateOfBirth).ToList();
                        if (CbSortDir.SelectedIndex == 2)
                            _list = _list.OrderByDescending(p => p.DateOfBirth).ToList();
                        break;
                    case 3:
                        if (CbSortDir.SelectedIndex == 1)
                            _list = _list.OrderBy(p => p.GPA).ToList();
                        if (CbSortDir.SelectedIndex == 2)
                            _list = _list.OrderByDescending(p => p.GPA).ToList();
                        break;
                    case 4:
                        if (CbSortDir.SelectedIndex == 1)
                            _list = _list.OrderBy(p => p.Groups.GroupNumber).ToList();
                        if (CbSortDir.SelectedIndex == 2)
                            _list = _list.OrderByDescending(p => p.Groups.GroupNumber).ToList();
                        break;
                    case 5:
                        if (CbSortDir.SelectedIndex == 1)
                            _list = _list.OrderBy(p => p.Course).ToList();
                        if (CbSortDir.SelectedIndex == 2)
                            _list = _list.OrderByDescending(p => p.Course).ToList();
                        break;
                }
            }
            if (CbFilterGender.SelectedIndex > 0)
            {
                switch (CbFilterGender.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.Gender.Equals("Муж.")).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.Gender.Equals("Жен.")).ToList();
                        break;
                }
            }
            if (CbFilterTown.SelectedIndex > 0)
            {
                switch (CbFilterTown.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.PermanentRegistrationAddress.Contains("Коломна") || p.PermanentRegistrationAddress.Contains("Коломенский")).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.PermanentRegistrationAddress.Contains("Московская")).ToList();
                        break;
                    case 3:
                        _list = _list.Where(p => p.PermanentRegistrationAddress.Contains("Московская") == false).ToList();
                        break;
                }
            }
            if (CbFilterAge.SelectedIndex > 0)
            {
                switch (CbFilterAge.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.Age < 18).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.Age > 18).ToList();
                        break;
                    case 3:
                        _list = _list.Where(p => p.Age <= 14).ToList();
                        break;
                    case 4:
                        _list = _list.Where(p => p.Age == 15).ToList();
                        break;
                    case 5:
                        _list = _list.Where(p => p.Age == 16).ToList();
                        break;
                    case 6:
                        _list = _list.Where(p => p.Age == 17).ToList();
                        break;
                    case 7:
                        _list = _list.Where(p => p.Age == 18).ToList();
                        break;
                    case 8:
                        _list = _list.Where(p => p.Age == 19).ToList();
                        break;
                    case 9:
                        _list = _list.Where(p => p.Age == 20).ToList();
                        break;
                    case 10:
                        _list = _list.Where(p => p.Age == 21).ToList();
                        break;
                    case 11:
                        _list = _list.Where(p => p.Age == 22).ToList();
                        break;
                    case 12:
                        _list = _list.Where(p => p.Age == 23).ToList();
                        break;
                    case 13:
                        _list = _list.Where(p => p.Age == 24).ToList();
                        break;
                    case 14:
                        _list = _list.Where(p => p.Age == 25).ToList();
                        break;
                    case 15:
                        _list = _list.Where(p => p.Age == 26).ToList();
                        break;
                    case 16:
                        _list = _list.Where(p => p.Age == 27).ToList();
                        break;
                    case 17:
                        _list = _list.Where(p => p.Age == 28).ToList();
                        break;
                    case 18:
                        _list = _list.Where(p => p.Age == 29).ToList();
                        break;
                    case 19:
                        _list = _list.Where(p => p.Age >= 30 && p.Age <= 34).ToList();
                        break;
                    case 20:
                        _list = _list.Where(p => p.Age >= 35 && p.Age <= 39).ToList();
                        break;
                    case 21:
                        _list = _list.Where(p => p.Age >= 40).ToList();
                        break;
                }
            }
            if (CbFilterTypeOfFinancing.SelectedIndex > 0)
            {
                switch (CbFilterTypeOfFinancing.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.TypeOfFinancing.Equals("Бюджет")).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.TypeOfFinancing.Equals("Договор")).ToList();
                        break;
                }
            }
            if (CbFilterEmployment.SelectedIndex > 0)
            {
                switch (CbFilterEmployment.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.IsEmployed == true).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.IsEmployed == false).ToList();
                        break;
                }
            }

            _list = _list.Where(p => p.Groups.GroupNumber.Equals(CbFilterGroup.Text)).ToList();

            dgStudents.ItemsSource = _list;
            lbCount.Content = $"Количество записей: {_list.Count}";
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedItems = dgStudents.SelectedItems.OfType<int>().ToList();
            lbSelected.Content = $"Выбрано: {selectedItems.Count()}";
            int s = 0;
            foreach (var r in dgStudents.SelectedItems)
            {
                s++;
            }
            lbSelected.Content = $"Выбрано: {s}";
        }

        private void TbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                _listSearch = _list.Where(Item => Item.Surname.Contains(tbSearch.Text)
                    || Item.Name.ToLower().Contains(tbSearch.Text.ToLower()) || Item.Patronymic.ToLower().Contains(tbSearch.Text.ToLower()) || Item.HasNote.ToLower().Contains(tbSearch.Text.ToLower())).ToList();
                dgStudents.ItemsSource = _listSearch;

                lbCount.Content = $"Количество записей: {_listSearch.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void Sort_Click(object sender, RoutedEventArgs e)
        {
            Update();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Document";
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Worksheets (.xlsx)|*.xlsx";

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                string filename = dlg.FileName;

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage excelPackage = new ExcelPackage();
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Выборка студентов по фильтру");

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
                    worksheet.Cells[1, 8].Value = "Email";
                    worksheet.Cells[1, 9].Value = "Гражданство";
                    worksheet.Cells[1, 9].AutoFitColumns();
                    worksheet.Cells[1, 10].Value = "ДУЛ";
                    worksheet.Cells[1, 11].Value = "Серия паспорта";
                    worksheet.Cells[1, 11].AutoFitColumns();
                    worksheet.Cells[1, 12].Value = "Номер паспорта";
                    worksheet.Cells[1, 12].AutoFitColumns();
                    worksheet.Cells[1, 13].Value = "Дата выдачи ДУЛа";
                    worksheet.Cells[1, 13].AutoFitColumns();
                    worksheet.Cells[1, 14].Value = "Кем выдан ДУЛ";
                    worksheet.Cells[1, 14].AutoFitColumns();
                    worksheet.Cells[1, 15].Value = "Адрес постоянной регистрации";
                    worksheet.Cells[1, 15].AutoFitColumns();
                    worksheet.Cells[1, 16].Value = "СНИЛС";
                    worksheet.Cells[1, 17].Value = "ИНН";
                    worksheet.Cells[1, 18].Value = "Базовое образование";
                    worksheet.Cells[1, 18].AutoFitColumns();
                    worksheet.Cells[1, 19].Value = "Средний балл";
                    worksheet.Cells[1, 19].AutoFitColumns();
                    worksheet.Cells[1, 20].Value = "Является сиротой";
                    worksheet.Cells[1, 20].AutoFitColumns();
                    worksheet.Cells[1, 21].Value = "Является инвалидом";
                    worksheet.Cells[1, 21].AutoFitColumns();
                    worksheet.Cells[1, 22].Value = "Причина инвалидности";
                    worksheet.Cells[1, 22].AutoFitColumns();
                    worksheet.Cells[1, 23].Value = "Форма обучения";
                    worksheet.Cells[1, 23].AutoFitColumns();
                    worksheet.Cells[1, 24].Value = "Вид финансирования";
                    worksheet.Cells[1, 24].AutoFitColumns();
                    worksheet.Cells[1, 25].Value = "Специальность";
                    worksheet.Cells[1, 26].Value = "Группа";
                    worksheet.Cells[1, 27].Value = "Статус";
                    worksheet.Cells[1, 28].Value = "Примечания";
                    worksheet.Cells[1, 29].Value = "Трудоустройство";
                    worksheet.Cells[1, 29].AutoFitColumns();

                    int row = 2;
                    foreach (var item in _list)
                    {
                        worksheet.Cells[row, 1].Value = item.Surname;
                        worksheet.Cells[row, 2].Value = item.Name;
                        worksheet.Cells[row, 3].Value = item.Patronymic;
                        worksheet.Cells[row, 4].Value = item.Gender;
                        worksheet.Cells[row, 5].Value = item.DateOfBirth;
                        worksheet.Cells[row, 5].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet.Cells[row, 6].Value = item.Age;
                        worksheet.Cells[row, 7].Value = item.Phone;
                        worksheet.Cells[row, 8].Value = item.Email;
                        worksheet.Cells[row, 9].Value = item.Nationality;
                        worksheet.Cells[row, 10].Value = item.IdentityDocument;
                        worksheet.Cells[row, 11].Value = item.PassportSeries;
                        worksheet.Cells[row, 12].Value = item.PassportNumber;
                        worksheet.Cells[row, 13].Value = item.DateOfIssue;
                        worksheet.Cells[row, 13].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet.Cells[row, 14].Value = item.IssuedBy;
                        worksheet.Cells[row, 15].Value = item.PermanentRegistrationAddress;
                        worksheet.Cells[row, 15].Style.WrapText = true;
                        worksheet.Cells[row, 16].Value = item.IIAN;
                        worksheet.Cells[row, 17].Value = item.ITN;
                        worksheet.Cells[row, 18].Value = item.BasicEducation;
                        worksheet.Cells[row, 18].Style.WrapText = true;
                        worksheet.Cells[row, 19].Value = item.GPA;
                        if (item.IsOrphan)
                            worksheet.Cells[row, 20].Value = "Да";
                        else
                            worksheet.Cells[row, 20].Value = "Нет";
                        if (item.IsInvalid)
                            worksheet.Cells[row, 21].Value = "Да";
                        else
                            worksheet.Cells[row, 21].Value = "Нет";
                        worksheet.Cells[row, 22].Value = item.CauseOfDisability;
                        worksheet.Cells[row, 22].Style.WrapText = true;
                        worksheet.Cells[row, 23].Value = item.FormOfStudy;
                        worksheet.Cells[row, 24].Value = item.TypeOfFinancing;
                        worksheet.Cells[row, 25].Value = item.Specialities.SpecialityName;
                        worksheet.Cells[row, 25].Style.WrapText = true;
                        worksheet.Cells[row, 26].Value = item.Groups.GroupNumber;
                        worksheet.Cells[row, 27].Value = item.Statuses.Status;
                        worksheet.Cells[row, 28].Value = item.Note;
                        worksheet.Cells[row, 29].Value = item.IfEmployed;
                        row++;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                excelPackage.SaveAs(new FileInfo(filename));
                excelPackage.Dispose();
                MessageBox.Show("Excel документ сформирован успешно");
            }
        }

        #endregion

        // Выпутившиеся группы
        #region
        private void ClearGraduates_Click(object sender, RoutedEventArgs e)
        {
            _listGraduatedGroups = _listGraduatedGroups.ToList();

            CbSortDirGraduates.SelectedIndex = 0;
            CbSortFieldGraduates.SelectedIndex = 0;

            CbFilterGenderGraduates.SelectedIndex = 0;
            CbFilterAgeGraduates.SelectedIndex = 0;
            CbFilterTownGraduates.SelectedIndex = 0;
            CbFilterTypeOfFinancingGraduates.SelectedIndex = 0;
            CbFilterEmploymentGraduates.SelectedIndex = 0;

            tbSearchGraduates.Clear();

            UpdateGraduates();
        }

        public void UpdateGraduates()
        {
            _listGraduates = GetContext().Students.Where(x => x.StatusID == 2).ToList();

            if (CbSortFieldGraduates.SelectedIndex > 0)
            {
                switch (CbSortFieldGraduates.SelectedIndex)
                {
                    case 1:
                        if (CbSortDirGraduates.SelectedIndex == 1)
                            _listGraduates = _listGraduates.OrderBy(p => p.Surname).ToList();
                        if (CbSortDirGraduates.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.Surname).ToList();
                        break;
                    case 2:
                        if (CbSortDir.SelectedIndex == 1)
                            _listGraduates = _listGraduates.OrderBy(p => p.DateOfBirth).ToList();
                        if (CbSortDir.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.DateOfBirth).ToList();
                        break;
                    case 3:
                        if (CbSortDir.SelectedIndex == 1)
                            _listGraduates = _listGraduates.OrderBy(p => p.GPA).ToList();
                        if (CbSortDir.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.GPA).ToList();
                        break;
                    case 4:
                        if (CbSortDir.SelectedIndex == 1)
                            _listGraduates = _listGraduates.OrderBy(p => p.Groups.GroupNumber).ToList();
                        if (CbSortDir.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.Groups.GroupNumber).ToList();
                        break;
                    case 5:
                        if (CbSortDir.SelectedIndex == 1)
                            _listGraduates = _listGraduates.OrderBy(p => p.Course).ToList();
                        if (CbSortDir.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.Course).ToList();
                        break;
                }
            }

            if (CbFilterGenderGraduates.SelectedIndex > 0)
            {
                switch (CbFilterGenderGraduates.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.Gender.Equals("Муж.")).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.Gender.Equals("Жен.")).ToList();
                        break;
                }
            }
            if (CbFilterTownGraduates.SelectedIndex > 0)
            {
                switch (CbFilterTown.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.PermanentRegistrationAddress.Contains("Коломна") || p.PermanentRegistrationAddress.Contains("Коломенский")).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.PermanentRegistrationAddress.Contains("Московская")).ToList();
                        break;
                    case 3:
                        _list = _list.Where(p => p.PermanentRegistrationAddress.Contains("Московская") == false).ToList();
                        break;
                }
            }
            if (CbFilterAgeGraduates.SelectedIndex > 0)
            {
                switch (CbFilterAgeGraduates.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.Age < 18).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.Age > 18).ToList();
                        break;
                    case 3:
                        _listGraduates = _listGraduates.Where(p => p.Age <= 14).ToList();
                        break;
                    case 4:
                        _listGraduates = _listGraduates.Where(p => p.Age == 15).ToList();
                        break;
                    case 5:
                        _listGraduates = _listGraduates.Where(p => p.Age == 16).ToList();
                        break;
                    case 6:
                        _listGraduates = _listGraduates.Where(p => p.Age == 17).ToList();
                        break;
                    case 7:
                        _listGraduates = _listGraduates.Where(p => p.Age == 18).ToList();
                        break;
                    case 8:
                        _listGraduates = _listGraduates.Where(p => p.Age == 19).ToList();
                        break;
                    case 9:
                        _listGraduates = _listGraduates.Where(p => p.Age == 20).ToList();
                        break;
                    case 10:
                        _listGraduates = _listGraduates.Where(p => p.Age == 21).ToList();
                        break;
                    case 11:
                        _listGraduates = _listGraduates.Where(p => p.Age == 22).ToList();
                        break;
                    case 12:
                        _listGraduates = _listGraduates.Where(p => p.Age == 23).ToList();
                        break;
                    case 13:
                        _listGraduates = _listGraduates.Where(p => p.Age == 24).ToList();
                        break;
                    case 14:
                        _listGraduates = _listGraduates.Where(p => p.Age == 25).ToList();
                        break;
                    case 15:
                        _listGraduates = _listGraduates.Where(p => p.Age == 26).ToList();
                        break;
                    case 16:
                        _listGraduates = _listGraduates.Where(p => p.Age == 27).ToList();
                        break;
                    case 17:
                        _listGraduates = _listGraduates.Where(p => p.Age == 28).ToList();
                        break;
                    case 18:
                        _listGraduates = _listGraduates.Where(p => p.Age == 29).ToList();
                        break;
                    case 19:
                        _listGraduates = _listGraduates.Where(p => p.Age >= 30 && p.Age <= 34).ToList();
                        break;
                    case 20:
                        _listGraduates = _listGraduates.Where(p => p.Age >= 35 && p.Age <= 39).ToList();
                        break;
                    case 21:
                        _listGraduates = _listGraduates.Where(p => p.Age >= 40).ToList();
                        break;
                }
            }
            if (CbFilterTypeOfFinancingGraduates.SelectedIndex > 0)
            {
                switch (CbFilterTypeOfFinancingGraduates.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.TypeOfFinancing.Equals("Бюджет")).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.TypeOfFinancing.Equals("Договор")).ToList();
                        break;
                }
            }
            if (CbFilterEmploymentGraduates.SelectedIndex > 0)
            {
                switch (CbFilterEmploymentGraduates.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.IsEmployed == true).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.IsEmployed == false).ToList();
                        break;
                }
            }

            _listGraduates = _listGraduates.Where(p => p.Groups.GroupNumber.Equals(CbFilterGroupGraduates.Text)).ToList();

            dgGraduates.ItemsSource = _listGraduates;
            lbCountGraduates.Content = $"Количество записей: {_listGraduates.Count}";
        }

        private void OnSelectionChangedGraduates(object sender, SelectionChangedEventArgs e)
        {
            var selectedItems = dgGraduates.SelectedItems.OfType<int>().ToList();
            lbSelectedGraduates.Content = $"Выбрано: {selectedItems.Count()}";
            int s = 0;
            foreach (var r in dgGraduates.SelectedItems)
            {
                s++;
            }
            lbSelectedGraduates.Content = $"Выбрано: {s}";
        }

        private void TbSearchGraduates_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                _listSearch = _listGraduates.Where(Item => Item.Surname.Contains(tbSearchGraduates.Text) || Item.Name.ToLower().Contains(tbSearchGraduates.Text.ToLower()) 
                || Item.Patronymic.ToLower().Contains(tbSearchGraduates.Text.ToLower()) || Item.HasNote.ToLower().Contains(tbSearch.Text.ToLower())).ToList();
                dgGraduates.ItemsSource = _listSearch;

                lbCountGraduates.Content = $"Количество записей: {_listSearch.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void SortGraduates_Click(object sender, RoutedEventArgs e)
        {
            UpdateGraduates();
        }

        private void SaveGraduates_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Document";
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Worksheets (.xlsx)|*.xlsx";

            bool? result = dlg.ShowDialog();

            if (result == true)
            {
                string filename = dlg.FileName;

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage excelPackage = new ExcelPackage();
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Выборка студентов по фильтру");

                //Заполнение
                try
                {
                    worksheet.Cells[1, 1].Value = "Фамилия";
                    worksheet.Cells[1, 2].Value = "Имя";
                    worksheet.Cells[1, 3].Value = "Отчество";
                    worksheet.Cells[1, 4].Value = "Пол";
                    worksheet.Cells[1, 5].Value = "Дата рождения";
                    worksheet.Cells[1, 6].Value = "Возраст";
                    worksheet.Cells[1, 7].Value = "Телефон";
                    worksheet.Cells[1, 8].Value = "Email";
                    worksheet.Cells[1, 9].Value = "Гражданство";
                    worksheet.Cells[1, 10].Value = "ДУЛ";
                    worksheet.Cells[1, 11].Value = "Серия паспорта";
                    worksheet.Cells[1, 12].Value = "Номер паспорта";
                    worksheet.Cells[1, 13].Value = "Дата выдачи ДУЛа";
                    worksheet.Cells[1, 14].Value = "Кем выдан ДУЛ";
                    worksheet.Cells[1, 15].Value = "Адрес постоянной регистрации";
                    worksheet.Cells[1, 16].Value = "СНИЛС";
                    worksheet.Cells[1, 17].Value = "ИНН";
                    worksheet.Cells[1, 18].Value = "Базовое образование";
                    worksheet.Cells[1, 19].Value = "Средний балл";
                    worksheet.Cells[1, 20].Value = "Является сиротой";
                    worksheet.Cells[1, 21].Value = "Является инвалидом";
                    worksheet.Cells[1, 22].Value = "Причина инвалидности";
                    worksheet.Cells[1, 23].Value = "Форма обучения";
                    worksheet.Cells[1, 24].Value = "Вид финансирования";
                    worksheet.Cells[1, 25].Value = "Специальность";
                    worksheet.Cells[1, 26].Value = "Группа";
                    worksheet.Cells[1, 27].Value = "Статус";
                    worksheet.Cells[1, 28].Value = "Примечания";
                    worksheet.Cells[1, 29].Value = "Трудоустройство";

                    int row = 2;
                    foreach (var item in _listGraduates)
                    {
                        worksheet.Cells[row, 1].Value = item.Surname;
                        worksheet.Cells[row, 2].Value = item.Name;
                        worksheet.Cells[row, 3].Value = item.Patronymic;
                        worksheet.Cells[row, 4].Value = item.Gender;
                        worksheet.Cells[row, 5].Value = item.DateOfBirth;
                        worksheet.Cells[row, 5].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet.Cells[row, 6].Value = item.Age;
                        worksheet.Cells[row, 7].Value = item.Phone;
                        worksheet.Cells[row, 8].Value = item.Email;
                        worksheet.Cells[row, 9].Value = item.Nationality;
                        worksheet.Cells[row, 10].Value = item.IdentityDocument;
                        worksheet.Cells[row, 11].Value = item.PassportSeries;
                        worksheet.Cells[row, 12].Value = item.PassportNumber;
                        worksheet.Cells[row, 13].Value = item.DateOfIssue;
                        worksheet.Cells[row, 13].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet.Cells[row, 14].Value = item.IssuedBy;
                        worksheet.Cells[row, 15].Value = item.PermanentRegistrationAddress;
                        worksheet.Cells[row, 16].Value = item.IIAN;
                        worksheet.Cells[row, 17].Value = item.ITN;
                        worksheet.Cells[row, 18].Value = item.BasicEducation;
                        worksheet.Cells[row, 19].Value = item.GPA;
                        if (item.IsOrphan)
                            worksheet.Cells[row, 20].Value = "Да";
                        else
                            worksheet.Cells[row, 20].Value = "Нет";
                        if (item.IsInvalid)
                            worksheet.Cells[row, 21].Value = "Да";
                        else
                            worksheet.Cells[row, 21].Value = "Нет";
                        worksheet.Cells[row, 22].Value = item.CauseOfDisability;
                        worksheet.Cells[row, 23].Value = item.FormOfStudy;
                        worksheet.Cells[row, 24].Value = item.TypeOfFinancing;
                        worksheet.Cells[row, 25].Value = item.Specialities.SpecialityName;
                        worksheet.Cells[row, 26].Value = item.Groups.GroupNumber;
                        worksheet.Cells[row, 27].Value = item.Statuses.Status;
                        worksheet.Cells[row, 28].Value = item.Note;
                        worksheet.Cells[row, 29].Value = item.IfEmployed;
                        row++;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

                excelPackage.SaveAs(new FileInfo(filename));
                excelPackage.Dispose();
                MessageBox.Show("Excel документ сформирован успешно");
            }
        }

        #endregion

        private void EditStudent_Click(object sender, RoutedEventArgs e)
        {
            Students current = (sender as Button)?.DataContext as Students;
            var wnd = new EditStudentWindow(current);
            wnd.ShowDialog();
            UpLoad();
            UpdateGraduates();
        }

        private void EditEmployment_Click(object sender, RoutedEventArgs e)
        {
            Students current = (sender as Button)?.DataContext as Students;
            var wnd = new EmploymentWindow(current);
            wnd.ShowDialog();
            Update();
            UpdateGraduates();
        }

        private void EmploymentHistory_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new EmploymentHistoryWindow(_currentUser);
            wnd.ShowDialog();
            Update();
            UpdateGraduates();
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new MainWindow();
            wnd.Show();
            Hide();
        }

        private void Help_Click(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start(@"..\..\Reference\reference.chm");
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите завершить работу в программе?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
            {
                e.Cancel = true;
            }
        }
    }
}