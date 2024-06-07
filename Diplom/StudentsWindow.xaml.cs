using iTextSharp.text.pdf;
using iTextSharp.text;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using static Diplom.AppData;
using static System.Net.Mime.MediaTypeNames;
using OfficeOpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Diplom.Models;
using Users = Diplom.Models.Users;
using Groups = Diplom.Models.Groups;
using System.Drawing;
using Font = iTextSharp.text.Font;
using Document = iTextSharp.text.Document;
using Paragraph = iTextSharp.text.Paragraph;
using System.Windows.Input;
using System.Text.RegularExpressions;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для StudentsWindow.xaml
    /// </summary>
    public partial class StudentsWindow : Window
    {
        private static Users _currentUser = new Users();
        private static List<Students> _list = new List<Students>();
        private static List<Students> _listSearch = new List<Students>();
        private static List<Students> _listMy = new List<Students>();
        private static List<Students> _listGraduates = new List<Students>();
        private static List<Groups> _listOutdatedGroups = new List<Groups>();
        private static List<Groups> _listGroups = new List<Groups>();
        private static List<Groups> _listMyGroups = new List<Groups>();
        private static List<Specialities> _listSpecialities = new List<Specialities>();
        private static List<string> listFilterGroup = new List<string>();
        private static List<string> listFilterGroup2 = new List<string>();
        private static List<string> listFilterMyGroup = new List<string>();
        private static List<string> listFilterSpeciality = new List<string>();

        List<Students> _students1 = new List<Students>();
        List<Students> _students2 = new List<Students>();
        List<Students> _students3 = new List<Students>();
        List<Students> _students = new List<Students>();
        List<Groups> _groups1 = new List<Groups>();
        List<Groups> _groups2 = new List<Groups>();
        List<Groups> _groups3 = new List<Groups>();
        List<Groups> _groups = new List<Groups>();

        public StudentsWindow(Users current)
        {
            InitializeComponent();
            _currentUser = current;
            if(_currentUser.TemporaryPassword != null)
            {
                _currentUser.TemporaryPassword = null;
                GetContext().SaveChanges();
            }
            UpLoad();
            Events();
        }

        public void UpLoad()
        {
            _listGroups = GetContext().Groups.ToList();
            _listOutdatedGroups = _listGroups.Where(x => x.GroupCourse > 10).ToList();
            _listMyGroups = GetContext().Groups.Where(x => x.FormMaster == _currentUser.ID).OrderByDescending(x => x.GroupNumber).ToList();
            _listSpecialities = GetContext().Specialities.ToList();
            tbSearch.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            tbSearch1.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);
            tbSearch2.PreviewTextInput += new TextCompositionEventHandler(textBoxText_PreviewTextInput);

            // Формирование выпадающих списков фильтров всех студентов
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
            if (CbFilterCourse.Items.Count == 0)
            {
                CbFilterCourse.Items.Insert(0, "Курс");
                CbFilterCourse.Items.Add("1 курс");
                CbFilterCourse.Items.Add("2 курс");
                CbFilterCourse.Items.Add("3 курс");
                CbFilterCourse.Items.Add("4 курс");
                CbFilterCourse.Items.Add("5 курс");
            }
            if (_listGroups.Count() + 2 != listFilterGroup.Count)
            {
                listFilterGroup = new List<string>();
                listFilterGroup.Insert(0, "Группа");
                listFilterGroup.Insert(1, "Не назначено");
                foreach (var group in _listGroups)
                {
                    listFilterGroup.Add(group.GroupNumber.ToString());
                }
                CbFilterGroup.ItemsSource = listFilterGroup;
            }
            if (CbFilterSpeciality.Items.Count == 0)
            {
                listFilterSpeciality = new List<string>();
                listFilterSpeciality.Insert(0, "Специальность");
                foreach (var speciality in _listSpecialities)
                {
                    listFilterSpeciality.Add(speciality.SpecialityName.ToString());
                }
                CbFilterSpeciality.ItemsSource = listFilterSpeciality;
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
            if (CbFilterFormOfStudy.Items.Count == 0)
            {
                CbFilterFormOfStudy.Items.Insert(0, "Форма обучения");
                CbFilterFormOfStudy.Items.Add("Очная");
                CbFilterFormOfStudy.Items.Add("Заочная");
            }
            if (CbFilterTypeOfFinancing.Items.Count == 0)
            {
                CbFilterTypeOfFinancing.Items.Insert(0, "Вид финансирования");
                CbFilterTypeOfFinancing.Items.Add("Бюджет");
                CbFilterTypeOfFinancing.Items.Add("Договор");
            }
            if (CbFilterIsOrphan.Items.Count == 0)
            {
                CbFilterIsOrphan.Items.Insert(0, "Является сиротой");
                CbFilterIsOrphan.Items.Add("Да");
                CbFilterIsOrphan.Items.Add("Нет");
            }
            if (CbFilterIsInvalid.Items.Count == 0)
            {
                CbFilterIsInvalid.Items.Insert(0, "Является инвалидом");
                CbFilterIsInvalid.Items.Add("Да");
                CbFilterIsInvalid.Items.Add("Нет");
            }

            #endregion

            // Формирование выпадающих списков фильтров выпускников
            #region

            if (CbSortDir2.Items.Count == 0)
            {
                CbSortDir2.Items.Insert(0, "Направление сортировки");
                CbSortDir2.Items.Add("По возрастанию");
                CbSortDir2.Items.Add("По убыванию");
            }

            if (CbSortField2.Items.Count == 0)
            {
                CbSortField2.Items.Insert(0, "Поле для сортировки");
                CbSortField2.Items.Add("Фамилия");
                CbSortField2.Items.Add("Дата рождения");
                CbSortField2.Items.Add("Средний балл");
                CbSortField2.Items.Add("Группа");
            }

            if (CbFilterGender2.Items.Count == 0)
            {
                CbFilterGender2.Items.Insert(0, "Пол");
                CbFilterGender2.Items.Add("Мужской");
                CbFilterGender2.Items.Add("Женский");
            }

            if (CbFilterTown2.Items.Count == 0)
            {
                CbFilterTown2.Items.Insert(0, "Город");
                CbFilterTown2.Items.Add("Коломна, Коломенский р-н");
                CbFilterTown2.Items.Add("Московская обл.");
                CbFilterTown2.Items.Add("Прочие регионы");
            }

            if (_listGroups.Count() + 1 != listFilterGroup2.Count)
            {
                listFilterGroup2 = new List<string>();
                listFilterGroup2.Insert(0, "Группа");
                foreach (var group in _listGroups)
                {
                    listFilterGroup2.Add(group.GroupNumber.ToString());
                }
                CbFilterGroup2.ItemsSource = listFilterGroup2;
            }

            if (CbFilterSpeciality2.Items.Count == 0)
            {
                listFilterSpeciality = new List<string>();
                listFilterSpeciality.Insert(0, "Специальность");
                foreach (var speciality in _listSpecialities)
                {
                    listFilterSpeciality.Add(speciality.SpecialityName.ToString());
                }
                CbFilterSpeciality2.ItemsSource = listFilterSpeciality;
            }

            if (CbFilterEmployment2.Items.Count == 0)
            {
                CbFilterEmployment2.Items.Insert(0, "Трудоустройство");
                CbFilterEmployment2.Items.Add("Трудоустроен");
                CbFilterEmployment2.Items.Add("Не трудоустроен");
            }

            if (CbFilterFormOfStudy2.Items.Count == 0)
            {
                CbFilterFormOfStudy2.Items.Insert(0, "Форма обучения");
                CbFilterFormOfStudy2.Items.Add("Очная");
                CbFilterFormOfStudy2.Items.Add("Заочная");
            }

            if (CbFilterTypeOfFinancing2.Items.Count == 0)
            {
                CbFilterTypeOfFinancing2.Items.Insert(0, "Вид финансирования");
                CbFilterTypeOfFinancing2.Items.Add("Бюджет");
                CbFilterTypeOfFinancing2.Items.Add("Договор");
            }

            if (CbFilterIsOrphan2.Items.Count == 0)
            {
                CbFilterIsOrphan2.Items.Insert(0, "Является сиротой");
                CbFilterIsOrphan2.Items.Add("Да");
                CbFilterIsOrphan2.Items.Add("Нет");
            }

            if (CbFilterIsInvalid2.Items.Count == 0)
            {
                CbFilterIsInvalid2.Items.Insert(0, "Является инвалидом");
                CbFilterIsInvalid2.Items.Add("Да");
                CbFilterIsInvalid2.Items.Add("Нет");
            }

            #endregion

            // Формирование выпадающих списков фильтров студентов подчиненной группы
            #region
            if (CbSortDir1.Items.Count == 0)
            {
                CbSortDir1.Items.Insert(0, "Направление сортировки");
                CbSortDir1.Items.Add("По возрастанию");
                CbSortDir1.Items.Add("По убыванию");
            }

            if (CbSortField1.Items.Count == 0)
            {
                CbSortField1.Items.Insert(0, "Поле для сортировки");
                CbSortField1.Items.Add("Фамилия");
                CbSortField1.Items.Add("Дата рождения");
                CbSortField1.Items.Add("Средний балл");
            }

            if (CbFilterTown1.Items.Count == 0)
            {
                CbFilterTown1.Items.Insert(0, "Город");
                CbFilterTown1.Items.Add("Коломна, Коломенский р-н");
                CbFilterTown1.Items.Add("Московская обл.");
                CbFilterTown1.Items.Add("Прочие регионы");
            }

            if (CbFilterAge1.Items.Count == 0)
            {
                CbFilterAge1.Items.Insert(0, "Возраст");
                CbFilterAge1.Items.Add("Несовершеннолетние");
                CbFilterAge1.Items.Add("Совершеннолетние");
                CbFilterAge1.Items.Add("14 и младше");
                CbFilterAge1.Items.Add("15");
                CbFilterAge1.Items.Add("16");
                CbFilterAge1.Items.Add("17");
                CbFilterAge1.Items.Add("18");
                CbFilterAge1.Items.Add("19");
                CbFilterAge1.Items.Add("20");
                CbFilterAge1.Items.Add("21");
                CbFilterAge1.Items.Add("22");
                CbFilterAge1.Items.Add("23");
                CbFilterAge1.Items.Add("24");
                CbFilterAge1.Items.Add("25");
                CbFilterAge1.Items.Add("26");
                CbFilterAge1.Items.Add("27");
                CbFilterAge1.Items.Add("28");
                CbFilterAge1.Items.Add("28");
                CbFilterAge1.Items.Add("29");
                CbFilterAge1.Items.Add("30-34");
                CbFilterAge1.Items.Add("35-39");
                CbFilterAge1.Items.Add("40 и старше");
            }

            if (CbFilterTypeOfFinancing1.Items.Count == 0)
            {
                CbFilterTypeOfFinancing1.Items.Insert(0, "Вид финансирования");
                CbFilterTypeOfFinancing1.Items.Add("Бюджет");
                CbFilterTypeOfFinancing1.Items.Add("Договор");
            }

            // Обновление списка групп для классного руководителя
            if (CbFilterGroup1.Items.Count == 0)
            {
                listFilterMyGroup = new List<string>();
                foreach (var group in _listMyGroups)
                {
                    listFilterMyGroup.Add(group.GroupNumber.ToString());
                }
                CbFilterGroup1.ItemsSource = listFilterMyGroup;
            }
            if (CbFilterEmployment1.Items.Count == 0)
            {
                CbFilterEmployment1.Items.Insert(0, "Трудоустройство");
                CbFilterEmployment1.Items.Add("Трудоустроен");
                CbFilterEmployment1.Items.Add("Не трудоустроен");
            }
            #endregion

            // Проверка данных на актуальность
            if (_listOutdatedGroups.Count > 0)
            {
                if (MessageBox.Show("В системе обнаружены неактуальные данные: найдены записи студентов, которые были выпущены более 5 лет назад. \nУдалить данных студентов?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    foreach (Groups group in _listOutdatedGroups)
                    {
                        List<Students> _listOutdatedStudents = GetContext().Students.Where(x => x.GroupID == group.ID).ToList();
                        foreach (Students student in _listOutdatedStudents)
                        {
                            List<Employment> studentsEmployment = GetContext().Employment.Where(p => p.StudentID == student.ID).ToList();
                            if (studentsEmployment.Count > 0)
                            {
                                GetContext().Employment.RemoveRange(studentsEmployment);
                            }
                        }
                        GetContext().Students.RemoveRange(_listOutdatedStudents);
                    }
                    if (MessageBox.Show("Удалить группы, в которых учились удаленные студенты?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        GetContext().Groups.RemoveRange(_listOutdatedGroups);
                    }
                    GetContext().SaveChanges();
                }
            }

            Update();
        }

        void textBoxText_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            if (!Regex.Match(e.Text, @"[а-яА-Яё]").Success) e.Handled = true;
        }

        public void Update()  // Для всех студентов
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
            if (CbFilterIsOrphan.SelectedIndex > 0)
            {
                switch (CbFilterIsOrphan.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.IsOrphan == true).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.IsOrphan == false).ToList();
                        break;
                }
            }
            if (CbFilterIsInvalid.SelectedIndex > 0)
            {
                switch (CbFilterIsInvalid.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.IsInvalid == true).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.IsInvalid == false).ToList();
                        break;
                }
            }
            if (CbFilterFormOfStudy.SelectedIndex > 0)
            {
                switch (CbFilterFormOfStudy.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.FormOfStudy.Equals("Очная")).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.FormOfStudy.Equals("Заочная")).ToList();
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
            if (CbFilterSpeciality.SelectedIndex > 0)
            {
                _list = _list.Where(p => p.Specialities.SpecialityName.Equals(CbFilterSpeciality.Text)).ToList();
            }
            if (CbFilterGroup.SelectedIndex > 0)
            {
                if (CbFilterGroup.SelectedIndex == 1)
                {
                    _list = _list.Where(p => p.GroupID == null).ToList();
                }
                else
                {
                    _list = _list.Where(p => p.GroupID != null).ToList();
                    _list = _list.Where(p => p.Groups.GroupNumber.Equals(CbFilterGroup.Text)).ToList();
                }
            }
            if (CbFilterCourse.SelectedIndex > 0)
            {
                switch (CbFilterCourse.SelectedIndex)
                {
                    case 1:
                        _list = _list.Where(p => p.Course == 1).ToList();
                        break;
                    case 2:
                        _list = _list.Where(p => p.Course == 2).ToList();
                        break;
                    case 3:
                        _list = _list.Where(p => p.Course == 3).ToList();
                        break;
                    case 4:
                        _list = _list.Where(p => p.Course == 4).ToList();
                        break;
                    case 5:
                        _list = _list.Where(p => p.Course == 5).ToList();
                        break;
                }
            }

            dgStudents.ItemsSource = _list;
            lbCount.Content = $"Количество записей: {_list.Count}";
        }

        public void UpdateGraduates()  // Для выпускников
        {
            _listGraduates = GetContext().Students.Where(x => x.StatusID == 2).ToList();

            if (CbSortField2.SelectedIndex > 0)
            {
                switch (CbSortField2.SelectedIndex)
                {
                    case 1:
                        if (CbSortDir2.SelectedIndex == 1)
                            _listGraduates = _listGraduates.OrderBy(p => p.Surname).ToList();
                        if (CbSortDir2.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.Surname).ToList();
                        break;
                    case 2:
                        if (CbSortDir2.SelectedIndex == 1)
                            _listGraduates = _listGraduates.OrderBy(p => p.DateOfBirth).ToList();
                        if (CbSortDir2.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.DateOfBirth).ToList();
                        break;
                    case 3:
                        if (CbSortDir2.SelectedIndex == 1)
                            _list = _list.OrderBy(p => p.GPA).ToList();
                        if (CbSortDir2.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.GPA).ToList();
                        break;
                    case 4:
                        if (CbSortDir2.SelectedIndex == 1)
                            _listGraduates = _listGraduates.OrderBy(p => p.Groups.GroupNumber).ToList();
                        if (CbSortDir2.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.Groups.GroupNumber).ToList();
                        break;
                    case 5:
                        if (CbSortDir2.SelectedIndex == 1)
                            _listGraduates = _listGraduates.OrderBy(p => p.Course).ToList();
                        if (CbSortDir2.SelectedIndex == 2)
                            _listGraduates = _listGraduates.OrderByDescending(p => p.Course).ToList();
                        break;
                }
            }
            if (CbFilterGender2.SelectedIndex > 0)
            {
                switch (CbFilterGender2.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.Gender.Equals("Муж.")).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.Gender.Equals("Жен.")).ToList();
                        break;
                }
            }
            if (CbFilterTown2.SelectedIndex > 0)
            {
                switch (CbFilterTown2.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.PermanentRegistrationAddress.Contains("Коломна") || p.PermanentRegistrationAddress.Contains("Коломенский")).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.PermanentRegistrationAddress.Contains("Московская")).ToList();
                        break;
                    case 3:
                        _listGraduates = _listGraduates.Where(p => p.PermanentRegistrationAddress.Contains("Московская") == false).ToList();
                        break;
                }
            }
            if (CbFilterEmployment2.SelectedIndex > 0)
            {
                switch (CbFilterEmployment2.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.IsEmployed == true).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.IsEmployed == false).ToList();
                        break;
                }
            }
            if (CbFilterIsOrphan2.SelectedIndex > 0)
            {
                switch (CbFilterIsOrphan2.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.IsOrphan == true).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.IsOrphan == false).ToList();
                        break;
                }
            }
            if (CbFilterIsInvalid2.SelectedIndex > 0)
            {
                switch (CbFilterIsInvalid2.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.IsInvalid == true).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.IsInvalid == false).ToList();
                        break;
                }
            }
            if (CbFilterFormOfStudy2.SelectedIndex > 0)
            {
                switch (CbFilterFormOfStudy2.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.FormOfStudy.Equals("Очная")).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.FormOfStudy.Equals("Заочная")).ToList();
                        break;
                }
            }
            if (CbFilterTypeOfFinancing2.SelectedIndex > 0)
            {
                switch (CbFilterTypeOfFinancing2.SelectedIndex)
                {
                    case 1:
                        _listGraduates = _listGraduates.Where(p => p.TypeOfFinancing.Equals("Бюджет")).ToList();
                        break;
                    case 2:
                        _listGraduates = _listGraduates.Where(p => p.TypeOfFinancing.Equals("Договор")).ToList();
                        break;
                }
            }
            if (CbFilterSpeciality2.SelectedIndex > 0)
            {
                _listGraduates = _listGraduates.Where(p => p.Specialities.SpecialityName.Equals(CbFilterSpeciality2.Text)).ToList();
            }
            if (CbFilterGroup2.SelectedIndex > 0)
            {
                if (CbFilterGroup2.SelectedIndex == 1)
                {
                    _listGraduates = _listGraduates.Where(p => p.GroupID == null).ToList();
                }
                else
                {
                    _listGraduates = _listGraduates.Where(p => p.GroupID != null).ToList();
                    _listGraduates = _listGraduates.Where(p => p.Groups.GroupNumber.Equals(CbFilterGroup2.Text)).ToList();
                }
            }

            dgGraduates.ItemsSource = _listGraduates;
            lbCount2.Content = $"Количество записей: {_listGraduates.Count}";
        }

        public void UpdateMyGroup() // Для студентов подчиненных групп
        {
            _listMy = GetContext().Students.Where(x => x.StatusID == 1 && x.Groups.FormMaster == _currentUser.ID).ToList();

            if (CbSortField1.SelectedIndex > 0)
            {
                switch (CbSortField1.SelectedIndex)
                {
                    case 1:
                        if (CbSortDir1.SelectedIndex == 1)
                            _listMy = _listMy.OrderBy(p => p.Surname).ToList();
                        if (CbSortDir1.SelectedIndex == 2)
                            _listMy = _listMy.OrderByDescending(p => p.Surname).ToList();
                        break;
                    case 2:
                        if (CbSortDir1.SelectedIndex == 1)
                            _listMy = _listMy.OrderBy(p => p.DateOfBirth).ToList();
                        if (CbSortDir1.SelectedIndex == 2)
                            _listMy = _listMy.OrderByDescending(p => p.DateOfBirth).ToList();
                        break;
                    case 3:
                        if (CbSortDir1.SelectedIndex == 1)
                            _listMy = _listMy.OrderBy(p => p.GPA).ToList();
                        if (CbSortDir1.SelectedIndex == 2)
                            _listMy = _listMy.OrderByDescending(p => p.GPA).ToList();
                        break;
                }
            }
            if (CbFilterTown1.SelectedIndex > 0)
            {
                switch (CbFilterTown1.SelectedIndex)
                {
                    case 1:
                        _listMy = _listMy.Where(p => p.PermanentRegistrationAddress.Contains("Коломна") || p.PermanentRegistrationAddress.Contains("Коломенский")).ToList();
                        break;
                    case 2:
                        _listMy = _listMy.Where(p => p.PermanentRegistrationAddress.Contains("Московская")).ToList();
                        break;
                    case 3:
                        _listMy = _listMy.Where(p => p.PermanentRegistrationAddress.Contains("Московская") == false).ToList();
                        break;
                }
            }
            if (CbFilterAge1.SelectedIndex > 0)
            {
                switch (CbFilterAge1.SelectedIndex)
                {
                    case 1:
                        _listMy = _listMy.Where(p => p.Age < 18).ToList();
                        break;
                    case 2:
                        _listMy = _listMy.Where(p => p.Age > 18).ToList();
                        break;
                    case 3:
                        _listMy = _listMy.Where(p => p.Age <= 14).ToList();
                        break;
                    case 4:
                        _listMy = _listMy.Where(p => p.Age == 15).ToList();
                        break;
                    case 5:
                        _listMy = _listMy.Where(p => p.Age == 16).ToList();
                        break;
                    case 6:
                        _listMy = _listMy.Where(p => p.Age == 17).ToList();
                        break;
                    case 7:
                        _listMy = _listMy.Where(p => p.Age == 18).ToList();
                        break;
                    case 8:
                        _listMy = _listMy.Where(p => p.Age == 19).ToList();
                        break;
                    case 9:
                        _listMy = _listMy.Where(p => p.Age == 20).ToList();
                        break;
                    case 10:
                        _listMy = _listMy.Where(p => p.Age == 21).ToList();
                        break;
                    case 11:
                        _listMy = _listMy.Where(p => p.Age == 22).ToList();
                        break;
                    case 12:
                        _listMy = _listMy.Where(p => p.Age == 23).ToList();
                        break;
                    case 13:
                        _listMy = _listMy.Where(p => p.Age == 24).ToList();
                        break;
                    case 14:
                        _listMy = _listMy.Where(p => p.Age == 25).ToList();
                        break;
                    case 15:
                        _listMy = _listMy.Where(p => p.Age == 26).ToList();
                        break;
                    case 16:
                        _listMy = _listMy.Where(p => p.Age == 27).ToList();
                        break;
                    case 17:
                        _listMy = _listMy.Where(p => p.Age == 28).ToList();
                        break;
                    case 18:
                        _listMy = _listMy.Where(p => p.Age == 29).ToList();
                        break;
                    case 19:
                        _listMy = _listMy.Where(p => p.Age >= 30 && p.Age <= 34).ToList();
                        break;
                    case 20:
                        _listMy = _listMy.Where(p => p.Age >= 35 && p.Age <= 39).ToList();
                        break;
                    case 21:
                        _listMy = _listMy.Where(p => p.Age >= 40).ToList();
                        break;
                }
            }
            if (CbFilterTypeOfFinancing1.SelectedIndex > 0)
            {
                switch (CbFilterTypeOfFinancing1.SelectedIndex)
                {
                    case 1:
                        _listMy = _listMy.Where(p => p.TypeOfFinancing.Equals("Бюджет")).ToList();
                        break;
                    case 2:
                        _listMy = _listMy.Where(p => p.TypeOfFinancing.Equals("Договор")).ToList();
                        break;
                }
            }
            if (CbFilterEmployment1.SelectedIndex > 0)
            {
                switch (CbFilterEmployment1.SelectedIndex)
                {
                    case 1:
                        _listMy = _listMy.Where(p => p.IsEmployed == true).ToList();
                        break;
                    case 2:
                        _listMy = _listMy.Where(p => p.IsEmployed == false).ToList();
                        break;
                }
            }
            _listMy = _listMy.Where(p => p.Groups.GroupNumber.Equals(CbFilterGroup1.Text)).ToList();

            dgMyStudents.ItemsSource = _listMy;
            lbCount1.Content = $"Количество записей: {_listMy.Count}";
        }

        // Слои окна
        #region
        private void Events()
        {
            btnAllStudents.Click += (s, e) =>
            {
                Clear();
                dgMyStudents.SelectedItems.Clear();
                dgGraduates.SelectedItems.Clear();
                SwitchLayers(nameof(studentsGrid));
                Update();
            };

            btnGraduates.Click += (s, e) =>
            {
                ClearGraduates();
                dgMyStudents.SelectedItems.Clear();
                dgStudents.SelectedItems.Clear();
                SwitchLayers(nameof(graduatesGrid));
                UpdateGraduates();
            };

            btnMyStudents.Click += (s, e) =>
            {
                ClearMyGroup();
                dgStudents.SelectedItems.Clear();
                dgGraduates.SelectedItems.Clear();
                SwitchLayers(nameof(myGroupsGrid));
                UpdateMyGroup();
            };
        }

        private void SwitchLayers(string LayerName)
        {
            List<Grid> layers = new List<Grid>()
            {
                studentsGrid,
                graduatesGrid,
                myGroupsGrid
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

        private void MyGroups_Click(object sender, RoutedEventArgs e)
        {
            SwitchLayers(nameof(myStudents));
        }
        #endregion

        // Действия
        #region
        private void Expel_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = dgStudents.SelectedItems;
            if (selectedItems.Count != 0)
            {
                if (MessageBox.Show("Вы уверены, что хотите отчислить выбранных студентов?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    try
                    {
                        foreach(Students student in selectedItems)
                        {
                            var studentsForExpelling = GetContext().Students.Where(p => p.ID == student.ID).First();
                            List<Employment> studentsEmployment = GetContext().Employment.Where(p => p.StudentID == student.ID).ToList();
                            if (studentsEmployment.Count > 0)
                            {
                                GetContext().Employment.RemoveRange(studentsEmployment);
                            }
                            GetContext().Students.Remove(studentsForExpelling);
                        }
                        GetContext().SaveChanges();
                        MessageBox.Show("Студенты отчислены");
                        UpLoad();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
            }
            else
            {
                MessageBox.Show("Студенты не выбраны", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SetGroup_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = dgStudents.SelectedItems;
            if (selectedItems.Count != 0)
            {
                var wnd = new SetGroupWindow(selectedItems);
                wnd.ShowDialog();
                UpLoad();
            }
            else
            {
                MessageBox.Show("Студенты не выбраны", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            Update();
        }

        private void NextCource_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _students = GetContext().Students.Where(x => x.StatusID == 1).ToList();
                _groups = GetContext().Groups.ToList();

                _students1 = _students.Where(x => x.Course == 1).OrderBy(p => p.Groups.GroupNumber).ThenBy(x => x.Surname).ToList();
                _students2 = _students.Where(x => x.Course == 2).OrderBy(p => p.Groups.GroupNumber).ThenBy(x => x.Surname).ToList();
                _students3 = _students.Where(x => x.Course == 3).OrderBy(p => p.Groups.GroupNumber).ThenBy(x => x.Surname).ToList();

                _groups1 = _groups.Where(x => x.GroupCourse == 1).OrderBy(p => p.GroupNumber).ToList();
                _groups2 = _groups.Where(x => x.GroupCourse == 2).OrderBy(p => p.GroupNumber).ToList();
                _groups3 = _groups.Where(x => x.GroupCourse == 3).OrderBy(p => p.GroupNumber).ToList();

                DateTime today = DateTime.Today;
                Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog
                {
                    FileName = $"ПЕРЕВОДНОЙ общий {today.Year}",
                    DefaultExt = ".pdf",
                    Filter = "Text documents (.pdf)|*.pdf"
                };

                bool? result = dlg.ShowDialog();

                if (result == true)
                {
                    string filename = dlg.FileName;

                    Document document = new Document();
                    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(filename, FileMode.Create));

                    string ttf = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "TIMES.TTF");
                    var baseFont = BaseFont.CreateFont(ttf, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                    Font font = new Font(baseFont, Font.DEFAULTSIZE, Font.NORMAL);
                    Font font1 = new Font(baseFont, Font.DEFAULTSIZE, Font.ITALIC);
                    Font font2 = new Font(baseFont, 14, Font.BOLD | Font.UNDERLINE);
                    Font font3 = new Font(baseFont, Font.DEFAULTSIZE, Font.BOLD);
                    Font font4 = new Font(baseFont, 11, Font.NORMAL);

                    StringFormat stringFormat = new StringFormat();
                    stringFormat.Alignment = StringAlignment.Center;
                    stringFormat.LineAlignment = StringAlignment.Center;

                    string orderN = "______";
                    string header = "Государственное бюджетное профессиональное образовательное учреждение\r\nМосковской области «Колледж «Коломна»\r\n\r\nПРИКАЗ";
                    string orderDN = $"«___» ______ {today.Year} года                                                                                            №{orderN}\r\n\r\n";
                    string aboutTransfer = "О переводе студентов очного отделения\r\n";
                    string basedOn = $"\tНа основании решения педагогического совета от __________ года, протокол №___\r\n";
                    string order = "ПРИКАЗЫВАЮ:";
                    string director = "\r\n\r\n\r\nДиректор колледжа                                                             М. А. Ширкалин";

                    string transferCource2 = $"\r\n1. Перевести на 2 курс:";
                    string transferCource3 = $"\r\n2. Перевести на 3 курс:";
                    string transferCource4 = $"\r\n3. Перевести на 4 курс:";
                    string nn = "№ п/п                  ФИО";
                    string contract = "На платной основе";
                    string zav = "\r\n\r\nАбрамова Ольга Ивановна,\r\nЕмельянова Вера Анатольевна,\r\nЗав СП4,\r\nт. 8-496-618-09-58.";

                    using (writer)
                    {
                        document.Open();
                        document.NewPage();
                        document.Add(new Paragraph(header, font)
                        {
                            Alignment = (int)TabStop.Alignment.RIGHT
                        });
                        document.Add(new Paragraph(orderDN, font)
                        {
                            Alignment = (int)TabStop.Alignment.RIGHT
                        });
                        document.Add(new Paragraph(aboutTransfer, font1)
                        {
                            Alignment = (int)TabStop.Alignment.RIGHT
                        });
                        document.Add(new Paragraph(basedOn, font)
                        {
                            Alignment = (int)TabStop.Alignment.ANCHOR
                        });
                        document.Add(new Paragraph(order, font));

                        document.Add(new Paragraph(transferCource2, font));

                        int index;
                        string transferGroup, group;

                        foreach (Groups gr in _groups1)
                        {
                            group = gr.GroupNumber;
                            transferGroup = $"В группу {group}:";
                            document.Add(new Paragraph(transferGroup, font2));
                            document.Add(new Paragraph(nn, font3));
                            index = 1;
                            int contractCount = _students1.Where(x => x.TypeOfFinancing == "Договор").Count();
                            foreach (Students st in _students1)
                            {
                                if (st.Groups.GroupNumber == group && st.TypeOfFinancing == "Бюджет")
                                {
                                    ChangeCaseOfNames changeCaseOfNames = new ChangeCaseOfNames(st.Name, st.Surname, st.Patronymic, st.Gender);
                                    changeCaseOfNames.ChangeCase();
                                    string str = $"{index}          {changeCaseOfNames.GetSurame} {changeCaseOfNames.GetName} {changeCaseOfNames.GetPatronymic}";
                                    document.Add(new Paragraph(str, font));
                                    index++;
                                }
                            }
                            if(contractCount > 0)
                            {
                                document.Add(new Paragraph(contract, font3));
                                foreach (Students st in _students1)
                                {
                                    if (st.Groups.GroupNumber == group && st.TypeOfFinancing == "Договор")
                                    {
                                        ChangeCaseOfNames changeCaseOfNames = new ChangeCaseOfNames(st.Name, st.Surname, st.Patronymic, st.Gender);
                                        changeCaseOfNames.ChangeCase();
                                        string str = $"{index}          {changeCaseOfNames.GetSurame} {changeCaseOfNames.GetName} {changeCaseOfNames.GetPatronymic}";
                                        document.Add(new Paragraph(str, font));
                                        index++;
                                    }
                                }
                            }
                            document.Add(new Paragraph(" ", font));
                        }

                        document.Add(new Paragraph(transferCource3, font));
                        foreach (Groups gr in _groups2)
                        {
                            group = gr.GroupNumber;
                            transferGroup = $"В группу {group}:";
                            document.Add(new Paragraph(transferGroup, font2));
                            document.Add(new Paragraph(nn, font3));
                            index = 1;
                            int contractCount = _students2.Where(x => x.TypeOfFinancing == "Договор").Count();
                            foreach (Students st in _students2)
                            {
                                if (st.Groups.GroupNumber == group && st.TypeOfFinancing == "Бюджет")
                                {
                                    ChangeCaseOfNames changeCaseOfNames = new ChangeCaseOfNames(st.Name, st.Surname, st.Patronymic, st.Gender);
                                    changeCaseOfNames.ChangeCase();
                                    string str = $"{index}          {changeCaseOfNames.GetSurame} {changeCaseOfNames.GetName} {changeCaseOfNames.GetPatronymic}";
                                    document.Add(new Paragraph(str, font));
                                    index++;
                                }
                            }
                            if (contractCount > 0)
                            {
                                document.Add(new Paragraph(contract, font3));
                                foreach (Students st in _students2)
                                {
                                    if (st.Groups.GroupNumber == group && st.TypeOfFinancing == "Договор")
                                    {
                                        ChangeCaseOfNames changeCaseOfNames = new ChangeCaseOfNames(st.Name, st.Surname, st.Patronymic, st.Gender);
                                        changeCaseOfNames.ChangeCase();
                                        string str = $"{index}          {changeCaseOfNames.GetSurame} {changeCaseOfNames.GetName} {changeCaseOfNames.GetPatronymic}";
                                        document.Add(new Paragraph(str, font));
                                        index++;
                                    }
                                }
                            }
                            document.Add(new Paragraph(" ", font));
                        }

                        document.Add(new Paragraph(transferCource4, font));
                        foreach (Groups gr in _groups3)
                        {
                            group = gr.GroupNumber;
                            transferGroup = $"В группу {group}:";
                            document.Add(new Paragraph(transferGroup, font2));
                            document.Add(new Paragraph(nn, font3));
                            index = 1;
                            int contractCount = _students3.Where(x => x.TypeOfFinancing == "Договор").Count();
                            foreach (Students st in _students3)
                            {
                                if (st.Groups.GroupNumber == group && st.TypeOfFinancing == "Бюджет")
                                {
                                    ChangeCaseOfNames changeCaseOfNames = new ChangeCaseOfNames(st.Name, st.Surname, st.Patronymic, st.Gender);
                                    changeCaseOfNames.ChangeCase();
                                    string str = $"{index}          {changeCaseOfNames.GetSurame} {changeCaseOfNames.GetName} {changeCaseOfNames.GetPatronymic}";
                                    document.Add(new Paragraph(str, font));
                                    index++;
                                }
                            }
                            if (contractCount > 0)
                            {
                                document.Add(new Paragraph(contract, font3));
                                foreach (Students st in _students3)
                                {
                                    if (st.Groups.GroupNumber == group && st.TypeOfFinancing == "Договор")
                                    {
                                        ChangeCaseOfNames changeCaseOfNames = new ChangeCaseOfNames(st.Name, st.Surname, st.Patronymic, st.Gender);
                                        changeCaseOfNames.ChangeCase();
                                        string str = $"{index}          {changeCaseOfNames.GetSurame} {changeCaseOfNames.GetName} {changeCaseOfNames.GetPatronymic}";
                                        document.Add(new Paragraph(str, font));
                                        index++;
                                    }
                                }
                            }
                            document.Add(new Paragraph(" ", font));
                        }
                        document.Add(new Paragraph(director, font)
                        {
                            Alignment = (int)TabStop.Alignment.RIGHT
                        });
                        document.Add(new Paragraph(zav, font4));


                        document.Close();
                    }
                    MessageBox.Show("Приказ сформирован");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void EditEmployment_Click(object sender, RoutedEventArgs e)
        {
            Students current = (sender as Button)?.DataContext as Students;
            var wnd = new EmploymentWindow(current);
            wnd.ShowDialog();
            Update();
            UpdateGraduates();
            UpdateMyGroup();
        }
        #endregion

        // Students
        #region
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
                _listSearch = _list.Where(Item => Item.Surname.ToLower().Contains(tbSearch.Text.ToLower()) || Item.Name.ToLower().Contains(tbSearch.Text.ToLower())
                    || Item.Patronymic.ToLower().Contains(tbSearch.Text.ToLower()) || Item.HasNote.ToLower().Contains(tbSearch.Text.ToLower())).ToList();
                dgStudents.ItemsSource = _listSearch;

                lbCount.Content = $"Количество записей: {_listSearch.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void Import_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new ImportWindow();
            wnd.ShowDialog();
            Update();
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog
            {
                FileName = "Document",
                DefaultExt = ".xlsx",
                Filter = "Excel Worksheets (.xlsx)|*.xlsx"
            };

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
                    worksheet.Cells[1, 6].AutoFitColumns();
                    worksheet.Cells[1, 7].Value = "Телефон";
                    worksheet.Cells[1, 8].Value = "Email";
                    worksheet.Cells[1, 9].Value = "Гражданство";
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
                    worksheet.Cells[1, 27].Value = "Классный руководитель";
                    worksheet.Cells[1, 27].AutoFitColumns();
                    worksheet.Cells[1, 28].Value = "Статус";
                    worksheet.Cells[1, 29].Value = "Примечания";
                    worksheet.Cells[1, 29].AutoFitColumns();
                    worksheet.Cells[1, 30].Value = "Трудоустройство";
                    worksheet.Cells[1, 30].AutoFitColumns();

                    int row = 2;
                    foreach (var item in _list)
                    {
                        worksheet.Cells[row, 1].Value = item.Surname;
                        worksheet.Cells[row, 1].AutoFitColumns();
                        worksheet.Cells[row, 2].Value = item.Name;
                        worksheet.Cells[row, 2].AutoFitColumns();
                        worksheet.Cells[row, 3].Value = item.Patronymic;
                        worksheet.Cells[row, 3].AutoFitColumns();
                        worksheet.Cells[row, 4].Value = item.Gender;
                        worksheet.Cells[row, 5].Value = item.DateOfBirth;
                        worksheet.Cells[row, 5].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet.Cells[row, 6].Value = item.Age;
                        worksheet.Cells[row, 7].Value = item.Phone;
                        worksheet.Cells[row, 7].AutoFitColumns();
                        worksheet.Cells[row, 8].Value = item.Email;
                        worksheet.Cells[row, 8].AutoFitColumns();
                        worksheet.Cells[row, 9].Value = item.Nationality;
                        worksheet.Cells[row, 9].AutoFitColumns();
                        worksheet.Cells[row, 10].Value = item.IdentityDocument;
                        worksheet.Cells[row, 10].AutoFitColumns();
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
                        if(item.IsOrphan)
                            worksheet.Cells[row, 20].Value = "Да";
                        else
                            worksheet.Cells[row, 20].Value = "Нет";
                        if(item.IsInvalid)
                            worksheet.Cells[row, 21].Value = "Да";
                        else
                            worksheet.Cells[row, 21].Value = "Нет";
                        worksheet.Cells[row, 22].Value = item.CauseOfDisability;
                        worksheet.Cells[row, 22].Style.WrapText = true;
                        worksheet.Cells[row, 23].Value = item.FormOfStudy;
                        worksheet.Cells[row, 24].Value = item.TypeOfFinancing;
                        worksheet.Cells[row, 25].Value = item.Specialities.SpecialityName;
                        worksheet.Cells[row, 25].Style.WrapText = true;
                        if (item.GroupID == null)
                        {
                            worksheet.Cells[row, 26].Value = "На распределении";
                        }
                        else
                        {
                            worksheet.Cells[row, 26].Value = item.Groups.GroupNumber;
                            worksheet.Cells[row, 27].Value = item.Groups.Users.Surname + " " + item.Groups.Users.Name + " " + item.Groups.Users.Patronymic;
                        }
                        worksheet.Cells[row, 28].Value = item.Statuses.Status;
                        worksheet.Cells[row, 28].AutoFitColumns();
                        if(item.Note != null)
                        {
                            worksheet.Cells[row, 29].Value = item.Note;
                            worksheet.Cells[row, 29].Style.WrapText = true;
                        }
                        worksheet.Cells[row, 30].Value = item.IfEmployed;
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

        private void EditStudent_Click(object sender, RoutedEventArgs e)
        {
            Students current = (sender as Button)?.DataContext as Students;
            var wnd = new EditStudentWindow(current);
            wnd.ShowDialog();
            Update();
            UpdateMyGroup();
        }

        private void Sort_Click(object sender, RoutedEventArgs e)
        {
            Update();
        }

        private void Clear()
        {
            CbSortDir.SelectedIndex = 0;
            CbSortField.SelectedIndex = 0;

            CbFilterGender.SelectedIndex = 0;
            CbFilterAge.SelectedIndex = 0;
            CbFilterGroup.SelectedIndex = 0;
            CbFilterCourse.SelectedIndex = 0;
            CbFilterSpeciality.SelectedIndex = 0;
            CbFilterTown.SelectedIndex = 0;
            CbFilterEmployment.SelectedIndex = 0;
            CbFilterFormOfStudy.SelectedIndex = 0;
            CbFilterTypeOfFinancing.SelectedIndex = 0;
            CbFilterIsOrphan.SelectedIndex = 0;
            CbFilterIsInvalid.SelectedIndex = 0;

            tbSearch.Clear();
        }

        private void Clear_Click(object sender, RoutedEventArgs e)
        {
            Clear();
            Update();
        }
        #endregion

        // My Groups
        #region
        private void OnSelectionChangedMyGroup(object sender, SelectionChangedEventArgs e)
        {
            var selectedItems = dgMyStudents.SelectedItems.OfType<int>().ToList();
            lbSelected1.Content = $"Выбрано: {selectedItems.Count()}";
            int s = 0;
            foreach (var r in dgMyStudents.SelectedItems)
            {
                s++;
            }
            lbSelected1.Content = $"Выбрано: {s}";
        }

        private void TbSearchMyGroup_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                _listSearch = _listMy.Where(Item => Item.Surname.ToLower().Contains(tbSearch1.Text.ToLower()) || Item.Name.ToLower().Contains(tbSearch1.Text.ToLower())
                    || Item.Patronymic.ToLower().Contains(tbSearch1.Text.ToLower()) || Item.HasNote.ToLower().Contains(tbSearch.Text.ToLower())).ToList();
                dgMyStudents.ItemsSource = _listSearch;

                lbCount1.Content = $"Количество записей: {_listSearch.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void SortMyGroup_Click(object sender, RoutedEventArgs e)
        {
            UpdateMyGroup();
        }

        private void ClearMyGroup()
        {
            CbSortDir1.SelectedIndex = 0;
            CbSortField1.SelectedIndex = 0;

            CbFilterAge1.SelectedIndex = 0;
            CbFilterTown1.SelectedIndex = 0;
            CbFilterEmployment1.SelectedIndex = 0;
            CbFilterTypeOfFinancing1.SelectedIndex = 0;

            tbSearch1.Clear();
        }

        private void ClearMyGroup_Click(object sender, RoutedEventArgs e)
        {
            ClearMyGroup();
            Update();
        }
        #endregion

        // Graduates
        #region
        private void OnSelectionChangedGraduates(object sender, SelectionChangedEventArgs e)
        {
            var selectedItems = dgGraduates.SelectedItems.OfType<int>().ToList();
            lbSelected2.Content = $"Выбрано: {selectedItems.Count()}";
            int s = 0;
            foreach (var r in dgGraduates.SelectedItems)
            {
                s++;
            }
            lbSelected2.Content = $"Выбрано: {s}";
        }

        private void TbSearchGraduates_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                _listSearch = _listGraduates.Where(Item => Item.Surname.ToLower().Contains(tbSearch2.Text.ToLower()) || Item.Name.ToLower().Contains(tbSearch2.Text.ToLower())
                    || Item.Patronymic.ToLower().Contains(tbSearch2.Text.ToLower()) || Item.HasNote.ToLower().Contains(tbSearch.Text.ToLower())).ToList();
                dgGraduates.ItemsSource = _listSearch;

                lbCount2.Content = $"Количество записей: {_listSearch.Count}";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }
        }

        private void ClearGraduates()
        {
            CbSortDir2.SelectedIndex = 0;
            CbSortField2.SelectedIndex = 0;

            CbFilterGender2.SelectedIndex = 0;
            CbFilterGroup2.SelectedIndex = 0;
            CbFilterSpeciality2.SelectedIndex = 0;
            CbFilterTown2.SelectedIndex = 0;
            CbFilterEmployment2.SelectedIndex = 0;
            CbFilterFormOfStudy2.SelectedIndex = 0;
            CbFilterTypeOfFinancing2.SelectedIndex = 0;
            CbFilterIsOrphan2.SelectedIndex = 0;
            CbFilterIsInvalid2.SelectedIndex = 0;

            tbSearch2.Clear();
        }

        private void ClearGraduates_Click(object sender, RoutedEventArgs e)
        {
            ClearGraduates();
            UpdateGraduates();
        }

        private void SortGraduates_Click(object sender, RoutedEventArgs e)
        {
            UpdateGraduates();
        }

        private void SaveGraduates_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog
            {
                FileName = "Document",
                DefaultExt = ".xlsx",
                Filter = "Excel Worksheets (.xlsx)|*.xlsx"
            };

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
                    worksheet.Cells[1, 6].AutoFitColumns();
                    worksheet.Cells[1, 7].Value = "Телефон";
                    worksheet.Cells[1, 8].Value = "Email";
                    worksheet.Cells[1, 9].Value = "Гражданство";
                    worksheet.Cells[1, 10].Value = "ДУЛ";
                    worksheet.Cells[1, 11].Value = "Серия паспорта";
                    worksheet.Cells[1, 12].Value = "Номер паспорта";
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
                    worksheet.Cells[1, 24].Value = "Вид финансирования";
                    worksheet.Cells[1, 24].AutoFitColumns();
                    worksheet.Cells[1, 25].Value = "Специальность";
                    worksheet.Cells[1, 25].AutoFitColumns();
                    worksheet.Cells[1, 26].Value = "Группа";
                    worksheet.Cells[1, 27].Value = "Классный руководитель";
                    worksheet.Cells[1, 27].AutoFitColumns();
                    worksheet.Cells[1, 28].Value = "Статус";
                    worksheet.Cells[1, 29].Value = "Примечания";
                    worksheet.Cells[1, 29].AutoFitColumns();
                    worksheet.Cells[1, 30].Value = "Трудоустройство";
                    worksheet.Cells[1, 30].AutoFitColumns();

                    int row = 2;
                    foreach (var item in _listGraduates)
                    {
                        worksheet.Cells[row, 1].Value = item.Surname;
                        worksheet.Cells[row, 1].AutoFitColumns();
                        worksheet.Cells[row, 2].Value = item.Name;
                        worksheet.Cells[row, 2].AutoFitColumns();
                        worksheet.Cells[row, 3].Value = item.Patronymic;
                        worksheet.Cells[row, 3].AutoFitColumns();
                        worksheet.Cells[row, 4].Value = item.Gender;
                        worksheet.Cells[row, 5].Value = item.DateOfBirth;
                        worksheet.Cells[row, 5].Style.Numberformat.Format = "dd.MM.yyyy";
                        worksheet.Cells[row, 6].Value = item.Age;
                        worksheet.Cells[row, 7].Value = item.Phone;
                        worksheet.Cells[row, 7].AutoFitColumns();
                        worksheet.Cells[row, 8].Value = item.Email;
                        worksheet.Cells[row, 8].AutoFitColumns();
                        worksheet.Cells[row, 9].Value = item.Nationality;
                        worksheet.Cells[row, 9].AutoFitColumns();
                        worksheet.Cells[row, 10].Value = item.IdentityDocument;
                        worksheet.Cells[row, 10].AutoFitColumns();
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
                        worksheet.Cells[row, 23].Value = item.FormOfStudy;
                        worksheet.Cells[row, 24].Value = item.TypeOfFinancing;
                        worksheet.Cells[row, 25].Value = item.Specialities.SpecialityName;
                        worksheet.Cells[row, 25].Style.WrapText = true;
                        if (item.GroupID == null)
                        {
                            worksheet.Cells[row, 26].Value = "На распределении";
                        }
                        else
                        {
                            worksheet.Cells[row, 26].Value = item.Groups.GroupNumber;
                            worksheet.Cells[row, 27].Value = item.Groups.Users.Surname + " " + item.Groups.Users.Name + " " + item.Groups.Users.Patronymic;
                        }
                        worksheet.Cells[row, 28].Value = item.Statuses.Status;
                        worksheet.Cells[row, 28].AutoFitColumns();
                        if (item.Note != null)
                        {
                            worksheet.Cells[row, 29].Value = item.Note;
                            worksheet.Cells[row, 29].Style.WrapText = true;
                        }
                        worksheet.Cells[row, 30].Value = item.IfEmployed;
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
        #endregion

        // Меню Аккаунт
        #region
        private void Profile_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new ProfileWindow(_currentUser);
            wnd.ShowDialog();
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new MainWindow();
            wnd.Show();
            Hide();
        }
        #endregion

        // Меню Списки
        #region
        private void Teachers_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new TeachersWindow(_currentUser);
            wnd.ShowDialog();
        }

        private void Groups_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new GroupsWindow();
            wnd.ShowDialog();
            UpLoad();
        }

        private void Specialities_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new SpecialitiesWindow();
            wnd.ShowDialog();
            UpLoad();
        }

        private void EmploymentHistory_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new EmploymentHistoryWindow(_currentUser);
            wnd.ShowDialog();
            Update();
        }
        #endregion

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
