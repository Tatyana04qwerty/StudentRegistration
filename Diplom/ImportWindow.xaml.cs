using ExcelDataReader;
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
using System.Data;
using System.IO;
using Microsoft.Win32;
using static Diplom.AppData;
using System.Globalization;
using Diplom.Models;

namespace Diplom
{
    /// <summary>
    /// Логика взаимодействия для ImportWindow.xaml
    /// </summary>
    public partial class ImportWindow : Window
    {
        IExcelDataReader edr;
        private static Specialities _currentSpec = new Specialities();
        private static Students _currentStudents = new Students();
        private static Students _studentToCheck = new Students();

        public ImportWindow()
        {
            InitializeComponent();
        }

        private void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "EXCEL Files (*.xlsx)|*.xlsx|EXCEL Files 2003 (*.xls)|*.xls|All files (*.*)|*.*"
            };
            if (openFileDialog.ShowDialog() != true)
                return;

            DbGrig.ItemsSource = ReadFile(openFileDialog.FileName);
        }

        private string GetDataGridCellInfo(int i, int j)
        {
            var ci = new DataGridCellInfo(DbGrig.Items[i], DbGrig.Columns[j]);
            var content = ci.Column.GetCellContent(ci.Item) as TextBlock;
            return content.Text;
        }

        private void ImportData_Click(object sender, RoutedEventArgs e)
        {
            if (DbGrig.ItemsSource == null)
            {
                MessageBox.Show("Нет данных для импорта");
            }
            else
            {
                try
                {
                    for (int i = 0; i < DbGrig.Items.Count - 1; i++)
                    {
                        int j = 0;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            string[] fio = GetDataGridCellInfo(i, j).Split(' ');
                            if (fio[0].Length > 50)
                            {
                                MessageBox.Show($"Фамилия превышает допустимое количество символов: 50\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            if (fio[0].Length > 50)
                            {
                                MessageBox.Show($"Имя превышает допустимое количество символов: 50\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            if (fio[0].Length > 50)
                            {
                                MessageBox.Show($"Отчество превышает допустимое количество символов: 50\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            _currentStudents.Surname = fio[0];
                            _currentStudents.Name = fio[1];
                            _currentStudents.Patronymic = fio[2];
                        }
                        else
                        {
                            MessageBox.Show($"ФИО не указано (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 10)
                            {
                                MessageBox.Show($"Пол превышает допустимое количество символов: 10\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.Gender = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"Пол не указан (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            _currentStudents.DateOfBirth = Convert.ToDateTime(GetDataGridCellInfo(i, j));
                        }
                        else
                        {
                            MessageBox.Show($"Дата рождения не указана (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 90)
                            {
                                MessageBox.Show($"Гражданство превышает допустимое количество символов: 90\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.Nationality = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"Гражданство не указано (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            _currentStudents.Phone = GetDataGridCellInfo(i, j);
                            _studentToCheck = GetContext().Students.Where(x => x.Surname == _currentStudents.Surname && x.Name == _currentStudents.Name
                                && x.Patronymic == _currentStudents.Patronymic && x.Phone == _currentStudents.Phone).FirstOrDefault();
                            if (_studentToCheck != null)
                            {
                                if (MessageBox.Show($"Такой студент уже числится в базе (Строка {i + 1}). Прервать операцию добавления?", "Внимание",
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                                    break;
                                else
                                    continue;

                            }
                            else if (GetDataGridCellInfo(i, j).Length > 16)
                            {
                                MessageBox.Show($"Телефон превышает допустимое количество символов: 16\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                        }
                        else
                        {
                            MessageBox.Show($"Телефон не указан (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 80)
                            {
                                MessageBox.Show($"Почта превышает допустимое количество символов: 80\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.Email = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            _currentStudents.Email = null;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 25)
                            {
                                MessageBox.Show($"ДУЛ превышает допустимое количество символов: 25\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.IdentityDocument = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"ДУЛ не указан (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 4)
                            {
                                MessageBox.Show($"Серия паспорта превышает допустимое количество символов: 4\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.PassportSeries = GetDataGridCellInfo(i, j).Replace(" ", "");
                        }
                        else
                        {
                            MessageBox.Show($"Серия паспорта не указана (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 6)
                            {
                                MessageBox.Show($"Номер паспорта превышает допустимое количество символов: 6\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.PassportNumber = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"Номер паспорта не указан (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 250)
                            {
                                MessageBox.Show($"Отдел выдачи паспорта превышает допустимое количество символов: 250\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.IssuedBy = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"Кем выдан ДУЛ не указано (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            _currentStudents.DateOfIssue = Convert.ToDateTime(GetDataGridCellInfo(i, j));
                        }
                        else
                        {
                            MessageBox.Show($"Дата выдачи ДУЛ-а не указана (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 300)
                            {
                                MessageBox.Show($"Адрес превышает допустимое количество символов: 300\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.PermanentRegistrationAddress = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"Адрес постоянной регистрации не указан (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 14)
                            {
                                MessageBox.Show($"СНИЛС превышает допустимое количество символов: 14\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.IIAN = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"СНИЛС не указан (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Replace(" ", "").Length > 12)
                            {
                                MessageBox.Show($"ИНН превышает допустимое количество символов: 12\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.ITN = GetDataGridCellInfo(i, j).Replace(" ", "");
                        }
                        else
                        {
                            _currentStudents.ITN = null;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            string str = GetDataGridCellInfo(i, j).Replace(',', '.');
                            _currentStudents.GPA = float.Parse(str, new NumberFormatInfo());
                        }
                        else
                        {
                            MessageBox.Show($"Средний балл не указан (Столбец{j + 1} строка {i + 1})");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j).Replace(" ", "") != "")
                        {
                            if (GetDataGridCellInfo(i, j) == "Да")
                                _currentStudents.IsOrphan = true;
                            else _currentStudents.IsOrphan = false;
                        }
                        else
                        {
                            MessageBox.Show($"Нет данных о сиротстве студента (Столбец{j + 1} строка {i + 1}). Присвоено: Нет");
                            _currentStudents.IsOrphan = false;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j).Replace(" ", "") != "")
                        {
                            if(GetDataGridCellInfo(i, j) == "Да")
                                _currentStudents.IsInvalid = true;
                            else _currentStudents.IsInvalid = false;
                        }
                        else
                        {
                            MessageBox.Show($"Нет данных об инвалидности студента (Столбец{j+1} строка {i+1}). Присвоено: Нет");
                            _currentStudents.IsInvalid = false;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            _currentStudents.CauseOfDisability = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            if (GetDataGridCellInfo(i, j).Length > 150)
                            {
                                MessageBox.Show($"Нозологическая группа превышает допустимое количество символов: 150\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.CauseOfDisability = null;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            string specialtyName = GetDataGridCellInfo(i, j);
                            _currentSpec = GetContext().Specialities.Where(x => x.SpecialityName == specialtyName).FirstOrDefault();
                            if(_currentSpec.ID != 0)
                            {
                                _currentStudents.SpecialityID = _currentSpec.ID;
                            }
                            else
                            {
                                MessageBox.Show($"Специальность не найдена (Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Специальность не указана");
                            continue;
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 15)
                            {
                                MessageBox.Show($"Форма обучения превышает допустимое количество символов: 15\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.FormOfStudy = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"Нет данных о форме обучения (Столбец{j + 1} строка {i + 1}). Присвоено: Бюджет");
                            _currentStudents.FormOfStudy = "Очная";
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 40)
                            {
                                MessageBox.Show($"Базовое образование превышает допустимое количество символов: 40\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.BasicEducation = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"Нет данных о базовом образовании (Столбец{j + 1} строка {i + 1}). Присвоено: Основное общее (9 классов)");
                            _currentStudents.BasicEducation = "Основное общее (9 классов)";
                        }
                        j++;
                        if (GetDataGridCellInfo(i, j) != "")
                        {
                            if (GetDataGridCellInfo(i, j).Length > 15)
                            {
                                MessageBox.Show($"Вид финансирования превышает допустимое количество символов: 15\n(Столбец{j + 1} строка {i + 1})");
                                continue;
                            }
                            else
                                _currentStudents.TypeOfFinancing = GetDataGridCellInfo(i, j);
                        }
                        else
                        {
                            MessageBox.Show($"Нет данных о виде финансирования (Столбец{j + 1} строка {i + 1}). Присвоено: Бюджет");
                            _currentStudents.TypeOfFinancing = "Бюджет";
                        }
                        _currentStudents.StatusID = 1;
                        _currentStudents.IsEmployed = false;
                        GetContext().Students.Add(_currentStudents);
                        GetContext().SaveChanges();
                    }
                    
                    MessageBox.Show("Студенты добавлены");
                    Hide();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void AddStudent_Click(object sender, RoutedEventArgs e)
        {
            var wnd = new AddStudentWindow();
            wnd.ShowDialog();
            Hide();
        }

        private DataView ReadFile(string fileNames)
        {
            try
            {
                var extension = fileNames.Substring(fileNames.LastIndexOf('.'));
                FileStream stream = File.Open(fileNames, FileMode.Open, FileAccess.Read);
                if (extension == ".xlsx")
                    edr = ExcelReaderFactory.CreateOpenXmlReader(stream);
                else if (extension == ".xls")
                    edr = ExcelReaderFactory.CreateBinaryReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                DataSet dataSet = edr.AsDataSet(conf);
                DataView dtView = dataSet.Tables[0].AsDataView();

                edr.Close();
                return dtView;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return null;
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }
    }
}
