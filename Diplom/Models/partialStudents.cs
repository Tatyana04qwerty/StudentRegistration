using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Diplom.AppData;

namespace Diplom.Models
{
    partial class Students
    {
        public string IfOrphan
        {
            get
            {
                string b;
                if (IsOrphan == true)
                {
                    b = "Да";
                }
                else
                {
                    b = "Нет";
                }
                return b;
            }
            set { }
        }
        public string IfInvalid
        {
            get
            {
                string b;
                if (IsInvalid == true)
                {
                    b = "Да";
                }
                else
                {
                    b = "Нет";
                }
                return b;
            }
            set { }
        }
        public string IfEmployed
        {
            get
            {
                string b;
                if (IsEmployed == true)
                {
                    b = "Трудоустроен";
                }
                else
                {
                    b = "Не трудоустроен";
                }
                return b;
            }
            set { }
        }
        public int Age
        {
            get
            {
                DateTime today = DateTime.Today;
                int age = today.Year - DateOfBirth.Year;
                if (DateOfBirth > today.AddYears(-age)) age--;
                return age;
            }
            set { }
        }
        public int Course
        {
            get
            {
                StringBuilder group = new StringBuilder();
                DateTime today = DateTime.Today;
                int month = today.Month;
                int year = today.Year;
                int course;
                Groups gr = GetContext().Groups.Where(x => x.ID == GroupID).FirstOrDefault();
                if (gr != null)
                {
                    for (int i = 1; i < 3; i++)
                    {
                        group.Append(Groups.GroupNumber[i]);
                    }
                    int groupNumber = Convert.ToInt32(group.ToString());
                    if (month > 8) course = year - groupNumber + 1999;
                    else course = year - groupNumber - 2000;
                    return course;
                }
                else
                    return 0;
            }
            set { }
        }
        public string HasNote
        {
            get
            {
                if (Note != null) return Note;
                else return "";
            }
            set { }
        }
    }
}
