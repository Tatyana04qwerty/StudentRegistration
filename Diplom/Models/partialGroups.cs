using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Diplom.AppData;

namespace Diplom.Models
{
    partial class Groups
    {
        public int Size
        {
            get
            {
                int groupSize = GetContext().Students.Where(x => x.GroupID == ID).Count();
                return groupSize;
            }
            set { }
        }
        public string Status
        {
            get
            {
                Students student = GetContext().Students.Where(x => x.GroupID == ID).FirstOrDefault();
                if(student != null)
                {
                    if (student.StatusID == 2)
                        return "Выпущена";
                    else
                        return "Обучается";
                }
                else
                    return "Обучается";
            }
            set { }
        }
        public string Speciality
        {
            get
            {
                Students student = GetContext().Students.Where(x => x.GroupID == ID).FirstOrDefault();
                if (student != null)
                {
                    Specialities speciality = GetContext().Specialities.Where(x => x.ID == student.SpecialityID).First();
                    return speciality.SpecialityName;
                }
                else return null;
            }
            set { }
        }
        public int GroupCourse
        {
            get
            {
                StringBuilder group = new StringBuilder();
                DateTime today = DateTime.Today;
                int month = today.Month;
                int year = today.Year;
                int course;
                for (int i = 1; i < 3; i++)
                {
                    group.Append(GroupNumber[i]);
                }
                int groupNumber = Convert.ToInt32(group.ToString());
                if (month > 8) course = year - groupNumber + 1999;
                else course = year - groupNumber - 2000;
                return course;
            }
            set { }
        }
    }
}