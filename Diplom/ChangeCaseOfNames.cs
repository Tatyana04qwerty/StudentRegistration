using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Diplom
{
    public class ChangeCaseOfNames
    {
        public StringBuilder name;
        public StringBuilder surname;
        public StringBuilder patronymic;
        public string gender;
        public string consonants = "бвгджзйклмнпрстфцчшщ";

        public ChangeCaseOfNames() { }

        public ChangeCaseOfNames(string name, string surname, string patronymic, string gender)
        {
            this.name = new StringBuilder(name);
            this.surname = new StringBuilder(surname);
            this.patronymic = new StringBuilder(patronymic);
            this.gender = gender;
        }

        public void ChangeCase()
        {
            ChangeName();
            ChangeSurame();
            ChangePatronymic();
        }

        public void ChangeName()
        {
            if (gender == "Жен.")
            {
                if (name[name.Length - 1] == 'а')
                {
                    name[name.Length - 1] = 'у';
                }
                else if (name[name.Length - 1] == 'я')
                {
                    name[name.Length - 1] = 'ю';
                }
            }
            else if (gender == "Муж.")
            {
                if(name.Equals("Лев"))
                {
                    name.Replace("Лев", "Льва");
                }
                else if (name[name.Length - 1] == 'а')
                {
                    name[name.Length - 1] = 'у';
                }
                else if (name[name.Length - 1] == 'я')
                {
                    name[name.Length - 1] = 'ю';
                }
                else if (name[name.Length - 1] == 'й')
                {
                    name[name.Length - 1] = 'я';
                }
                else if (consonants.Contains(name[name.Length - 1]))
                {
                    name.Append('а');
                }
            }
        }

        public void ChangeSurame()
        {
            if (gender == "Жен.")
            {
                if (surname[surname.Length - 2] == 'а' && surname[surname.Length - 1] == 'я')
                {
                    surname[surname.Length - 2] = 'у';
                    surname[surname.Length - 1] = 'ю';
                }
                else if (surname[surname.Length - 1] == 'а')
                {
                    surname[surname.Length - 1] = 'у';
                }
            }
            else if (gender == "Муж.")
            {
                if ((surname[surname.Length - 2] == 'и' || surname[surname.Length - 2] == 'ы') && surname[surname.Length - 1] == 'й')
                {
                    surname[surname.Length - 2] = 'о';
                    surname[surname.Length - 1] = 'г';
                    surname.Append('о');
                }
                else if (surname[surname.Length - 1] == 'й' || surname[surname.Length - 1] == 'ь')
                {
                    surname[surname.Length - 1] = 'я';
                }
                else if (surname[surname.Length - 1] == 'я')
                {
                    surname[surname.Length - 1] = 'ю';
                }
                else if (consonants.Contains(surname[surname.Length - 1]))
                {
                    surname.Append('а');
                }
            }
        }

        public void ChangePatronymic()
        {
            if (gender == "Жен.")
            {
                patronymic[patronymic.Length - 1] = 'у';
            }
            else if (gender == "Муж.")
            {
                patronymic.Append('а');
            }
        }

        public string GetName
        {
            get
            {
                return name.ToString();
            }
        }

        public string GetSurame
        {
            get
            {
                return surname.ToString();
            }
        }

        public string GetPatronymic
        {
            get
            {
                return patronymic.ToString();
            }
        }
    }
}
