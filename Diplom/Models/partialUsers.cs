using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Diplom.Models
{
    partial class Users
    {
        public int GroupsAmount
        {
            get
            {
                int groupAmount = Groups.Where(x => x.FormMaster == ID).Count();
                return groupAmount;
            }
            set { }
        }
    }
}
