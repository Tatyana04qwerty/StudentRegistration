using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Diplom.Models;

namespace Diplom
{
    internal class AppData
    {
        private static ContingentEntities2 _context;

        public static ContingentEntities2 GetContext()
        {
            if (_context == null)
                _context = new ContingentEntities2();
            return _context;
        }
    }
}
