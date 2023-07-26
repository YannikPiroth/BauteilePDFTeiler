using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BauteilePDFTeiler
{
    internal class Bauteil
    {
        private string _name;
        private int _seite;

        public string Name
        {
            get { return _name; }
            set { _name = value; }
        }
        public int Seite
        {
            get { return _seite; }
            set { _seite = value; }
        }
    }
}
