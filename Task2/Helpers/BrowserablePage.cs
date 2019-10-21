using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task2.Helpers
{
    public abstract class BroswerablePage
    {
        public virtual string URL { get { return string.Empty; } }
    }
}
