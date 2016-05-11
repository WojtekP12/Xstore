using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FromExcelToTXT
{
    enum Brand
    {
        gucci = 2, slp, tm, amq, bal, bv, br, sr
    }

    public static class BrandParser
    {
        public static T ParseEnum<T>(string value)
        {
            return (T)Enum.Parse(typeof(T), value, true);
        }
    }
}
