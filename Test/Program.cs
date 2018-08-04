using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            string a = "梁";
            var b = Validation.IsChinese(a);
            Console.WriteLine(b);
            Console.ReadKey();
        }
    }
}
