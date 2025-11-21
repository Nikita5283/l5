using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace l5
{
    internal class Inspector
    {
        public static double CheckDouble(string message)
        {
            double x;

            while (true)
            {
                Console.Write(message);
                string inp = Console.ReadLine().Replace('.', ',');

                if (double.TryParse(inp, out x))
                {
                    return x;
                }

                Console.WriteLine("Неверный тип данных! Должен быть double");
            }
        }

        public static int CheckInt32(string message)
        {
            int x;

            while (true)
            {
                Console.Write(message);
                string inp = Console.ReadLine().Replace('.', ',');

                if (int.TryParse(inp, out x))
                {
                    return x;
                }

                Console.WriteLine("Неверный тип данных! Должен быть int");
            }
        }
    }
}
