using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleFullTextSearch
{
    class FullTextSearch
    {
        static void Main(string[] args)
        {
           
            Search();
        }

        public static void Search()
        {
            string output = null;

            var CityName = "Blue Lake";

            var CityNameloweCase = CityName.ToLower();

            foreach (var part in CityNameloweCase.Split(' '))
            {
                if (part.Length > 3)
                {

                    for (int i = 3; i <= part.Length; i++)
                    {
                        output += " " + part.Substring(0, i);
                    }
                }
                else
                {
                    output += " " + part;
                }
            }

            //foreach (string word in words)
            //{
            //    if (word.Contains(" "))
            //    {
            //        for (int i = 3; i <= word.Length; i++)
            //        {
            //            output.Append(word.Substring(0, i));
            //            output.Append(' ');
            //        }
            //    }
            //    else
            //        output.Append(word + ' ');
            //}
            Console.WriteLine(output);
            Console.ReadKey();
            
        }
    }
}
