using System;
using System.IO;

namespace Reader
{
    class Program
    {
        static void Main(string[] args)
        {

            // Document.ExcelReader();
            read();
            //Console.ReadKey();
        }

        static void read()
        {
            string[] updated;
            string[] text = System.IO.File.ReadAllLines(@"C:\Users\slip4\Desktop\unsorted.txt");
            var sorted = Array.Sort(text);
            File.WriteAllLines(@"C:\Users\slip4\Desktop\sorted.txt", sorted);
        }

       
    }
}
