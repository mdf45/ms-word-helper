using System;
using System.IO;

namespace MyTest
{
    class Program
    {
        static readonly Models.WordHelper _wordHelper;

        static Program()
        {
            _wordHelper = new Models.WordHelper();
        }

        static void Main(string[] args)
        {
            string file = string.Empty;

            while (!File.Exists(file) || !file.Contains(".doc"))
            {
                if (file != string.Empty)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Файла не существует!");
                    Console.ResetColor();
                }
                Console.Write("Введите путь до файла: ");
                file = Console.ReadLine();
            }

            _wordHelper.Logic(file);

            Console.ReadKey();
        }
    }
}
