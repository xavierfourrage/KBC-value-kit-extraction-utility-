using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ValueKit_ExtractionUI
{
    class Utilities
    {
        public bool Confirm(string title)
        {
            ConsoleKey response;
            do
            {
                WriteInYellow($"{title} [y/n] ");
                Console.ForegroundColor = ConsoleColor.White;
                response = Console.ReadKey(false).Key;
                if (response != ConsoleKey.Enter)
                {
                    Console.WriteLine();
                }
            } while (response != ConsoleKey.Y && response != ConsoleKey.N);

            return (response == ConsoleKey.Y);
        }

        public void PressEnterToExit()
        {
            WriteInGreen("Press Enter to exit");
            ConsoleKeyInfo keyInfo = Console.ReadKey();
            while (keyInfo.Key != ConsoleKey.Enter)
                keyInfo = Console.ReadKey();
        }
        public void WriteInGreen(string text)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(text);
        }

        public void WriteInYellow(string text)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(text);
        }

        public void WriteInBlue(string text)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(text);
        }

        public void WriteInWrite(string text)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(text);
        }

        public void WriteInRed(string text)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(text);
        }


    }
}
