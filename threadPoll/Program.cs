using System;
using ThreadPoolLibrary;

namespace threadPoll
{
    class Program
    {
        static void Main(string[] args)
        {
            ThreadsPool tp = new ThreadsPool();            
            WinAPI.СhechSystemTime(); //апи функция получение даты системы
            var key = ' ';
            while (key != 'q')
            {
                Console.WriteLine("\nMenu:");
                Console.WriteLine("1 - Run theads pool");
                Console.WriteLine("2 - Check info about threads");
                Console.WriteLine("q - To exit");
                key = Console.ReadKey().KeyChar;
                switch(key)
                {
                    case '1':
                        tp.AddToPool();
                        break;
                    case '2':
                        tp.InfoAboutThreadsInPool();
                        break;
                    case 'q': break;
                    default: Console.WriteLine("Invalid input"); break;
                }
            }
        }
    }
}
