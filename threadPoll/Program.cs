using System;

namespace threadPoll
{
    class Program
    {
        static void Main(string[] args)
        {
            ThreadsPool tp = new ThreadsPool();
            WorkWithThreads wwt = new WorkWithThreads();
            WinAPI.СhechSystemTime(); //апи функция получение даты системы
            var key = ' ';
            while (key != 'q')
            {
                Console.WriteLine("\nMenu:");
                Console.WriteLine("1 - Run theads");
                Console.WriteLine("2 - Check info about thread");
                Console.WriteLine("q - To exit");
                //Console.WriteLine("3 - Pause/Start theads pool");
                //Console.WriteLine("4 - Own pool\n");
                key = Console.ReadKey().KeyChar;
                switch(key)
                {
                    case '1':
                        //wwt.CreateThread(null);
                        tp.AddToPool();
                        //test.CreateThread();
                        break;
                    case '2':
                        wwt.CheckStatus();
                        tp.InfoAboutThreadsInPool();
                        break;
                    case '3':
                        wwt.PauseOrStartThread();
                        break;
                    case 'q': break;
                    default: Console.WriteLine("Invalid input"); break;
                }
            }
        }
    }
}
