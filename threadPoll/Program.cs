using System;
namespace threadPoll
{
    class Program
    {
        static void Main(string[] args)
        {

            ThreadsPool tp = new ThreadsPool();
            WorkWithThreads wwt = new WorkWithThreads();
            Test test = new Test();

            var key = ' ';
            while (key != 'q')
            {
                Console.WriteLine("Menu:");
                Console.WriteLine("1 - Run theads");
                Console.WriteLine("2 - Check thread status");                
                Console.WriteLine("3 - Pause/Start theads\n");
                key = Console.ReadKey().KeyChar;
                switch(key)
                {
                    case '1':
                        wwt.CreateThread();
                        tp.CreateThread();
                        test.CreateThread();
                        break;
                    case '2':
                        wwt.CheckStatus();
                        tp.CheckStatus();
                        break;
                    case '3':
                        //wr1.PauseOrStartThread();
                        tp.PauseOrStartThread();
                        break;
                    case 'q': break;
                    default: Console.WriteLine("Invalid input"); break;
                }
            }
        }
    }
}
