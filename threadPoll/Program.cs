using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace threadPoll
{
    class Program
    {
        static void Main(string[] args)
        {

            //StreamWriter sw = new StreamWriter(@"D:\text.txt");
            //for (int i = 0; i <= 10; i++)
            //{
            //    result = i * i;
            //    sw.WriteLine(result);
            //}
            //sw.Close();

            //ThreadStart writeSecond = new ThreadStart(SecondThread);
            //Thread thread = new Thread(writeSecond);
            //thread.Start();


            //while (true)
            //{
            //    Console.WriteLine("Выполняется поток 1");
            //    Thread.Sleep(200);
            //}
            Worker1 wr1 = new Worker1();
            Worker2 wr2 = new Worker2();



            var key = ' ';
            while (key != 'q')
            {
                Console.WriteLine("Menu:");
                Console.WriteLine("1 - Run theads");
                Console.WriteLine("2 - Check thread status");                
                Console.WriteLine("3 - Pause/Start theads");
                key = Console.ReadKey().KeyChar;
                switch(key)
                {
                    case '1':
                       // wr1.ThreadCounter();
                        wr2.ThreadCounter();
                        break;
                    case '2':
                       // wr1.CheckStatus();
                        wr2.CheckStatus();
                        break;
                    case '3':
                        //wr1.PauseOrStartThread();
                        wr2.PauseOrStartThread();
                        break;
                    case 'q': break;
                    default: Console.WriteLine("Invalid input"); break;
                }
            }
        }

        //public static void SecondThread()
        //{
        //    while (true)
        //    {
        //        Console.WriteLine(new string(' ', 25) + "Выполняется поток 2");
        //        Thread.Sleep(300);
        //    }
        //}

        //public static void ThirdThread()
        //{
        //    while (true)
        //    {
        //        Console.WriteLine(new string(' ', 50) + "Выполняется поток 3");
        //        Thread.Sleep(500);
        //    }
        //}
    }
}
