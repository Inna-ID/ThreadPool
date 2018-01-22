using System;
using System.Text;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace threadPoll
{
    class ThreadsPool
    {
        private Thread _thread;
        int workerThreads;
        int portThreads;
        WorkWithThreads wwt = new WorkWithThreads();

        public void AddToPool()
        {
            ThreadPool.SetMinThreads(10, 0); //задаем мин. кол-во потоков в пуле
            ThreadPool.SetMaxThreads(10, 0); //задаем мач. кол-во потоков в пуле
            Console.WriteLine($"{ new string(' ', 40)} Work from thread pool...");
            //добавляем задачи в пул потоков
            ThreadPool.QueueUserWorkItem(new WaitCallback(FindPrimeNumbers));
            ThreadPool.QueueUserWorkItem(new WaitCallback(ToSecondDegree));
            ThreadPool.QueueUserWorkItem(new WaitCallback(ToThirdDegree));
            Task.Run(() => ToThirdDegree(null));

            CheckThreads();
        }

        public void CheckThreads()
        {
            ThreadPool.GetAvailableThreads(out workerThreads, out portThreads);
            int availableThreads = workerThreads;
            if (availableThreads == 10)
            {
                Console.WriteLine("tasks ends!");
            }
        }

        public void InfoAboutThreadsInPool()
        {
            //получаем максимального кол-ва потоков
            ThreadPool.GetMaxThreads(out workerThreads, out portThreads);
            int maxThreads = workerThreads;
            Console.WriteLine($"Max threads in the thread pool: {workerThreads}");
            //получаем доступное кол-ва потоков
            ThreadPool.GetAvailableThreads(out workerThreads, out portThreads);
            int availableThreads = workerThreads;
            Console.WriteLine($"Available worker threads: {workerThreads}");
            //пишем в консоль сколько потоков из пула используется
            Console.WriteLine($"Threads count used in the pool: {maxThreads - availableThreads}");

        }

        public void FindPrimeNumbers(Object stateInfo)
        {
            WorkWithWord www = new WorkWithWord(@"D:\Doc5.docx");
            List<int> arrayOfPrimes = new List<int>();
            int min = 1;
            int max = 100;
            int j;

            for (int i = min; i <= max; i++)
            {
                for (j = 2; j <= i; j++)
                {
                    if (i % j == 0) break;
                }
                if (j == i)
                {
                    arrayOfPrimes.Add(i);
                }
            }

            foreach (int item in arrayOfPrimes)
            {
                Thread.Sleep(500);
                Console.WriteLine($"{new string(' ', 70)}{item}");
            }
            Console.WriteLine(Thread.CurrentThread.IsThreadPoolThread ? "yes" : "not ");//!!!!!!!!!!!
            // пишем список протых чисел в ворд, разделяя элементы ,
            www.WriteToWord($"Prime numbers\n{string.Join(", ", arrayOfPrimes)}");
        }


        private void ToThirdDegree(Object stateInfo)
        {
            WorkWithWord www = new WorkWithWord(@"D:\Doc4.docx");
            string strArray = "";
            string str = "";             
            for (int i = 0; i < 20; i++)
            {
                Thread.Sleep(500);
                str = $"i^3 = {i *  i *i}\n";
                strArray += str;
                Console.WriteLine($"{new string(' ', 30)} i^3 = {i * i * i}");
            }
            //записываем результат в ворд
            www.WriteToWord($"Table of third degree\n{strArray}");
        }
        private void ToSecondDegree(Object stateInfo)
        {
            WorkWithWord www = new WorkWithWord(@"D:\Doc3.docx");
            string strArray = "";
            string str = "";
            for (int i = 0; i < 20; i++)
            {
                Thread.Sleep(500);
                str = $"i^2 = {i * i}\n";
                strArray += str;
                Console.WriteLine($"{new string(' ', 50)} {str}");
            }
            //записываем результат в ворд
            www.WriteToWord($"Table of second degree\n{strArray}");
        }
        

    }
}