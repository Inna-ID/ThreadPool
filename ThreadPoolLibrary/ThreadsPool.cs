using System;
using System.Threading;

namespace ThreadPoolLibrary
{
    public class ThreadsPool
    {
        private Thread _thread;
        int workerThreads;
        int portThreads;
        TasksForThreads task = new TasksForThreads();
        
        public void AddToPool()
        {
            //задаем минимальное количество потоков, которые должны запускаться сразу после создания пула
            ThreadPool.SetMinThreads(7, 0);
            //задаем максимальное количество потоков, доступных в пуле
            ThreadPool.SetMaxThreads(10, 0);

            Console.WriteLine($"{ new string(' ', 40)} Work from thread pool...");

            //добавляем задачи в пул потоков
            ThreadPool.QueueUserWorkItem(new WaitCallback(task.FindPrimeNumbers));
            ThreadPool.QueueUserWorkItem(new WaitCallback(task.ToSecondDegree));
            ThreadPool.QueueUserWorkItem(new WaitCallback(task.Сalculation));
            ThreadPool.QueueUserWorkItem(new WaitCallback(task.ReadFromFile));
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
    }
}