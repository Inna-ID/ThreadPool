using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Office.Interop.Word;

namespace threadPoll
{    
    class Worker2
    {
        private Thread _thread;
        private Application app;
        private Document document;
        private bool IsApplicaionClosed = false;
        private object path;
        List<int> arrayOfPrimes = new List<int>();



        public Worker2(string path = null)
        {
            this.path = path ?? @"D:\Doc2.docx";
        }

        public Document OpenNewDoc()
        {
            if (app == null || IsApplicaionClosed)
                app = new Application();
            return app.Documents.Add(Visible: true);
        }

        public void WriteToWord(Object stateInfo)
        {
            try
            {
                document = OpenNewDoc();
                Range range = document.Range();
                FindPrimeNumbers();

                var str = string.Join(", ", arrayOfPrimes);
                range.Text = $"Prime numbers: {str}";

                document.SaveAs(path);
                document.Close();
                CloseWord();                
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        public void CloseWord()
        {
            app.Quit();
            IsApplicaionClosed = true;
        }

        public List<int> FindPrimeNumbers()
        {
            int min = 1;
            int max = 30;
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
            Console.WriteLine($"Thread {_thread.Name} is completed.");

            return arrayOfPrimes;
        }

        public void CreateThread(string threadName = null)
        {

            if (_thread == null || _thread.ThreadState == ThreadState.Stopped)
            {
                _thread = new Thread(WriteToWord);
                _thread.Name = threadName ?? "Worker2";
                _thread.Start();
            }
            //ThreadPool.QueueUserWorkItem(new WaitCallback(WriteToWord));
            ThreadPool.QueueUserWorkItem(new WaitCallback(ThreadProc2));
            ThreadPool.QueueUserWorkItem(new WaitCallback(ThreadProc));
            ThreadPool.QueueUserWorkItem(new WaitCallback(ThreadProc3));

            //Console.WriteLine($"Thread {_thread.Name} is running...");
        }

         private void ThreadProc(Object stateInfo)
        {
            Console.WriteLine("Hello from the thread pool.");
            for (int i = 0; i < 5; i++)
            {
                Thread.Sleep(1000);
                Console.Write(".");
            }
        }
        private void ThreadProc2(Object stateInfo)
        {
            Console.WriteLine("Hello from the thread pool.");
            for (int i = 0; i < 10; i++)
            {
                Thread.Sleep(500);
                Console.WriteLine($"{new string(' ', 50)} {i*i}");
            }
        }

        private void ThreadProc3(Object stateInfo)
        {
            Console.WriteLine("Hello from the thread pool.");
            for (int i = 0; i < 30; i++)
            {
                Thread.Sleep(200);
                Console.WriteLine($"{new string(' ', 60)}hi");
            }
        }

        public void PauseOrStartThread()
        {
            if(_thread == null)
            {
                Console.WriteLine("The thread can not be paused because it is not created");
                return;
            }
            if (_thread.ThreadState == ThreadState.WaitSleepJoin)
            {
                _thread.Suspend();
                Console.WriteLine($"Thread {_thread.Name} suspended");
            }
            else if(_thread.IsAlive && _thread.ThreadState == ThreadState.Suspended)
            {
                _thread.Resume();
            }
        }

        public void CheckStatus()
        {
            if (_thread != null)
            {
                Console.WriteLine($"Thread {_thread.Name} status: {_thread.ThreadState.ToString()}");
            }
            else
            {
                Console.WriteLine($"Thread 2 is not created");
            }
        }

    }
}