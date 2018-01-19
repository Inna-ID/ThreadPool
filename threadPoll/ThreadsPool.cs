using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Office.Interop.Word;

namespace threadPoll
{
    class ThreadsPool
    {
        private Thread _thread;
        private Application app;
        private Document document;
        private bool IsApplicaionClosed = false;
        private object path;
        SystemTime st = new SystemTime();


        public ThreadsPool(string path = null)
        {
            this.path = path ?? @"D:\Doc2.docx";
        }

        public Document OpenNewDoc()
        {
            if (app == null || IsApplicaionClosed)
                app = new Application();
            return app.Documents.Add();
        }

        public void WriteToWord(Object stateInfo)
        {
            try
            {
                document = OpenNewDoc();
                Range range = document.Range(); //object of range of text in the word doc
                var str = string.Join(", ", FindPrimeNumbers());
                range.Text = $"Prime numbers: {str} \n {chechSystemTime()}";
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
            List<int> arrayOfPrimes = new List<int>();
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
            
            Console.WriteLine($"{ new string(' ', 70)} Work from thread pool...");
            foreach (int item in arrayOfPrimes)
            {
                Thread.Sleep(500);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"{new string(' ', 70)}{item}");
            }
            return arrayOfPrimes;
        }

        public void CreateThread(string threadName = null)
        {
            ThreadPool.QueueUserWorkItem(new WaitCallback(WriteToWord));            
        }

        public string chechSystemTime()
        {
            LibWrap.GetSystemTime(st);            
            var date = $"{st.day}.{st.month}.{st.year} | {st.hour} : {st.minute}";
            return date;
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