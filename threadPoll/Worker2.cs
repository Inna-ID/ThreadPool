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
        private List<int> arrayOfPrimes = new List<int>();

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

        public void WriteToWord()
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
            //var arrayOfPrimes = new List<int>();
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
            _thread = new Thread(WriteToWord);
            // IF dont work
            if (_thread.ThreadState != ThreadState.WaitSleepJoin && _thread.ThreadState != ThreadState.Suspended && _thread.ThreadState != ThreadState.Running)
            {
                _thread.Name = threadName ?? "Worker2";
                _thread.Start();
            }
            else
            {
                Console.WriteLine($"Thread {_thread.Name} already running...");
            }
            
            Console.WriteLine($"Thread {_thread.Name} is running...");
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
                Console.WriteLine($"The thread {_thread.Name} suspended");
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