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
                var range = document.Range();
                document.SaveAs2(ref path);
                int min = 1;
                int max = 20;
                var arrayOfSimpleNum = new List<int>();
                int j;

                for (int i = min; i <= max; i++)
                {
                    for (j = 2; j <= i; j++)
                    {                        
                        if (i % j == 0) break;
                    }
                    if (j == i)
                    {
                        arrayOfSimpleNum.Add(i);
                    }
                }
                foreach (int item in arrayOfSimpleNum)
                {
                    Thread.Sleep(500);
                    
                    Console.WriteLine($"{new string(' ', 70)}{item}");
                }
                Console.WriteLine($"The thread {_thread.Name} is completed.");
                range.Text = string.Join(", ", arrayOfSimpleNum);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            { 
                document.Save();
                document.Close();
                CloseWord();
            }
        }

        public void CloseWord()
        {
            app.Quit();
            IsApplicaionClosed = true;
        }

        public void ThreadCounter(string threadName = null)
        {
            ThreadPool.QueueUserWorkItem(arg => WriteToWord());
            ThreadStart threadStart = new ThreadStart(WriteToWord);
            _thread = new Thread(threadStart);
            _thread.Name = threadName ?? "Worker2";

            if (_thread.ThreadState == ThreadState.WaitSleepJoin || _thread.ThreadState == ThreadState.Suspended)
            {
                return;
            }

            _thread.Start();
            Console.WriteLine($"The thread {_thread.Name} is running...");
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