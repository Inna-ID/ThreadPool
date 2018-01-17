using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Office.Interop.Word;

namespace threadPoll
{
    class Worker1
    {
        private Thread _thread;
        private Application app;
        private Document document;
        private bool IsApplicaionClosed = false;
        private object path;

        public Worker1(string path = null)
        {
            this.path = path ?? @"D:\Doc1.docx";
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
                int result;
                int[] array = new int[10];

                for (int i = 1; i < 10; i++)
                {
                    Thread.Sleep(1000);
                    result = i * i;
                    array[i] = result;
                    
                    Console.WriteLine($"{new string(' ', 40)}{i} в квадрате = {result}");

                }
                Console.WriteLine($"Поток {_thread.Name} завершен.");
                range.Text = "Таблица квадратов";
                range.Text = array.ToString();
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
            ThreadStart threadStart = new ThreadStart(WriteToWord);
            _thread = new Thread(threadStart);
            _thread.Name = threadName ?? "Worker1";

            if (_thread.ThreadState == ThreadState.WaitSleepJoin || _thread.ThreadState == ThreadState.Suspended)
            {
                return;
            }

            _thread.Start();
            Console.WriteLine($"Поток {_thread.Name} выполняется...");
        }

        public void PauseOrStartThread()
        {
            if (_thread == null)
            {
                Console.WriteLine("Поток невозможно приостановить, т.к он не создан");
                return;
            }
            if (_thread.ThreadState == ThreadState.WaitSleepJoin)
            {
                _thread.Suspend();
                Console.WriteLine($"Поток {_thread.Name} приоставновлен");
            }
            else if (_thread.IsAlive && _thread.ThreadState == ThreadState.Suspended)
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
                Console.WriteLine($"Thread 1 is not created");
            }
        }

    }
}