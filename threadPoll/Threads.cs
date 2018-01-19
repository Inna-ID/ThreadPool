using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Office.Interop.Word;

namespace threadPoll
{
    class WorkWithThreads
    {
        private Thread _thread;
        private Application app;
        private Document document;
        private bool IsApplicaionClosed = false;
        private object path;
        SystemTime st = new SystemTime();
        
        public WorkWithThreads(string path = null)
        {
            this.path = path ?? @"D:\Doc1.docx";
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
                var str = string.Join(", ", ToSecondDegree());
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

        public List<int> ToSecondDegree()
        {            
            Console.WriteLine($"{new string(' ', 30)}Thread {_thread.Name} is running...");
            List<int> array = new List<int>();
            for (int i = 0; i <= 10 ; i++)
            {
                array.Add(i*i);
                Thread.Sleep(100);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{new string(' ', 30)}{i*i}");
            }                 
            return array;
        }

        public void CreateThread(string threadName = null)
        {
            //if thread not created or completed
            if (_thread == null || _thread.ThreadState == ThreadState.Stopped)
            {
                _thread = new Thread(WriteToWord);
                _thread.Name = threadName ?? "Worker1";
                _thread.Start();
            }            
        }

        public string chechSystemTime()
        {
            LibWrap.GetSystemTime(st);
            var date = $"{st.day}.{st.month}.{st.year} | {st.hour} : {st.minute}";
            return date;
        }
       
        public void PauseOrStartThread()
        {
            if (_thread == null)
            {
                Console.WriteLine("The thread can not be paused because it is not created");
                return;
            }
            if (_thread.ThreadState == ThreadState.WaitSleepJoin)
            {
                _thread.Suspend();
                Console.WriteLine($"Thread {_thread.Name} suspended");
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
                Console.WriteLine($"Thread 2 is not created");
            }
        }
    }
}