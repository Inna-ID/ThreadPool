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

        //DeskWallPaper dwp = new DeskWallPaper();

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
                range.Text = $"The second degree of numbers: {str} \n {chechSystemTime()}";
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
            Console.WriteLine($"Thread {_thread.Name} is running...");
            List<int> array = new List<int>();
            for (int i = 0; i <= 20; i++)
            {
                array.Add(i*i);
                Thread.Sleep(500);
                //Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{new string(' ', 10)}{i*i}");
            }            
            return array;
        }

        private void Hello(Object stateInfo)
        {
            WorkWithWord www = new WorkWithWord(@"D:\Doc1.docx");
            string str = "";
            for (int i = 0; i <= 20; i++)
            {
                Thread.Sleep(100);
                str += "hello\n";
                Console.WriteLine($"{new string(' ', 10)} Hello!");
            }
            //записываем результат в ворд
            www.WriteToWord($"Table of third degree\n{str}");
        }

        public void CreateThread(Object stateInfo /*string threadName = null*/)
        {
            //if thread not created or completed
            if (_thread == null || _thread.ThreadState == ThreadState.Stopped)
            {
                _thread = new Thread(Hello);
                _thread.Name = /*threadName ??*/ "Worker1";
                _thread.Start();
            }           
        }

        //public void ChangePic()
        //{
        //    SystemParametersInfo(SPI_SETDESKWALLPAPER, 0, @"D:\zu6PBjylpow.jpg", SPIF_UPDATEINIFILE | SPIF_SENDWININICHANGE);
        //}

        
        //public void AddThreadToList()
        //{
        //    List<Thread> threadList = new List<Thread>();

        //    if (_thread == null || _thread.ThreadState == ThreadState.Stopped)
        //    {
        //        _thread = new Thread(WriteToWord);
        //        _thread.Name = "Worker1";
        //        _thread.Start();
        //    }

        //    threadList.Add(_thread);
        //    Console.WriteLine("hello from my own thread pool");
        //}

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
                try
                {
                    Console.WriteLine($"Thread {_thread.Name} status: {_thread.ThreadState.ToString()}");
                    if (_thread.ThreadState != ThreadState.Stopped)
                    {
                        Console.WriteLine($"Thread is background: {_thread.IsBackground}");
                        Console.WriteLine($"Thread is thread pool thread: {_thread.IsThreadPoolThread}");
                    }
                    
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
            }
            else
            {
                Console.WriteLine($"Thread is not created");
            }
        }
    }
}