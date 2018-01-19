using System;
using System.Collections.Generic;
using System.Threading;

namespace threadPoll
{
    class Test
    {
        private Thread _thread;
        WorkWithWord www = new WorkWithWord();

        public List<string> hello()
        {
            List<string> str = new List<string>();
            for (int i = 0; i <= 10; i++)
            {
                str.Add("Hello");
                Console.WriteLine("Hello");
            }
            return str;
        }

        public void CreateThread(string threadName = null)
        {
            if (_thread == null || _thread.ThreadState == ThreadState.Stopped)
            {
                _thread = new Thread(www.WriteToWord);
                _thread.Name = threadName ?? "Test";
                _thread.Start();
            }
        }
    }
}
