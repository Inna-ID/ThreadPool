using System;
using System.Runtime.InteropServices; //for use Win API

namespace threadPoll
{
    [StructLayout(LayoutKind.Sequential)]
    public class SystemTime
    {
        public ushort year;
        public ushort month;
        public ushort day;
        public ushort hour;
        public ushort minute;
    }

    public class LibWrap
    {
        [DllImport("Kernel32.dll")]
        public static extern void GetSystemTime([In, Out] SystemTime st); //take function GetSystemTime from Kernel32.dll library
    }
}
